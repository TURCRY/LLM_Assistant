from __future__ import annotations

import ast
import unittest
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_piece_helpers():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted_assignments = {"PIECE_FILE_RE", "PIECE_REF_TEXT_SUFFIXES"}
    wanted_functions = {
        "_ascii_piece_ref_text",
        "compact_spaces",
        "_normalize_piece_suffix",
        "piece_reference_piece",
        "piece_reference_metadata",
        "detect_piece_ref_details_from_filename",
        "detect_piece_ref_from_filename",
        "build_libelle_affichage",
        "document_registry_fingerprint",
        "copy_ingestion_uploaded_originals",
    }
    body = []
    for node in tree.body:
        if isinstance(node, ast.Assign):
            names = {target.id for target in node.targets if isinstance(target, ast.Name)}
            if names & wanted_assignments:
                body.append(node)
        elif isinstance(node, ast.FunctionDef) and node.name in wanted_functions:
            body.append(node)
    module = ast.Module(body=body, type_ignores=[])
    ast.fix_missing_locations(module)

    ns = {
        "Path": Path,
        "re": __import__("re"),
        "unicodedata": __import__("unicodedata"),
        "pj": lambda *parts: str(Path(str(parts[0])).joinpath(*[str(p) for p in parts[1:] if str(p)])),
        "normalize_document_role": lambda role: str(role or ""),
        "document_type_from_role": lambda role: str(role or ""),
        "file_role_payload": lambda role, source="ui_ingestion": {
            "document_role": str(role or ""),
            "type_document": str(role or ""),
            "qualification_source": source if role else "",
        },
        "source_roots_for_folder": lambda *args, **kwargs: {},
        "ensure_source_roots": lambda roots: {},
        "file_page_count_record": lambda path: {"page_count": 1, "page_count_source": "test"},
        "file_stat_record": lambda path: {"present": Path(path).exists()},
        "insert_transmission_documents_sqlite": lambda *args, **kwargs: {
            "sqlite_path": "",
            "inserted": [],
            "existing": [],
        },
        "affaire_sqlite_path_from_root": lambda root, cfg=None: Path(root) / "Documents.db",
        "write_ingestion_log": lambda root, event, cfg=None: str(Path(root) / "log.json"),
        "append_transmission_record": lambda root, event, cfg=None: str(Path(root) / "journal.json"),
        "user_machine_label": lambda: "test",
        "make_transmission_id": lambda aff_id: f"{aff_id}-T",
        "JURIDICTION_FOLDER_REL": "_Juridiction",
    }
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns


class PieceReferenceSuffixTests(unittest.TestCase):
    def setUp(self):
        self.ns = _load_piece_helpers()

    def test_detects_real_annex_names_without_float_conversion(self):
        cases = [
            ("Piece  1 Annexe  1.1_Programme_version modifiee - signes.pdf", 1, "1.1", "1 annexe 1.1"),
            ("Piece  1 Annexe  2.10_PC_Pieces ecrites et graphiques - signes.pdf", 1, "2.10", "1 annexe 2.10"),
            ("Piece  1 Annexe 2.11_PC_Recepisse depot - signes.pdf", 1, "2.11", "1 annexe 2.11"),
            ("Piece  1 Annexe 7.4_Attestation CCRD - signes.pdf", 1, "7.4", "1 annexe 7.4"),
        ]
        for filename, numero, sous_piece, reference in cases:
            with self.subTest(filename=filename):
                self.assertEqual(self.ns["detect_piece_ref_from_filename"](filename), (numero, sous_piece))
                details = self.ns["detect_piece_ref_details_from_filename"](filename)
                self.assertEqual(details["sous_piece"], sous_piece)
                self.assertIsInstance(details["sous_piece"], str)
                self.assertEqual(details["reference_piece"], reference)

    def test_detects_alpha_text_and_hyphen_suffixes(self):
        cases = [
            ("PIECE 1.A", (1, "A"), "1.A"),
            ("PIECE 1.B", (1, "B"), "1.B"),
            ("PIECE 3 bis", (3, "bis"), "3 bis"),
            ("PIECE 3 ter", (3, "ter"), "3 ter"),
            ("PIECE 4-1", (4, "1"), "4-1"),
            ("  piece   1   annexe   2.10   suite.pdf", (1, "2.10"), "1 annexe 2.10"),
            ("Piece 1 Annexe 2.10_Pieces ecrites accentuees.pdf", (1, "2.10"), "1 annexe 2.10"),
        ]
        for filename, expected_pair, expected_reference in cases:
            with self.subTest(filename=filename):
                self.assertEqual(self.ns["detect_piece_ref_from_filename"](filename), expected_pair)
                self.assertEqual(
                    self.ns["detect_piece_ref_details_from_filename"](filename)["reference_piece"],
                    expected_reference,
                )

    def test_no_false_positive_without_piece_reference(self):
        self.assertEqual(self.ns["detect_piece_ref_from_filename"]("rapport technique 2026.pdf"), (None, ""))

    def test_simple_piece_and_display_labels_are_backward_compatible(self):
        self.assertEqual(self.ns["detect_piece_ref_from_filename"]("Piece 1.pdf"), (1, ""))
        self.assertEqual(self.ns["piece_reference_piece"](1, ""), "1")
        self.assertEqual(self.ns["build_libelle_affichage"](1, "", "Plan masse"), "Piece 1 - Plan masse")
        self.assertEqual(self.ns["build_libelle_affichage"](1, "2.10", "Plans"), "Piece 1 annexe 2.10 - Plans")
        self.assertEqual(self.ns["build_libelle_affichage"](1, "A", "Plans"), "Piece 1.A - Plans")
        self.assertEqual(self.ns["build_libelle_affichage"](3, "bis", "Notice"), "Piece 3 bis - Notice")

    def test_synthesis_block_exposes_and_transmits_sous_piece(self):
        text = APP_PATH.read_text(encoding="utf-8-sig")
        self.assertIn('"sous_piece": sous_piece', text)
        self.assertIn('"reference_piece": reference_piece', text)
        self.assertIn('selected_document_piece_refs = {}', text)
        self.assertIn('"numero_piece": st.column_config.TextColumn("numero_piece")', text)
        self.assertIn('"sous_piece": st.column_config.TextColumn("sous_piece")', text)
        self.assertIn('"reference_piece": st.column_config.TextColumn("reference_piece", disabled=True)', text)

    def test_copy_contract_accepts_and_persists_piece_refs(self):
        text = APP_PATH.read_text(encoding="utf-8-sig")
        tree = ast.parse(text)
        copy_func = next(
            node for node in tree.body
            if isinstance(node, ast.FunctionDef) and node.name == "copy_ingestion_uploaded_originals"
        )
        self.assertIn("document_piece_refs", [arg.arg for arg in copy_func.args.args])
        self.assertIn("piece_ref_meta = {", text)
        self.assertGreaterEqual(text.count("**piece_ref_meta,"), 2)
        self.assertIn('"numero_piece", "sous_piece", "reference_piece", "piece_ref_style"', text)

        refs = {
            "Piece  1 Annexe  2.10_Plans.pdf": {
                "numero_piece": 1,
                "sous_piece": "2.10",
                "reference_piece": "1 annexe 2.10",
                "piece_ref_style": "annexe",
            }
        }
        self.assertEqual(refs["Piece  1 Annexe  2.10_Plans.pdf"]["sous_piece"], "2.10")

    def test_fingerprints_distinguish_annex_suffixes(self):
        fp = self.ns["document_registry_fingerprint"]
        self.assertNotEqual(
            fp({"numero_piece": 1, "sous_piece": "2.10", "fichier_source": "a.pdf"}),
            fp({"numero_piece": 1, "sous_piece": "2.11", "fichier_source": "a.pdf"}),
        )


if __name__ == "__main__":
    unittest.main()
