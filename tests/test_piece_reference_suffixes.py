from __future__ import annotations

import ast
import io
import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_piece_helpers():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted_assignments = {"PIECE_FILE_RE", "PIECE_REF_TEXT_SUFFIXES"}
    wanted_functions = {
        "_ascii_piece_ref_text",
        "compact_spaces",
        "_normalize_piece_suffix",
        "is_blank_editor_value",
        "coerce_editor_int",
        "piece_reference_piece",
        "piece_reference_metadata",
        "piece_reference_key",
        "piece_reference_key_from_item",
        "piece_reference_parent_key",
        "piece_title_lookup_get",
        "piece_reference_details_from_key",
        "detect_piece_ref_details_from_filename",
        "detect_piece_ref_from_filename",
        "fallback_piece_title_from_filename",
        "build_libelle_affichage",
        "extract_piece_titles_from_deepseek_text",
        "piece_title_lookup_from_split_rows",
        "selected_document_labels_from_rows",
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
        "hashlib": __import__("hashlib"),
        "datetime": datetime,
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

    def test_deepseek_extraction_keeps_sub_piece_references_distinct(self):
        text = """Bordereau :
Pièce 10
Note LGI sur les 7 réserves à la livraison
Pièce 10.1
2022_12_16_SEBIA_SB_00_PRO_ENS_DOC_TN_073A_A Bilan thermique et aéraulique
Pièce 10.2
2024_07_26_Mail LG vers SEBIA_Réserve 2521
"""
        titles = self.ns["extract_piece_titles_from_deepseek_text"](text)
        self.assertEqual(titles["10"], "Note LGI sur les 7 réserves à la livraison")
        self.assertEqual(
            titles["10.1"],
            "2022_12_16_SEBIA_SB_00_PRO_ENS_DOC_TN_073A_A Bilan thermique et aéraulique",
        )
        self.assertEqual(titles["10.2"], "2024_07_26_Mail LG vers SEBIA_Réserve 2521")

    def test_sub_piece_prefers_own_ocr_title_over_parent_title(self):
        lookup = {
            "10": "Note LGI sur les 7 réserves à la livraison",
            "10.1": "2022_12_16_SEBIA_SB_00_PRO_ENS_DOC_TN_073A_A Bilan thermique et aéraulique",
        }
        self.assertEqual(
            self.ns["piece_title_lookup_get"](lookup, 10, "1"),
            "2022_12_16_SEBIA_SB_00_PRO_ENS_DOC_TN_073A_A Bilan thermique et aéraulique",
        )

    def test_sub_piece_uses_filename_before_parent_fallback(self):
        lookup = {"10": "Note LGI sur les 7 réserves à la livraison"}
        filename_title = self.ns["fallback_piece_title_from_filename"](
            "Pièce n°10.1 - 2022_12_16_SEBIA_SB_00_PRO_ENS_DOC_TN_073A_A Bilan thermique et aéraulique.pdf",
            10,
            "1",
            "dot",
        )
        self.assertEqual(
            filename_title,
            "2022_12_16_SEBIA_SB_00_PRO_ENS_DOC_TN_073A_A Bilan thermique et aéraulique",
        )
        self.assertEqual(
            self.ns["piece_title_lookup_get"](lookup, 10, "1", "dot", allow_parent_fallback=False),
            "",
        )

    def test_parent_fallback_is_last_resort_for_sub_piece(self):
        lookup = {"10": "Note LGI sur les 7 réserves à la livraison"}
        filename_title = self.ns["fallback_piece_title_from_filename"]("Pièce n°10.1.pdf", 10, "1", "dot")
        self.assertEqual(filename_title, "")
        self.assertEqual(
            self.ns["piece_title_lookup_get"](lookup, 10, "1", "dot"),
            "Note LGI sur les 7 réserves à la livraison",
        )

    def test_filename_fallback_preserves_annexes_and_simple_pieces(self):
        self.assertEqual(
            self.ns["fallback_piece_title_from_filename"](
                "Piece 1 Annexe 1.1_Programme_version modifiée - signés.pdf",
                1,
                "1.1",
                "annexe",
            ),
            "Programme_version modifiée - signés",
        )
        self.assertEqual(
            self.ns["fallback_piece_title_from_filename"](
                "Pièce n°2 - PV de livraison unité expedition du 25 juin 2024.pdf",
                2,
                "",
            ),
            "PV de livraison unité expedition du 25 juin 2024",
        )

    def test_selected_document_labels_are_distinct_and_manual_values_win(self):
        rows = [
            {
                "fichier_source": "Piece 10 Note LGI.pdf",
                "libelle_retenu": "PIECE n°10 Note LGI sur les 7 réserves à la livraison",
            },
            {
                "fichier_source": "Pièce n°10.1 - Bilan.pdf",
                "libelle_retenu": "PIECE n°10.1 Libellé corrigé manuel",
            },
            {
                "fichier_source": "Pièce n°10.2 - Mail.pdf",
                "libelle_retenu": "PIECE n°10.2 2024_07_26_Mail LG vers SEBIA_Réserve 2521",
            },
        ]
        labels = self.ns["selected_document_labels_from_rows"](rows)
        self.assertEqual(labels["Piece 10 Note LGI.pdf"], "PIECE n°10 Note LGI sur les 7 réserves à la livraison")
        self.assertEqual(labels["Pièce n°10.1 - Bilan.pdf"], "PIECE n°10.1 Libellé corrigé manuel")
        self.assertEqual(labels["Pièce n°10.2 - Mail.pdf"], "PIECE n°10.2 2024_07_26_Mail LG vers SEBIA_Réserve 2521")

    def test_copy_persists_distinct_libelle_final_in_transmission_event(self):
        tmp_root = Path(tempfile.mkdtemp(prefix="llm_assistant_piece_labels_", dir=r"C:\CodexWorkspace"))
        try:
            class Uploaded(io.BytesIO):
                def __init__(self, name: str):
                    super().__init__(b"%PDF-1.4\n")
                    self.name = name

            files = [
                Uploaded("Piece 10 Note LGI.pdf"),
                Uploaded("Pièce n°10.1 - Bilan.pdf"),
                Uploaded("Pièce n°10.2 - Mail.pdf"),
            ]
            labels = {
                "Piece 10 Note LGI.pdf": "PIECE n°10 Note LGI sur les 7 réserves à la livraison",
                "Pièce n°10.1 - Bilan.pdf": "PIECE n°10.1 2022_12_16_SEBIA Bilan thermique",
                "Pièce n°10.2 - Mail.pdf": "PIECE n°10.2 2024_07_26_Mail LG vers SEBIA",
            }
            refs = {
                "Piece 10 Note LGI.pdf": {"numero_piece": 10, "sous_piece": "", "piece_ref_style": ""},
                "Pièce n°10.1 - Bilan.pdf": {"numero_piece": 10, "sous_piece": "1", "piece_ref_style": "dot"},
                "Pièce n°10.2 - Mail.pdf": {"numero_piece": 10, "sous_piece": "2", "piece_ref_style": "dot"},
            }
            event = self.ns["copy_ingestion_uploaded_originals"](
                str(tmp_root),
                "2026-A60",
                {"code_partie": 2, "nom": "LEON GROSSE IMMOBILIER", "folder_rel": "02_Partie"},
                files,
                {"date_transmission_expert": "2026-06-15", "type_transmission": "lettre"},
                {},
                {name: "piece" for name in labels},
                labels,
                {},
                refs,
            )
            copied = {item["name"]: item for item in event["copied"]}
            self.assertEqual(copied["Piece 10 Note LGI.pdf"]["libelle_final"], labels["Piece 10 Note LGI.pdf"])
            self.assertEqual(copied["Pièce n°10.1 - Bilan.pdf"]["libelle_final"], labels["Pièce n°10.1 - Bilan.pdf"])
            self.assertEqual(copied["Pièce n°10.2 - Mail.pdf"]["libelle_affichage"], labels["Pièce n°10.2 - Mail.pdf"])
            self.assertEqual(copied["Pièce n°10.1 - Bilan.pdf"]["reference_piece"], "10.1")
            self.assertEqual(copied["Pièce n°10.2 - Mail.pdf"]["reference_piece"], "10.2")
        finally:
            shutil.rmtree(tmp_root, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
