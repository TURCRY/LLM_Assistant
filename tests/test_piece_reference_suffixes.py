from __future__ import annotations

import ast
import io
import shutil
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_piece_helpers():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted_assignments = {
        "DOCUMENT_TITLE_EXTENSIONS",
        "PIECE_FILE_RE",
        "PIECE_REF_TEXT_SUFFIXES",
        "INGESTION_DOCUMENT_SUBJECT_ORDER",
    }
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
        "strip_document_extension",
        "clean_document_filename_label",
        "normalize_document_label",
        "label_contains_piece_reference",
        "piece_label_with_reference",
        "document_display_label",
        "parse_date_for_sort",
        "valid_expert_doc_id",
        "expert_doc_id_sort_key",
        "etat2_document_sort_key",
        "document_registry_fingerprint",
        "document_state_date",
        "is_real_non_piece_document",
        "doc_dedupe_key",
        "document_priority_score",
        "dedupe_documents_prefer_validated_rows",
        "_document_ingestion_sort_text",
        "_document_ingestion_subject_from_filename",
        "_document_ingestion_display_subject",
        "_document_ingestion_display_reference",
        "document_ingestion_business_label_from_filename",
        "_document_ingestion_suffix_sort_key",
        "document_ingestion_sort_key",
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
        "date": date,
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

    def test_document_display_label_keeps_business_piece_reference(self):
        doc = {
            "numero_piece": 1,
            "sous_piece": "",
            "libelle_final": "Introduction - Pièce 1 - SEBIA Contrat Promotion Immobilière",
        }
        self.assertEqual(
            self.ns["document_display_label"](doc),
            "Introduction - Pièce 1 - SEBIA Contrat Promotion Immobilière",
        )

    def test_document_display_label_keeps_business_suffix_reference(self):
        doc = {
            "numero_piece": 5,
            "sous_piece": "a",
            "libelle_final": "Réserve 2521 - Pièce 5.a - Re SEBIA",
        }
        self.assertEqual(
            self.ns["document_display_label"](doc),
            "Réserve 2521 - Pièce 5.a - Re SEBIA",
        )

    def test_document_display_label_matches_suffix_case_without_doubling(self):
        doc = {
            "numero_piece": 5,
            "sous_piece": "A",
            "libelle_final": "réserve 2521 - pièce 5.a - Re SEBIA",
        }
        self.assertEqual(
            self.ns["document_display_label"](doc),
            "réserve 2521 - pièce 5.a - Re SEBIA",
        )

    def test_document_display_label_does_not_match_partial_piece_number(self):
        doc = {
            "numero_piece": 1,
            "sous_piece": "",
            "libelle_final": "Réserve 4908 - Pièce 10 - SEBIA LIS2",
        }
        self.assertEqual(
            self.ns["document_display_label"](doc),
            "Piece 1 - Réserve 4908 - Pièce 10 - SEBIA LIS2",
        )

    def test_document_display_label_keeps_historical_prefixing(self):
        doc = {"numero_piece": 1, "sous_piece": "", "libelle_final": "Plan masse"}
        self.assertEqual(self.ns["document_display_label"](doc), "Piece 1 - Plan masse")

    def test_document_display_label_keeps_corrected_business_label_priority(self):
        doc = {
            "numero_piece": 1,
            "sous_piece": "",
            "libelle_corrige": "Introduction - Pièce 1 - Libellé corrigé",
            "libelle_final": "Plan masse",
        }
        self.assertEqual(
            self.ns["document_display_label"](doc),
            "Introduction - Pièce 1 - Libellé corrigé",
        )

    def test_document_state_labels_keep_business_reference(self):
        docs = [
            {
                "numero_piece": 1,
                "sous_piece": "",
                "libelle_affichage": "Introduction - Pièce 1 - SEBIA Contrat Promotion Immobilière",
            },
            {
                "numero_piece": 5,
                "sous_piece": "A",
                "libelle_affichage": "Réserve 2521 - Pièce 5.a - Re SEBIA",
            },
            {
                "numero_piece": 10,
                "sous_piece": "",
                "libelle_affichage": "Réserve 4908 - Pièce 10 - SEBIA LIS2",
            },
        ]
        etat2 = [{"libelle_retenu": self.ns["document_display_label"](doc)} for doc in docs]
        etat3 = [{"libelle": self.ns["document_display_label"](doc)} for doc in docs]
        etat4 = [{"objet": self.ns["document_display_label"](doc)} for doc in docs]
        self.assertEqual(etat2[0]["libelle_retenu"], "Introduction - Pièce 1 - SEBIA Contrat Promotion Immobilière")
        self.assertEqual(etat3[1]["libelle"], "Réserve 2521 - Pièce 5.a - Re SEBIA")
        self.assertEqual(etat4[2]["objet"], "Réserve 4908 - Pièce 10 - SEBIA LIS2")

    def test_dedupe_keeps_same_piece_reference_with_distinct_expert_ids(self):
        docs = [
            {
                "expert_doc_id": "01-0001",
                "date_transmission_expert": "2026-07-03",
                "deposant": "SEBIA",
                "numero_piece": 1,
                "sous_piece": "",
                "fichier_source": "Piece 1 - Introduction.pdf",
                "libelle_retenu": "Introduction - Pièce 1 - Contrat",
            },
            {
                "expert_doc_id": "01-0016",
                "date_transmission_expert": "2026-07-03",
                "deposant": "SEBIA",
                "numero_piece": 1,
                "sous_piece": "",
                "fichier_source": "Piece 1 - Reserve 2521.pdf",
                "libelle_retenu": "Réserve 2521 - Pièce 1 - Stockage",
            },
        ]
        selected, diag = self.ns["dedupe_documents_prefer_validated_rows"](docs)
        self.assertEqual([doc["expert_doc_id"] for doc in selected], ["01-0001", "01-0016"])
        self.assertEqual(diag["nombre_lignes_apres_dedoublonnage"], 2)
        self.assertEqual(diag["lignes_supprimees"], [])

    def test_dedupe_collapses_duplicate_expert_id(self):
        docs = [
            {
                "expert_doc_id": "01-0001",
                "date_transmission_expert": "2026-07-03",
                "deposant": "SEBIA",
                "numero_piece": 1,
                "sous_piece": "",
                "fichier_source": "Piece 1 - Introduction.pdf",
                "libelle_retenu": "Introduction - Pièce 1 - Contrat",
            },
            {
                "expert_doc_id": "01-0001",
                "date_transmission_expert": "2026-07-03",
                "deposant": "SEBIA",
                "numero_piece": 1,
                "sous_piece": "",
                "fichier_source": "Piece 1 - Introduction copie.pdf",
                "libelle_retenu": "Introduction - Pièce 1 - Contrat",
            },
        ]
        selected, diag = self.ns["dedupe_documents_prefer_validated_rows"](docs)
        self.assertEqual(len(selected), 1)
        self.assertEqual(selected[0]["expert_doc_id"], "01-0001")
        self.assertEqual(len(diag["lignes_supprimees"]), 1)

    def test_dedupe_without_expert_id_keeps_historical_piece_reference_fallback(self):
        docs = [
            {
                "date_transmission_expert": "2026-07-03",
                "deposant": "SEBIA",
                "numero_piece": 1,
                "sous_piece": "",
                "fichier_source": "Piece 1 - Introduction.pdf",
                "libelle_retenu": "Introduction - Pièce 1 - Contrat",
            },
            {
                "date_transmission_expert": "2026-07-03",
                "deposant": "SEBIA",
                "numero_piece": 1,
                "sous_piece": "",
                "fichier_source": "Piece 1 - Reserve 2521.pdf",
                "libelle_retenu": "Réserve 2521 - Pièce 1 - Stockage",
            },
        ]
        selected, diag = self.ns["dedupe_documents_prefer_validated_rows"](docs)
        self.assertEqual(len(selected), 1)
        self.assertEqual(diag["nombre_lignes_apres_dedoublonnage"], 1)
        self.assertEqual(len(diag["lignes_supprimees"]), 1)

    def test_argumentation_transmission_distinct_expert_ids_are_not_deduped(self):
        transmissions = Path(r"C:\Affaires\2026-A60\AA_Expert_Admin\_Logs\transmissions.jsonl")
        if not transmissions.exists():
            self.skipTest("Journal transmissions 2026-A60 absent sur cette machine")
        import json

        record = None
        for line in transmissions.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            current = json.loads(line)
            copied = [item for item in (current.get("copied") or []) if isinstance(item, dict)]
            if any(str(item.get("expert_doc_id") or "").startswith("01-") for item in copied):
                record = current
                break
        if not record:
            self.skipTest("Transmission partie 01 absente sur cette machine")

        docs = []
        for item in record.get("copied") or []:
            if not isinstance(item, dict):
                continue
            docs.append({
                "expert_doc_id": item.get("expert_doc_id") or "",
                "date_transmission_expert": record.get("date_transmission_expert") or "",
                "deposant": (record.get("party") or {}).get("nom") or "",
                "numero_piece": item.get("numero_piece"),
                "sous_piece": item.get("sous_piece") or "",
                "fichier_source": item.get("source_name") or item.get("name") or "",
                "libelle_retenu": item.get("libelle_affichage") or item.get("libelle_final") or "",
                "libelle_affichage": item.get("libelle_affichage") or "",
                "libelle_final": item.get("libelle_final") or "",
                "document_role": item.get("document_role") or "",
                "type_document": item.get("type_document") or "",
            })

        selected, diag = self.ns["dedupe_documents_prefer_validated_rows"](docs)
        selected_ids = {doc.get("expert_doc_id") for doc in selected}
        self.assertEqual(len(docs), 109)
        self.assertEqual(len(selected), 109)
        self.assertEqual(
            sum(1 for doc in selected if str(doc.get("expert_doc_id") or "").startswith("01-") and doc.get("numero_piece")),
            108,
        )
        self.assertFalse([
            row for row in diag["lignes_supprimees"]
            if str(row.get("expert_doc_id") or "").startswith("01-")
        ])
        for expert_doc_id in {
            "01-0001", "01-0016", "01-0032", "01-0043",
            "01-0066", "01-0073", "01-0082", "01-0097",
            "01-0008", "01-0009", "01-0010", "01-0011",
            "01-0021", "01-0050", "01-0051",
        }:
            self.assertIn(expert_doc_id, selected_ids)

    def test_etat2_sort_orders_expert_doc_ids_naturally_within_same_date(self):
        docs = [
            {"expert_doc_id": "01-0002", "date_transmission_expert": "2026-07-03", "deposant": "SEBIA"},
            {"expert_doc_id": "01-0010", "date_transmission_expert": "2026-07-03", "deposant": "SEBIA"},
            {"expert_doc_id": "01-0001", "date_transmission_expert": "2026-07-03", "deposant": "SEBIA"},
        ]
        ordered = sorted(docs, key=self.ns["etat2_document_sort_key"])
        self.assertEqual([doc["expert_doc_id"] for doc in ordered], ["01-0001", "01-0002", "01-0010"])

    def test_etat2_sort_keeps_current_chronological_direction(self):
        docs = [
            {"expert_doc_id": "01-0001", "date_transmission_expert": "2026-07-02", "deposant": "SEBIA"},
            {"expert_doc_id": "01-0002", "date_transmission_expert": "2026-07-03", "deposant": "SEBIA"},
        ]
        ordered = sorted(docs, key=self.ns["etat2_document_sort_key"])
        self.assertEqual([doc["date_transmission_expert"] for doc in ordered], ["2026-07-03", "2026-07-02"])

    def test_etat2_sort_without_expert_doc_id_uses_historical_fallback(self):
        docs = [
            {
                "date_transmission_expert": "2026-07-03",
                "deposant": "SEBIA",
                "numero_piece": 10,
                "fichier_source": "Piece 10.pdf",
            },
            {
                "date_transmission_expert": "2026-07-03",
                "deposant": "SEBIA",
                "numero_piece": 2,
                "fichier_source": "Piece 2.pdf",
            },
        ]
        ordered = sorted(docs, key=self.ns["etat2_document_sort_key"])
        self.assertEqual([doc["fichier_source"] for doc in ordered], ["Piece 2.pdf", "Piece 10.pdf"])

    def test_argumentation_transmission_etat2_orders_by_expert_doc_id(self):
        transmissions = Path(r"C:\Affaires\2026-A60\AA_Expert_Admin\_Logs\transmissions.jsonl")
        if not transmissions.exists():
            self.skipTest("Journal transmissions 2026-A60 absent sur cette machine")
        import json

        record = None
        for line in transmissions.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            current = json.loads(line)
            copied = [item for item in (current.get("copied") or []) if isinstance(item, dict)]
            if any(str(item.get("expert_doc_id") or "").startswith("01-") for item in copied):
                record = current
                break
        if not record:
            self.skipTest("Transmission partie 01 absente sur cette machine")

        docs = [
            {
                "expert_doc_id": item.get("expert_doc_id") or "",
                "date_transmission_expert": record.get("date_transmission_expert") or "",
                "deposant": (record.get("party") or {}).get("nom") or "",
                "numero_piece": item.get("numero_piece"),
                "sous_piece": item.get("sous_piece") or "",
                "fichier_source": item.get("source_name") or item.get("name") or "",
            }
            for item in (record.get("copied") or [])
            if isinstance(item, dict)
        ]
        ordered = sorted(docs, key=self.ns["etat2_document_sort_key"])
        self.assertEqual(len(ordered), 109)
        self.assertEqual([doc["expert_doc_id"] for doc in ordered], [f"01-{idx:04d}" for idx in range(1, 110)])

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

    def test_document_ingestion_sort_subject_order(self):
        names = [
            "Piece 1 - Reserve 4908 - Cahier des charges.pdf",
            "Piece 1 - Réserve 2521 - Stockage matières.pdf",
            "Piece 1 - Introduction - Contrat.pdf",
        ]
        ordered = sorted(names, key=self.ns["document_ingestion_sort_key"])
        self.assertEqual(ordered, [
            "Piece 1 - Introduction - Contrat.pdf",
            "Piece 1 - Réserve 2521 - Stockage matières.pdf",
            "Piece 1 - Reserve 4908 - Cahier des charges.pdf",
        ])

    def test_document_ingestion_business_label_from_prepared_name(self):
        cases = [
            (
                "Piece 1 - Introduction - SEBIA Contrat Promotion Immobilière version modifiée - signés.pdf",
                "Introduction - Pièce 1 - SEBIA Contrat Promotion Immobilière version modifiée - signés",
            ),
            (
                "Piece 1 - Reserve 2521 - TR PSEBIA - Stockage matière premières dangereuses.pdf",
                "Réserve 2521 - Pièce 1 - TR PSEBIA - Stockage matière premières dangereuses",
            ),
            (
                "Piece 05.a - Reserve 4908 - SEBIA LIS2 Tableau de suivi des observations.pdf",
                "Réserve 4908 - Pièce 5.a - SEBIA LIS2 Tableau de suivi des observations",
            ),
            (
                "Piece 5 - Introduction - Annexe - 24 08 08 - SEBIA - Suivi Des Ecarts - Ecart PCA.pdf",
                "Introduction - Pièce 5 - Annexe - 24 08 08 - SEBIA - Suivi Des Ecarts - Ecart PCA",
            ),
            (
                "Piece 1 - Réserve 5584 - dimension porte restaurant avant travaux.pdf",
                "Réserve 5584 - Pièce 1 - dimension porte restaurant avant travaux",
            ),
        ]
        for filename, expected in cases:
            with self.subTest(filename=filename):
                self.assertEqual(self.ns["document_ingestion_business_label_from_filename"](filename), expected)

    def test_document_ingestion_business_label_falls_back_for_other_names(self):
        self.assertEqual(self.ns["document_ingestion_business_label_from_filename"]("Piece 1 - Sujet inconnu - Beta.pdf"), "")
        self.assertEqual(self.ns["document_ingestion_business_label_from_filename"]("rapport technique 2026.pdf"), "")

    def test_document_ingestion_sort_piece_numbers_naturally(self):
        names = [
            "Piece 10 - Introduction - Dix.pdf",
            "Piece 2 - Introduction - Deux.pdf",
        ]
        self.assertEqual(
            sorted(names, key=self.ns["document_ingestion_sort_key"]),
            [
                "Piece 2 - Introduction - Deux.pdf",
                "Piece 10 - Introduction - Dix.pdf",
            ],
        )

    def test_document_ingestion_sort_suffixes_after_main_piece(self):
        names = [
            "Piece 5.b - Reserve 2521 - B.pdf",
            "Piece 5 - Reserve 2521 - Main.pdf",
            "Piece 5.a - Reserve 2521 - A.pdf",
        ]
        self.assertEqual(
            sorted(names, key=self.ns["document_ingestion_sort_key"]),
            [
                "Piece 5 - Reserve 2521 - Main.pdf",
                "Piece 5.a - Reserve 2521 - A.pdf",
                "Piece 5.b - Reserve 2521 - B.pdf",
            ],
        )

    def test_document_ingestion_sort_master_before_annexe(self):
        names = [
            "Piece 5 - Introduction - Annexe - Ecart PCA.pdf",
            "Piece 5 - Introduction - SEBIA - Suivi des écarts.pdf",
        ]
        self.assertEqual(
            sorted(names, key=self.ns["document_ingestion_sort_key"]),
            [
                "Piece 5 - Introduction - SEBIA - Suivi des écarts.pdf",
                "Piece 5 - Introduction - Annexe - Ecart PCA.pdf",
            ],
        )

    def test_document_ingestion_sort_unknown_names_are_stable(self):
        names = [
            "zeta.pdf",
            "alpha.pdf",
            "Piece 1 - Sujet inconnu - Beta.pdf",
        ]
        self.assertEqual(
            sorted(names, key=self.ns["document_ingestion_sort_key"]),
            [
                "Piece 1 - Sujet inconnu - Beta.pdf",
                "alpha.pdf",
                "zeta.pdf",
            ],
        )

    def test_document_ingestion_sort_representative_sample(self):
        names = [
            "Piece 5 - Introduction - Annexe - Ecart PCA.pdf",
            "Piece 1 - Reserve 2521 - TR.pdf",
            "Piece 5.b - Reserve 2521 - B.pdf",
            "Piece 5 - Introduction - Mail maître.pdf",
            "Piece 5.a - Reserve 2521 - A.pdf",
            "Piece 1 - Introduction - Contrat.pdf",
        ]
        self.assertEqual(
            sorted(names, key=self.ns["document_ingestion_sort_key"]),
            [
                "Piece 1 - Introduction - Contrat.pdf",
                "Piece 5 - Introduction - Mail maître.pdf",
                "Piece 5 - Introduction - Annexe - Ecart PCA.pdf",
                "Piece 1 - Reserve 2521 - TR.pdf",
                "Piece 5.a - Reserve 2521 - A.pdf",
                "Piece 5.b - Reserve 2521 - B.pdf",
            ],
        )

    def test_argumentation_cohort_108_sort_simulation(self):
        manifest = Path(r"C:\Affaires\2026-A60\01_Partie_01_SEBIA\_Depot_Argumentation_SEBIA\prepare_argumentation_cohort.json")
        if not manifest.exists():
            self.skipTest("Cohorte Argumentation SEBIA absente sur cette machine")
        import json

        rows = json.loads(manifest.read_text(encoding="utf-8"))
        names = [row["target_name"] for row in rows if row.get("target_name")]
        ordered = sorted(names, key=self.ns["document_ingestion_sort_key"])

        def subject(name: str) -> str:
            parts = Path(name).stem.split(" - ")
            return parts[1] if len(parts) >= 3 else ""

        transitions = {}
        last_subject = None
        for idx, name in enumerate(ordered, start=1):
            current = subject(name)
            if current != last_subject:
                transitions[current] = f"01-{idx:04d}"
                last_subject = current

        self.assertEqual(len(ordered), 108)
        self.assertEqual(transitions["Introduction"], "01-0001")
        self.assertEqual(transitions["Reserve 2521"], "01-0016")
        self.assertEqual(transitions["Reserve 4226"], "01-0032")
        self.assertEqual(transitions["Reserve 4908"], "01-0043")
        self.assertEqual(transitions["Reserve 5431"], "01-0066")
        self.assertEqual(transitions["Reserve 5584"], "01-0073")
        self.assertEqual(transitions["Reserve 5951"], "01-0082")
        self.assertEqual(transitions["Reserve 5959"], "01-0097")
        self.assertEqual(f"01-{len(ordered):04d}", "01-0108")

    def test_argumentation_cohort_108_business_labels(self):
        manifest = Path(r"C:\Affaires\2026-A60\01_Partie_01_SEBIA\_Depot_Argumentation_SEBIA\prepare_argumentation_cohort.json")
        if not manifest.exists():
            self.skipTest("Cohorte Argumentation SEBIA absente sur cette machine")
        import json

        rows = json.loads(manifest.read_text(encoding="utf-8"))
        names = [row["target_name"] for row in rows if row.get("target_name")]
        labels = {
            name: self.ns["document_ingestion_business_label_from_filename"](name)
            for name in names
        }
        self.assertEqual(len(labels), 108)
        self.assertFalse([name for name, label in labels.items() if not label])
        self.assertEqual(
            labels["Piece 1 - Introduction - SEBIA Contrat Promotion Immobilière version modifiée - signés.pdf"],
            "Introduction - Pièce 1 - SEBIA Contrat Promotion Immobilière version modifiée - signés",
        )
        self.assertEqual(
            labels["Piece 1 - Reserve 2521 - TR PSEBIA - Stockage matière premières dangereuses.pdf"],
            "Réserve 2521 - Pièce 1 - TR PSEBIA - Stockage matière premières dangereuses",
        )
        self.assertEqual(
            labels["Piece 05.a - Reserve 4908 - SEBIA LIS2 Tableau de suivi des observations 27 02 2024.pdf"],
            "Réserve 4908 - Pièce 5.a - SEBIA LIS2 Tableau de suivi des observations 27 02 2024",
        )
        self.assertEqual(
            labels["Piece 5 - Introduction - Annexe - 24 08 08 - SEBIA - Suivi Des Ecarts - Ecart PCA.pdf"],
            "Introduction - Pièce 5 - Annexe - 24 08 08 - SEBIA - Suivi Des Ecarts - Ecart PCA",
        )
        self.assertEqual(
            labels["Piece 1 - Reserve 5584 - dimension porte restaurant avant travaux.pdf"],
            "Réserve 5584 - Pièce 1 - dimension porte restaurant avant travaux",
        )
        self.assertFalse([label for label in labels.values() if label.startswith("Piece ")])


if __name__ == "__main__":
    unittest.main()
