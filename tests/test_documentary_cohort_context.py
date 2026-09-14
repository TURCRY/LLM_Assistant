from __future__ import annotations

import ast
import hashlib
import json
import re
import tempfile
import unittest
import unicodedata
from datetime import date
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_helpers():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted_assignments = {
        "WINDOWS_FORBIDDEN",
        "PIECE_REF_TEXT_SUFFIXES",
    }
    wanted_functions = {
        "sanitize_filename",
        "is_blank_editor_value",
        "safe_text",
        "party_code",
        "compact_spaces",
        "coerce_editor_int",
        "_ascii_piece_ref_text",
        "_normalize_piece_suffix",
        "piece_reference_piece",
        "piece_reference_key",
        "piece_reference_parent_key",
        "piece_title_lookup_get",
        "valid_expert_doc_id",
        "expert_doc_id_sort_key",
        "documentary_cohort_context",
        "documentary_cohort_context_id",
        "documentary_cohort_widget_suffix",
        "party_from_cohort_context",
        "cohort_dependent_session_keys",
        "reset_cohort_dependent_session_state",
        "sync_documentary_cohort_session",
        "ingestion_context_from_values",
        "validate_cohort_ingestion_context",
        "expected_split_child_filenames",
        "filter_split_child_paths_for_current_context",
        "document_registry_override_key",
        "documentary_maintenance_row_key",
        "documentary_maintenance_selection_signature",
        "documentary_deletion_confirmation_text",
        "documentary_deletion_preflight",
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
        "hashlib": hashlib,
        "json": json,
        "Path": Path,
        "re": re,
        "unicodedata": unicodedata,
    }
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns


class DocumentaryCohortContextTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = _load_helpers()

    def _party(self, code=25, name="Gérard TORDJMAN"):
        return {
            "code_partie": code,
            "nom": name,
            "folder_rel": f"{code:02d}_Partie_{code:02d}_Gerard_TORDJMAN",
        }

    def _context(self, code=25, when=date(2025, 10, 24), author="Maître Alberta SMALL", files=None):
        return self.ns["documentary_cohort_context"](
            "2025-J48",
            self._party(code),
            when,
            author,
            files or ["dire.pdf", "multi.pdf"],
        )

    def test_cohort_context_propagates_to_ingestion(self):
        ctx = self._context()
        ingestion = self.ns["ingestion_context_from_values"]("2025-J48", self._party(25), date(2025, 10, 24), "Maître Alberta SMALL")
        ok, divergences = self.ns["validate_cohort_ingestion_context"](ctx, ingestion)
        self.assertTrue(ok)
        self.assertEqual([], divergences)
        self.assertEqual("25", ingestion["code_partie"])

    def test_switching_to_cohort_b_clears_cohort_a_state(self):
        state = {
            "pdf_current_cohort_signature_2025-J48": self._context(code=1)["context_id"],
            "split_rows": [{"numero_piece": 1}],
            "ingestion_pieces": ["PIECE_01_Piece_1.pdf"],
            "piece_title_suggestions": {1: "Piece 1"},
        }
        result = self.ns["sync_documentary_cohort_session"](state, "2025-J48", self._context(code=2))
        self.assertTrue(result["changed"])
        self.assertNotIn("split_rows", state)
        self.assertNotIn("ingestion_pieces", state)
        self.assertNotIn("piece_title_suggestions", state)

    def test_changing_party_updates_downstream_code(self):
        ctx = self._context(code=7)
        ingestion = self.ns["ingestion_context_from_values"]("2025-J48", self._party(7), date(2025, 10, 24), "Maître Alberta SMALL")
        self.assertEqual("07", ingestion["code_partie"])
        self.assertTrue(self.ns["validate_cohort_ingestion_context"](ctx, ingestion)[0])

    def test_changing_date_updates_ingestion_date(self):
        ingestion = self.ns["ingestion_context_from_values"]("2025-J48", self._party(25), date(2026, 1, 2), "Maître Alberta SMALL")
        self.assertEqual("2026-01-02", ingestion["date_transmission"])

    def test_changing_author_updates_ingestion_author(self):
        ingestion = self.ns["ingestion_context_from_values"]("2025-J48", self._party(25), date(2025, 10, 24), "Autre conseil")
        self.assertEqual("Autre conseil", ingestion["auteur_transmission"])

    def test_no_double_source_party_and_target_keys_survive_reset(self):
        state = {
            "ingestion_party": 0,
            "classement_originaux_partie_cible_2025-J48": 1,
            "split_rows": [{"numero_piece": 1}],
        }
        removed = self.ns["reset_cohort_dependent_session_state"](state, "2025-J48")
        self.assertIn("ingestion_party", removed)
        self.assertNotIn("split_rows", state)

    def test_artificial_divergence_blocks_validation(self):
        ctx = self._context(code=25)
        ingestion = self.ns["ingestion_context_from_values"]("2025-J48", self._party(26), date(2025, 10, 24), "Maître Alberta SMALL")
        ok, divergences = self.ns["validate_cohort_ingestion_context"](ctx, ingestion)
        self.assertFalse(ok)
        self.assertEqual(["code_partie"], [row["champ"] for row in divergences])

    def test_streamlit_reruns_keep_same_context_stable(self):
        state = {}
        ctx = self._context()
        first = self.ns["sync_documentary_cohort_session"](state, "2025-J48", ctx)
        second = self.ns["sync_documentary_cohort_session"](state, "2025-J48", ctx)
        self.assertFalse(first["changed"])
        self.assertFalse(second["changed"])
        self.assertEqual(ctx["context_id"], state["pdf_current_cohort_signature_2025-J48"])

    def test_changing_dire_file_changes_context_and_clears_splits(self):
        state = {
            "pdf_current_cohort_signature_2025-J48": self._context(files=["ancien_dire.pdf", "multi.pdf"])["context_id"],
            "ingestion_manual_split_rows": [{"numero_piece": 1}],
        }
        result = self.ns["sync_documentary_cohort_session"](
            state,
            "2025-J48",
            self._context(files=["nouveau_dire.pdf", "multi.pdf"]),
        )
        self.assertTrue(result["changed"])
        self.assertNotIn("ingestion_manual_split_rows", state)

    def test_switching_from_party_25_to_03_updates_downstream_readonly_context(self):
        parties = [
            self._party(25, "Gérard TORDJMAN"),
            self._party(3, "Philippe INGOLD et Véronique INGOLD"),
        ]
        previous = self._context(code=25, files=["ancien.pdf"])
        current = self.ns["documentary_cohort_context"](
            "2025-J48",
            parties[1],
            date(2026, 2, 27),
            "Conseil INGOLD",
            ["ingold.pdf"],
        )
        state = {
            "pdf_current_cohort_signature_2025-J48": previous["context_id"],
            "default_code_partie": "25",
            "split_code_partie_2025-J48": "25",
            "classement_originaux_partie_cible_display_2025-J48": "25 – Gérard TORDJMAN",
            "ingestion_party_display_2025-J48": "25 – Gérard TORDJMAN",
            "ingestion_date_transmission_expert_2025-J48": date(2025, 10, 24),
            "ingestion_auteur_transmission_2025-J48": "Ancien conseil",
        }
        result = self.ns["sync_documentary_cohort_session"](state, "2025-J48", current)
        self.assertTrue(result["changed"])
        self.assertNotIn("default_code_partie", state)
        self.assertNotIn("split_code_partie_2025-J48", state)
        self.assertNotIn("classement_originaux_partie_cible_display_2025-J48", state)
        self.assertNotIn("ingestion_party_display_2025-J48", state)
        self.assertNotIn("ingestion_date_transmission_expert_2025-J48", state)
        self.assertNotIn("ingestion_auteur_transmission_2025-J48", state)
        downstream_party = self.ns["party_from_cohort_context"](current, parties)
        ingestion = self.ns["ingestion_context_from_values"](
            "2025-J48",
            downstream_party,
            date.fromisoformat(current["date_transmission"]),
            current["auteur_transmission"],
        )
        self.assertEqual("03", downstream_party["code_partie"])
        self.assertEqual("03", ingestion["code_partie"])
        self.assertTrue(self.ns["validate_cohort_ingestion_context"](current, ingestion)[0])
        self.assertNotEqual(
            self.ns["documentary_cohort_widget_suffix"](previous),
            self.ns["documentary_cohort_widget_suffix"](current),
        )

    def test_split_filter_rejects_generic_piece_01_and_02_from_old_context(self):
        rows = [
            {"numero_piece": 1, "libelle_final": "Contrat"},
            {"numero_piece": 2, "libelle_final": "Facture"},
        ]
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            good = root / "PIECE n°1 Contrat.pdf"
            stale_1 = root / "PIECE_01_Piece_1.pdf"
            stale_2 = root / "PIECE_02_Piece_2.pdf"
            for path in (good, stale_1, stale_2):
                path.write_text("x", encoding="utf-8")
            selected, diag = self.ns["filter_split_child_paths_for_current_context"](
                [good, stale_1, stale_2],
                rows,
                {},
                "multi.pdf",
            )
        self.assertEqual(["PIECE n°1 Contrat.pdf"], [path.name for path in selected])
        self.assertEqual(2, len(diag["rejected"]))

    def test_sqlite_receives_current_context_code_via_ingestion_context(self):
        ingestion = self.ns["ingestion_context_from_values"]("2025-J48", self._party(42), date(2025, 10, 24), "Maître Alberta SMALL")
        self.assertEqual("42", ingestion["code_partie"])

    def test_expert_doc_id_for_correct_party_prefix_preflight(self):
        rows = [{"expert_doc_id": "25-0001"}, {"expert_doc_id": "26-0001"}]
        preflight = self.ns["documentary_deletion_preflight"](rows, ["25-0001"], "réingestion")
        self.assertTrue(preflight["ok"])
        self.assertEqual(["25-0001"], [row["expert_doc_id"] for row in preflight["rows"]])


if __name__ == "__main__":
    unittest.main()
