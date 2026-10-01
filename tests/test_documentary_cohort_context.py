from __future__ import annotations

import ast
import hashlib
import json
import re
import tempfile
import unittest
import unicodedata
import uuid
from datetime import date
from datetime import datetime
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
        "file_page_count_record",
        "sha256_file",
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
        "documentary_cohort_stable_id",
        "documentary_cohort_context_from_record",
        "restore_documentary_cohort_identity",
        "split_manifest_path",
        "load_split_manifest",
        "split_manifest_parent_key",
        "split_manifest_parent_name",
        "resolve_split_manifest_parent",
        "reconcile_fragmented_split_manifests",
        "piece_number_range_from_parent_filename",
        "migrate_unmanifested_split_parents",
        "upsert_split_manifest",
        "split_manifest_parent_statuses",
        "documentary_cohort_widget_suffix",
        "party_from_cohort_context",
        "cohort_dependent_session_keys",
        "reset_cohort_dependent_session_state",
        "sync_documentary_cohort_session",
        "ingestion_context_from_values",
        "validate_cohort_ingestion_context",
        "expected_split_child_filenames",
        "filter_split_child_paths_for_current_context",
        "detect_piece_ref_details_from_filename",
        "fallback_piece_title_from_filename",
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
        "uuid": uuid,
        "datetime": datetime,
        "fitz": None,
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

    def test_j47_legacy_single_parent_keeps_six_name_matched_children(self):
        rows = [
            {"numero_piece": numero, "fichier_sortie": f"PIECE_{numero:02d}_Titre_{numero}.pdf"}
            for numero in range(2, 8)
        ]
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            children = []
            for row in rows:
                path = root / row["fichier_sortie"]
                path.write_bytes(b"pdf")
                children.append(path)
            unrelated = root / "PIECE_08_Autre_parent.pdf"
            unrelated.write_bytes(b"pdf")
            selected, diag = self.ns["filter_split_child_paths_for_current_context"](
                [*children, unrelated], rows, {}, "304 PIECES N°2 A 7.PDF"
            )
        self.assertEqual(6, len(selected))
        self.assertEqual(["PIECE_08_Autre_parent.pdf"], [row["fichier"] for row in diag["rejected"]])

    def test_expected_split_names_accepts_nom_cible_defensively(self):
        names = self.ns["expected_split_child_filenames"](
            [{"numero_piece": 1, "nom_cible": "PIECE_01_Piece 1.pdf"}], {}
        )
        self.assertIn("piece_01_piece 1.pdf", names)

    def test_parent_piece_ranges_support_j47_and_j54_names(self):
        parse = self.ns["piece_number_range_from_parent_filename"]
        self.assertEqual((2, 7), parse("304 PIECES N°2 A 7.PDF"))
        self.assertEqual((1, 15), parse("102 Nos Pièces 1 à 15.pdf"))
        self.assertEqual((16, 32), parse("Nos Pièces 16 à 32.pdf"))

    def test_guided_split_fixes_global_strategy_without_ui_selector(self):
        source = APP_PATH.read_text(encoding="utf-8-sig")
        self.assertNotIn('st.selectbox("Stratégie de numérotation"', source)
        self.assertIn('strategy = "global"', source)
        self.assertIn('"state": {"last_global": 0}', source)

    def test_existing_complete_parent_range_migrates_without_resplitting(self):
        context = self._context(code=1, when=date(2025, 10, 6), files=["Nos Pièces 1 à 15.pdf"])
        context["affaire"] = "2026-J54"
        context["cohort_id"] = "migration-j54"
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "sources"
            splits = root / "splits"
            source.mkdir()
            splits.mkdir()
            (source / "Nos Pièces 1 à 15.pdf").write_bytes(b"parent")
            for numero in range(1, 16):
                (splits / f"PIECE_{numero:02d}_Piece {numero}.pdf").write_bytes(b"child" + bytes([numero]))
            result = self.ns["migrate_unmanifested_split_parents"](
                splits, source, context, "2026-J54", ["Nos Pièces 1 à 15.pdf"]
            )
            manifest = self.ns["load_split_manifest"](splits, context)
            status = self.ns["split_manifest_parent_statuses"](
                manifest, splits, ["Nos Pièces 1 à 15.pdf"]
            )[0]
        self.assertEqual(["Nos Pièces 1 à 15.pdf"], result["migrated"])
        self.assertEqual([], result["rejected"])
        self.assertTrue(status["complete"])
        self.assertEqual(15, status["existing_child_count"])

    def test_j54_three_parent_manifest_is_cumulative_across_reruns(self):
        context = self._context(
            code=1,
            when=date(2025, 10, 6),
            files=[
                "Nos Pièces 1 à 15.pdf",
                "Nos Pièces 16 à 32.pdf",
                "Nos Pièces 33 à 43.pdf",
            ],
        )
        context["affaire"] = "2026-J54"
        context["cohort_id"] = "2026-J54-partie-01-2025-10-06"
        plans = [
            ("Nos Pièces 1 à 15.pdf", range(1, 16)),
            ("Nos Pièces 16 à 32.pdf", range(16, 33)),
            ("Nos Pièces 33 à 43.pdf", range(33, 44)),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            for parent_name, numbers in plans:
                parent = root / parent_name
                parent.write_bytes(f"parent:{parent_name}".encode())
                pieces, created = [], []
                for numero in numbers:
                    filename = f"PIECE_{numero:02d}_Piece {numero}.pdf"
                    child = root / filename
                    child.write_bytes(f"child:{parent_name}:{numero}".encode())
                    pieces.append({
                        "numero": numero,
                        "start_page": numero,
                        "end_page": numero,
                        "filename": filename,
                        "title": f"Pièce {numero}",
                    })
                    created.append(str(child))
                self.ns["upsert_split_manifest"](
                    root, context, "2026-J54", parent_name, parent, pieces, created
                )
                reloaded = self.ns["load_split_manifest"](root, dict(context))
                statuses = self.ns["split_manifest_parent_statuses"](reloaded, root)
                self.assertEqual(len(statuses), plans.index((parent_name, numbers)) + 1)
            final_manifest = self.ns["load_split_manifest"](root, context)
            final_statuses = self.ns["split_manifest_parent_statuses"](final_manifest, root)
        self.assertEqual(3, len(final_statuses))
        self.assertEqual(43, sum(status["existing_child_count"] for status in final_statuses))
        self.assertTrue(all(status["complete"] for status in final_statuses))

    def test_manifest_parent_identity_prevents_same_piece_number_cross_contamination(self):
        context = self._context(code=1)
        context["cohort_id"] = "same-number-cohort"
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            for parent_name, child_name in (("parent-A.pdf", "A_PIECE_01.pdf"), ("parent-B.pdf", "B_PIECE_01.pdf")):
                parent = root / parent_name
                child = root / child_name
                parent.write_bytes(parent_name.encode())
                child.write_bytes(child_name.encode())
                self.ns["upsert_split_manifest"](
                    root,
                    context,
                    "2026-J54",
                    parent_name,
                    parent,
                    [{"numero": 1, "start_page": 1, "end_page": 1, "filename": child_name}],
                    [child],
                )
            manifest = self.ns["load_split_manifest"](root, context)
            a_status = self.ns["split_manifest_parent_statuses"](manifest, root, ["parent-A.pdf"])[0]
            b_status = self.ns["split_manifest_parent_statuses"](manifest, root, ["parent-B.pdf"])[0]
        self.assertEqual(["A_PIECE_01.pdf"], [child["filename"] for child in a_status["children"]])
        self.assertEqual(["B_PIECE_01.pdf"], [child["filename"] for child in b_status["children"]])

    def test_manifest_attachment_survives_later_label_change(self):
        context = self._context(code=1)
        context["cohort_id"] = "label-change-cohort"
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            parent = root / "parent.pdf"
            child = root / "PIECE_01_Ancien titre.pdf"
            parent.write_bytes(b"parent")
            child.write_bytes(b"child")
            self.ns["upsert_split_manifest"](
                root,
                context,
                "2026-J54",
                parent.name,
                parent,
                [{"numero": 1, "start_page": 1, "end_page": 2, "filename": child.name, "title": "Ancien titre"}],
                [child],
            )
            status = self.ns["split_manifest_parent_statuses"](
                self.ns["load_split_manifest"](root, context), root, [parent.name]
            )[0]
        self.assertTrue(status["complete"])
        self.assertEqual("PIECE_01_Ancien titre.pdf", status["children"][0]["filename"])

    def test_journal_reconstruction_preserves_stable_cohort_identity(self):
        record = {
            "affaire": "2026-J54",
            "cohort_id": "stable-j54-cohort",
            "context_id": "context-j54",
            "code_partie": "01",
            "date_transmission": "2025-10-06",
            "auteur_transmission": "Conseil",
            "files": ["Nos Pièces 1 à 15.pdf"],
        }
        restored = self.ns["documentary_cohort_context_from_record"](record, "2026-J54")
        self.assertEqual("2026-J54", restored["affaire"])
        self.assertEqual("Conseil", restored["auteur_transmission"])
        self.assertEqual("stable-j54-cohort", self.ns["documentary_cohort_stable_id"](restored))

    def test_persistent_cohort_id_survives_file_removal_and_requalification(self):
        original = self._context(files=["dire.pdf", "parent.pdf"])
        record = {
            **original,
            "cohort_id": "immutable-transmission-id",
            "files": ["dire.pdf", "parent.pdf"],
        }
        changed = self._context(files=["parent.pdf"])
        restored = self.ns["restore_documentary_cohort_identity"](changed, [record])
        self.assertNotEqual(original["context_id"], changed["context_id"])
        self.assertEqual("immutable-transmission-id", restored["cohort_id"])

    def _write_one_parent_manifest(self, root, cohort_id, parent_name, parent_bytes, child_name):
        context = self._context(code=1, when=date(2025, 10, 6), files=[parent_name])
        context.update({"affaire": "2026-J54", "cohort_id": cohort_id})
        parent = root / parent_name
        child = root / child_name
        parent.write_bytes(parent_bytes)
        child.write_bytes((child_name + cohort_id).encode())
        self.ns["upsert_split_manifest"](
            root,
            context,
            "2026-J54",
            parent_name,
            parent,
            [{"numero": 1, "start_page": 1, "end_page": 1, "filename": child_name}],
            [child],
        )
        return context

    def test_fragmented_manifest_recovers_unique_parent_by_name_and_sha256(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            old_context = self._write_one_parent_manifest(
                root, "old-cohort", "parent-A.pdf", b"parent-a", "A_PIECE_01.pdf"
            )
            current_context = self._write_one_parent_manifest(
                root, "current-cohort", "parent-B.pdf", b"parent-b", "B_PIECE_01.pdf"
            )
            current_context["files"] = ["parent-A.pdf", "parent-B.pdf"]
            diag = self.ns["reconcile_fragmented_split_manifests"](
                root, current_context, ["parent-A.pdf", "parent-B.pdf"], root
            )
            manifest = self.ns["load_split_manifest"](root, current_context)
        self.assertEqual(["parent-A.pdf"], [row["parent_document"] for row in diag["recovered"]])
        self.assertEqual(2, len(manifest["parents"]))
        self.assertEqual("old-cohort", old_context["cohort_id"])

    def test_document_prefix_alias_requires_identical_parent_sha256(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._write_one_parent_manifest(
                root, "old-cohort", "parent.pdf", b"same-parent", "OLD_PIECE_01.pdf"
            )
            current = self._write_one_parent_manifest(
                root, "current-cohort", "seed.pdf", b"seed", "SEED_PIECE_01.pdf"
            )
            (root / "102 parent.pdf").write_bytes(b"same-parent")
            diag = self.ns["reconcile_fragmented_split_manifests"](
                root, current, ["102 parent.pdf"], root
            )
            recovered = self.ns["load_split_manifest"](root, current)
        self.assertTrue(diag["recovered"][0]["alias_documentaire"])
        self.assertIn(self.ns["split_manifest_parent_key"]("parent.pdf"), recovered["parents"])
        self.assertNotIn(self.ns["split_manifest_parent_key"]("102 parent.pdf"), recovered["parents"])

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._write_one_parent_manifest(
                root, "old-cohort", "parent.pdf", b"old-parent", "OLD_PIECE_01.pdf"
            )
            current = self._write_one_parent_manifest(
                root, "current-cohort", "seed.pdf", b"seed", "SEED_PIECE_01.pdf"
            )
            (root / "102 parent.pdf").write_bytes(b"different-parent")
            diag = self.ns["reconcile_fragmented_split_manifests"](
                root, current, ["102 parent.pdf"], root
            )
        self.assertEqual([], diag["recovered"])
        self.assertEqual("aucun_manifest_compatible", diag["not_found"][0]["reason"])

    def test_fragmented_manifest_refuses_ambiguous_parent(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            self._write_one_parent_manifest(
                root, "old-one", "parent.pdf", b"same-parent", "ONE_PIECE_01.pdf"
            )
            self._write_one_parent_manifest(
                root, "old-two", "parent.pdf", b"same-parent", "TWO_PIECE_01.pdf"
            )
            current = self._write_one_parent_manifest(
                root, "current", "seed.pdf", b"seed", "SEED_PIECE_01.pdf"
            )
            diag = self.ns["reconcile_fragmented_split_manifests"](
                root, current, ["parent.pdf"], root
            )
            manifest = self.ns["load_split_manifest"](root, current)
        self.assertEqual([], diag["recovered"])
        self.assertEqual(2, diag["ambiguous"][0]["candidate_count"])
        self.assertEqual(1, len(manifest["parents"]))

    def test_j54_prefixed_ui_parents_resolve_against_current_manifest_without_duplication(self):
        context = self._context(code=1, when=date(2025, 10, 6))
        context.update({"affaire": "2026-J54", "cohort_id": "db79f45cbea5d22b"})
        plans = [
            ("Nos Pièces 1 à 15.pdf", "102 Nos Pièces 1 à 15.pdf", range(1, 16)),
            ("Nos Pièces 16 à 32.pdf", "103 Nos Pièces 16 à 32.pdf", range(16, 33)),
            ("Nos Pièces 33 à 43.pdf", "104 Nos Pièces 33 à 43.pdf", range(33, 44)),
        ]
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sources = root / "sources"
            splits = root / "splits"
            sources.mkdir()
            splits.mkdir()
            for manifest_name, selected_name, numbers in plans:
                parent_bytes = f"parent:{manifest_name}".encode()
                manifest_parent = sources / manifest_name
                selected_parent = sources / selected_name
                manifest_parent.write_bytes(parent_bytes)
                selected_parent.write_bytes(parent_bytes)
                pieces, created = [], []
                for numero in numbers:
                    filename = f"PIECE_{numero:02d}_Piece {numero}.pdf"
                    child = splits / filename
                    child.write_bytes(f"child:{numero}".encode())
                    pieces.append({
                        "numero": numero,
                        "start_page": numero,
                        "end_page": numero,
                        "filename": filename,
                    })
                    created.append(child)
                self.ns["upsert_split_manifest"](
                    splits, context, "2026-J54", manifest_name, manifest_parent, pieces, created
                )
            selected_names = [selected for _, selected, _ in plans]
            manifest_before = self.ns["load_split_manifest"](splits, context)
            statuses = self.ns["split_manifest_parent_statuses"](
                manifest_before, splits, selected_names, sources
            )
            manifest_after = self.ns["load_split_manifest"](splits, context)

        self.assertEqual(3, len(statuses))
        self.assertEqual([15, 17, 11], [status["expected_child_count"] for status in statuses])
        self.assertEqual([15, 17, 11], [status["existing_child_count"] for status in statuses])
        self.assertTrue(all(status["complete"] for status in statuses))
        self.assertEqual(43, sum(status["existing_child_count"] for status in statuses))
        self.assertEqual(selected_names, [status["parent_document"] for status in statuses])
        self.assertEqual(
            [manifest_name for manifest_name, _, _ in plans],
            [status["manifest_parent_document"] for status in statuses],
        )
        self.assertEqual(["sha256_alias_current"] * 3, [status["matched_by"] for status in statuses])
        self.assertEqual(3, len(manifest_before["parents"]))
        self.assertEqual(manifest_before, manifest_after)

    def test_current_manifest_alias_with_different_sha256_is_refused(self):
        context = self._context(code=1, when=date(2025, 10, 6))
        context.update({"affaire": "2026-J54", "cohort_id": "current"})
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            sources = root / "sources"
            splits = root / "splits"
            sources.mkdir()
            splits.mkdir()
            manifest_parent = sources / "Nos Pièces 16 à 32.pdf"
            selected_parent = sources / "103 Nos Pièces 16 à 32.pdf"
            child = splits / "PIECE_16.pdf"
            manifest_parent.write_bytes(b"original-parent")
            selected_parent.write_bytes(b"different-parent")
            child.write_bytes(b"child")
            self.ns["upsert_split_manifest"](
                splits,
                context,
                "2026-J54",
                manifest_parent.name,
                manifest_parent,
                [{"numero": 16, "start_page": 1, "end_page": 1, "filename": child.name}],
                [child],
            )
            manifest = self.ns["load_split_manifest"](splits, context)
            statuses = self.ns["split_manifest_parent_statuses"](
                manifest, splits, [selected_parent.name], sources
            )
        self.assertEqual([], statuses)

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
