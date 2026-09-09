from __future__ import annotations

import ast
import csv
import getpass
import hashlib
import io
import json
import os
import re
import shutil
import socket
import sqlite3
import tempfile
import unittest
import unicodedata
import uuid
from datetime import date, datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_documentary_helpers():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    keep_assignments = {
        "WINDOWS_FORBIDDEN",
        "JURIDICTION_FOLDER_REL",
        "JURIDICTION_SOURCE_RECORD",
        "DOCUMENT_ROLE_ALIASES",
        "DOCUMENT_TITLE_EXTENSIONS",
        "PIECE_FILE_RE",
        "EXPLICIT_PIECE_TITLE_RE",
        "NUMBERED_BCP_TITLE_RE",
        "NUMBER_ONLY_BCP_RE",
        "DOCUMENTARY_HINT_RE",
        "PIECE_REF_TEXT_SUFFIXES",
        "INGESTION_DOCUMENT_SUBJECT_ORDER",
        "LEGACY_CLASSIFIABLE_ACTIONS",
        "PCFIXE_SMB_TEST_TIMEOUT_SECONDS",
        "ANNOTATION_JOBS_DIAGNOSTIC_TTL_SECONDS",
        "ANNOTATION_FILE_PROFILE_TTL_SECONDS",
        "ANNOTATION_JOB_SCAN_LIMIT_LIGHT",
        "ANNOTATION_JOB_SCAN_LIMIT_DETAILED",
        "ANNOTATION_UNC_LIST_TIMEOUT_SECONDS",
        "ANNOTATION_SLOW_BLOCK_SECONDS",
        "ANNOTATION_JOBS_SMB_TCP_TIMEOUT_SECONDS",
        "ANNOTATION_LOCAL_PATH_TIMEOUT_SECONDS",
        "ANNOTATION_UNC_RESOURCE_TIMEOUT_SECONDS",
        "ANNOTATION_UNC_RESOURCE_RETRY_TIMEOUT_SECONDS",
        "ANNOTATION_JOBS_FOLDERS",
    }
    body = []
    for node in tree.body:
        if isinstance(node, ast.Assign):
            names = {target.id for target in node.targets if isinstance(target, ast.Name)}
            if names & keep_assignments:
                body.append(node)
        elif isinstance(node, ast.FunctionDef):
            body.append(node)
    module = ast.Module(body=body, type_ignores=[])
    ast.fix_missing_locations(module)
    ns = {
        "csv": csv,
        "date": date,
        "datetime": datetime,
        "getpass": getpass,
        "hashlib": hashlib,
        "io": io,
        "json": json,
        "os": os,
        "Path": Path,
        "re": re,
        "shutil": shutil,
        "socket": socket,
        "sqlite3": sqlite3,
        "unicodedata": unicodedata,
        "uuid": uuid,
        "pj": lambda *parts: str(Path(str(parts[0])).joinpath(*[str(p) for p in parts[1:] if str(p)])),
    }
    exec(compile(module, str(APP_PATH), "exec"), ns)
    ns["configured_admin_path"] = lambda root, cfg=None, key="logs": str(
        Path(root) / ((cfg or {}).get("paths") or {}).get(key, r"AA_Expert_Admin\_Logs")
    )
    ns["generate_document_state_exports"] = lambda root, cfg, state_no: {"state_no": state_no, "ok": True}
    return ns


class DocumentaryMaintenanceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = _load_documentary_helpers()

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.cfg = {
            "project_id": "AFF-A",
            "paths": {
                "logs": r"AA_Expert_Admin\_Logs",
                "sqlite": r"_DB\project.sqlite",
            },
        }
        (self.root / "AA_Expert_Admin" / "_Logs").mkdir(parents=True)
        (self.root / "01_Partie_01_Alpha").mkdir()
        (self.root / "02_Partie_02_Beta").mkdir()

    def tearDown(self):
        self.tmp.cleanup()

    def _record(
        self,
        expert_doc_id: str,
        *,
        code: int = 1,
        transmission_id: str = "T1",
        filename: str = "piece.pdf",
        label: str = "Piece A",
        date_value: str = "2026-01-02",
        deposant: str = "Alpha Avocat",
        role: str = "piece",
    ) -> dict:
        folder = f"{code:02d}_Partie_{code:02d}_{'Alpha' if code == 1 else 'Beta'}"
        physical = self.root / folder / filename
        if not physical.exists():
            physical.write_text("pdf placeholder", encoding="utf-8")
        return {
            "action": "classify_originals_to_party",
            "aff_id": "AFF-A",
            "source_type": "partie",
            "transmission_id": transmission_id,
            "date_transmission_expert": date_value,
            "type_transmission": "lettre" if role == "lettre_dire" else "dire",
            "auteur_transmission": deposant,
            "party": {"code_partie": code, "nom": f"Partie {code}", "folder_rel": folder},
            "copied": [
                {
                    "transmission_id": transmission_id,
                    "source_name": filename,
                    "destination": str(physical),
                    "expert_doc_id": expert_doc_id,
                    "numero_expert": expert_doc_id,
                    "document_role": role,
                    "type_document": self.ns["document_type_from_role"](role),
                    "libelle_final": label,
                    "libelle_affichage": label,
                    "page_count": 1,
                }
            ],
            "files": [
                {
                    "transmission_id": transmission_id,
                    "source_name": filename,
                    "destination": str(physical),
                    "expert_doc_id": expert_doc_id,
                    "numero_expert": expert_doc_id,
                    "document_role": role,
                    "type_document": self.ns["document_type_from_role"](role),
                    "libelle_final": label,
                    "libelle_affichage": label,
                    "page_count": 1,
                }
            ],
        }

    def _insert_sqlite(self, expert_doc_id: str, code: str = "01", filename: str = "piece.pdf") -> Path:
        db_path = self.ns["ensure_documents_sqlite_schema"](str(self.root), self.cfg)
        conn = sqlite3.connect(str(db_path))
        try:
            conn.execute(
                """
                INSERT INTO Documents(
                    id_document, numero_expert, code_partie, description, date_reception,
                    emetteur, pages, est_dire, chemin_nas, sha256, chemin_local,
                    role_document, nom_original, nom_cible
                )
                VALUES (?, ?, ?, ?, ?, ?, 1, 0, '', ?, ?, 'piece', ?, ?)
                """,
                (
                    f"id-{expert_doc_id}",
                    expert_doc_id,
                    code,
                    filename,
                    "2026-01-02",
                    "Alpha Avocat",
                    f"sha-{expert_doc_id}",
                    str(self.root / f"{int(code):02d}_Partie_{int(code):02d}_Alpha" / filename),
                    filename,
                    filename,
                ),
            )
            conn.commit()
        finally:
            conn.close()
        return db_path

    def _states_ids(self, records):
        states = self.ns["build_document_states"](records, "AFF-A", str(self.root), self.cfg)
        return {
            "etat1": {row.get("expert_doc_id") for row in states["etat1_dires_messages_courriers"]},
            "etat2": {row.get("expert_doc_id") for row in states["etat2_detail_documents_fournis"]},
            "etat3": {row.get("expert_doc_id") for row in states["etat3_tableau_recapitulatif"]},
            "etat4": {row.get("expert_doc_id") for row in states["etat4_documents_recus"]},
        }

    def _delete(self, records, ids, reason="erreur historique", rebuild=False):
        rows = self.ns["build_document_maintenance_rows"](records, "AFF-A", str(self.root), self.cfg)
        preflight = self.ns["documentary_deletion_preflight"](rows, ids, reason)
        self.assertTrue(preflight["ok"], preflight)
        return self.ns["apply_documentary_logical_deletion"](
            str(self.root),
            self.cfg,
            preflight["rows"],
            reason,
            rebuild_states=rebuild,
        )

    def test_single_logical_deletion_disappears_from_states_1_to_4(self):
        records = [
            self._record("01-001", filename="lettre.pdf", label="Lettre", role="lettre_dire"),
            self._record("01-002", filename="piece.pdf", label="Piece"),
        ]
        self._insert_sqlite("01-001", "01", "lettre.pdf")
        self._delete(records, ["01-001"])
        ids = self._states_ids(records)
        self.assertNotIn("01-001", ids["etat1"])
        self.assertNotIn("01-001", ids["etat2"])
        self.assertNotIn("01-001", ids["etat3"])
        self.assertNotIn("01-001", ids["etat4"])

    def test_other_documents_unchanged_after_deletion(self):
        records = [self._record("01-001"), self._record("01-002", filename="other.pdf", label="Other")]
        before = self._states_ids(records)
        self._delete(records, ["01-001"])
        after = self._states_ids(records)
        self.assertIn("01-002", before["etat2"])
        self.assertIn("01-002", after["etat2"])

    def test_multiple_deletion(self):
        records = [self._record("01-001"), self._record("01-002", filename="b.pdf")]
        self._delete(records, ["01-001", "01-002"])
        ids = self._states_ids(records)
        self.assertNotIn("01-001", ids["etat2"])
        self.assertNotIn("01-002", ids["etat2"])

    def test_exact_duplicate_can_be_deleted_individually_by_expert_doc_id(self):
        records = [
            self._record("01-001", filename="same.pdf", label="Même libellé", transmission_id="T1"),
            self._record("01-002", filename="same.pdf", label="Même libellé", transmission_id="T2"),
        ]
        self._delete(records, ["01-001"])
        ids = self._states_ids(records)
        self.assertNotIn("01-001", ids["etat2"])
        self.assertIn("01-002", ids["etat2"])

    def test_same_label_other_document_not_deleted(self):
        records = [
            self._record("01-001", filename="a.pdf", label="Même libellé"),
            self._record("02-001", code=2, filename="b.pdf", label="Même libellé"),
        ]
        self._delete(records, ["01-001"])
        self.assertIn("02-001", self._states_ids(records)["etat2"])

    def test_wrong_deposant_deleted_then_same_number_not_reused_in_sqlite(self):
        records = [self._record("01-0001", deposant="Mauvais déposant")]
        self._insert_sqlite("01-0001", "01")
        self._delete(records, ["01-0001"])
        db_path = self.ns["affaire_sqlite_path_from_root"](str(self.root), self.cfg)
        conn = sqlite3.connect(str(db_path))
        try:
            status = conn.execute("SELECT statut_document FROM Documents WHERE numero_expert='01-0001'").fetchone()[0]
            next_id = self.ns["next_sqlite_numero_expert"](conn, "01")
        finally:
            conn.close()
        self.assertEqual("supprimé", status)
        self.assertEqual("01-0002", next_id)

    def test_physical_file_is_preserved(self):
        records = [self._record("01-001", filename="keep.pdf")]
        physical = Path(records[0]["copied"][0]["destination"])
        self._delete(records, ["01-001"])
        self.assertTrue(physical.exists())

    def test_no_renumbering_in_states(self):
        records = [self._record("01-001"), self._record("01-003", filename="c.pdf")]
        self._delete(records, ["01-001"])
        self.assertIn("01-003", self._states_ids(records)["etat2"])

    def test_stale_preflight_is_blocked(self):
        records = [self._record("01-001")]
        rows = self.ns["build_document_maintenance_rows"](records, "AFF-A", str(self.root), self.cfg)
        stale = self.ns["documentary_deletion_preflight"](rows, ["01-001"], "motif", expected_signature="stale")
        self.assertFalse(stale["ok"])
        self.assertIn("preflight_perime", stale["blockers"])

    def test_confirmation_text_is_bound_to_exact_selection(self):
        records = [self._record("01-001"), self._record("01-002", filename="b.pdf")]
        rows = self.ns["build_document_maintenance_rows"](records, "AFF-A", str(self.root), self.cfg)
        one = self.ns["documentary_deletion_preflight"](rows, ["01-001"], "motif")
        two = self.ns["documentary_deletion_preflight"](rows, ["01-001", "01-002"], "motif")
        self.assertEqual("SUPPRIMER 01-001", one["confirmation_text"])
        self.assertEqual("SUPPRIMER 01-001, 01-002", two["confirmation_text"])
        self.assertNotEqual(one["selection_signature"], two["selection_signature"])

    def test_append_only_journal_created(self):
        records = [self._record("01-001")]
        result = self._delete(records, ["01-001"])
        log_path = Path(result["override_log"])
        self.assertTrue(log_path.exists())
        self.assertIn("documentary_entries_logical_delete", log_path.read_text(encoding="utf-8"))

    def test_no_cross_case_contamination(self):
        other_root = self.root.parent / (self.root.name + "_other")
        other_root.mkdir()
        try:
            (other_root / "AA_Expert_Admin" / "_Logs").mkdir(parents=True)
            other_cfg = {"project_id": "AFF-B", "paths": {"logs": r"AA_Expert_Admin\_Logs", "sqlite": r"_DB\project.sqlite"}}
            self._delete([self._record("01-001")], ["01-001"])
            self.assertFalse(self.ns["documents_registry_overrides_path"](str(other_root), other_cfg).exists())
        finally:
            shutil.rmtree(other_root, ignore_errors=True)

    def test_fingerprinted_tombstone_does_not_delete_same_expert_id_other_document(self):
        docs = [
            {
                "expert_doc_id": "01-001",
                "transmission_id": "T1",
                "fichier_source": "a.pdf",
                "chemin": str(self.root / "01_Partie_01_Alpha" / "a.pdf"),
                "numero_piece": "1",
                "statut_document": "actif",
            },
            {
                "expert_doc_id": "01-001",
                "transmission_id": "T2",
                "fichier_source": "b.pdf",
                "chemin": str(self.root / "01_Partie_01_Alpha" / "b.pdf"),
                "numero_piece": "2",
                "statut_document": "actif",
            },
        ]
        fp_first = self.ns["document_registry_fingerprint"](docs[0])
        overrides = {
            self.ns["document_registry_override_key"]("01-001", fp_first): {
                "expert_doc_id": "01-001",
                "registry_fingerprint": fp_first,
                "statut_document": "supprimé",
                "suppression_logique": True,
            }
        }
        applied = self.ns["apply_document_registry_overrides"](docs, overrides)
        self.assertEqual(applied[0]["statut_document"], "supprimé")
        self.assertEqual(applied[1]["statut_document"], "actif")

    def test_preflight_blocks_ambiguous_expert_doc_id_without_row_key(self):
        rows = [
            {
                "expert_doc_id": "01-001",
                "code_partie": "01",
                "transmissions": "T1",
                "registry_fingerprint": "T1|a.pdf|path-a|1|",
            },
            {
                "expert_doc_id": "01-001",
                "code_partie": "01",
                "transmissions": "T2",
                "registry_fingerprint": "T2|b.pdf|path-b|2|",
            },
        ]
        ambiguous = self.ns["documentary_deletion_preflight"](rows, ["01-001"], "motif")
        self.assertFalse(ambiguous["ok"])
        self.assertIn("selection_ambigue", ambiguous["blockers"])
        row_key = self.ns["documentary_maintenance_row_key"](rows[0])
        targeted = self.ns["documentary_deletion_preflight"](rows, [row_key], "motif")
        self.assertTrue(targeted["ok"], targeted)
        self.assertEqual(targeted["rows"], [rows[0]])

    def test_sqlite_marking_uses_primary_key_not_numero_expert(self):
        db_path = self.ns["ensure_documents_sqlite_schema"](str(self.root), self.cfg)
        conn = sqlite3.connect(str(db_path))
        try:
            for id_document, filename in (("pk-a", "a.pdf"), ("pk-b", "b.pdf")):
                conn.execute(
                    """
                    INSERT INTO Documents(
                        id_document, numero_expert, code_partie, description, date_reception,
                        emetteur, pages, est_dire, chemin_nas, sha256, chemin_local,
                        role_document, nom_original, nom_cible
                    )
                    VALUES (?, '01-001', '01', ?, '2026-01-02', 'Alpha Avocat',
                            1, 0, '', ?, ?, 'piece', ?, ?)
                    """,
                    (
                        id_document,
                        filename,
                        f"sha-{id_document}",
                        str(self.root / "01_Partie_01_Alpha" / filename),
                        filename,
                        filename,
                    ),
                )
            conn.commit()
        finally:
            conn.close()
        result = self.ns["mark_documentary_sqlite_entries_deleted"](
            str(self.root),
            self.cfg,
            [{"expert_doc_id": "01-001", "sqlite_id_document": "pk-b"}],
            "erreur ciblée",
        )
        self.assertEqual(result["marked_deleted_expert_doc_ids"], ["01-001"])
        conn = sqlite3.connect(str(db_path))
        try:
            statuses = dict(conn.execute("SELECT id_document, statut_document FROM Documents").fetchall())
        finally:
            conn.close()
        self.assertIn(statuses["pk-a"], (None, "actif"))
        self.assertEqual(statuses["pk-b"], "supprimé")

    def test_skipped_item_tombstoned_by_fingerprint_does_not_resurrect_with_new_generated_id(self):
        skipped_path = self.root / "01_Partie_01_Alpha" / "skipped.pdf"
        kept_path = self.root / "01_Partie_01_Alpha" / "kept.pdf"
        skipped_path.write_text("skipped physical pdf", encoding="utf-8")
        kept_path.write_text("kept physical pdf", encoding="utf-8")
        records = [
            self._record("01-001", filename="kept.pdf", label="Kept", transmission_id="T1"),
            {
                "action": "classify_originals_to_party",
                "aff_id": "AFF-A",
                "source_type": "partie",
                "transmission_id": "T2",
                "date_transmission_expert": "2026-01-02",
                "auteur_transmission": "Alpha Avocat",
                "party": {"code_partie": 1, "nom": "Partie 1", "folder_rel": "01_Partie_01_Alpha"},
                "skipped": [
                    {
                        "transmission_id": "T2",
                        "source_name": "skipped.pdf",
                        "destination": str(skipped_path),
                        "reason": "destination existe déjà",
                        "document_role": "piece",
                        "type_document": "piece",
                        "numero_piece": 7,
                        "libelle_final": "Skipped generated document",
                        "libelle_affichage": "Skipped generated document",
                        "page_count": 1,
                    }
                ],
            },
        ]
        first_rows = self.ns["build_documents_registry"](records, "AFF-A", str(self.root), self.cfg)
        skipped_row = next(row for row in first_rows if row["transmission_id"] == "T2")
        generated_id = skipped_row["expert_doc_id"]
        fingerprint = skipped_row["registry_fingerprint"]
        self.assertEqual("generated", skipped_row["expert_doc_id_source"])
        self.assertEqual(fingerprint, self.ns["document_registry_fingerprint"]({
            "transmission_id": "T2",
            "fichier_source": "skipped.pdf",
            "chemin": str(skipped_path),
            "numero_piece": 7,
            "sous_piece": "",
        }))
        self.assertEqual(
            fingerprint,
            self.ns["document_registry_fingerprint"]({
                "transmission_id": "T2",
                "fichier_source": "skipped.pdf",
                "chemin": str(skipped_path),
                "numero_piece": 7,
                "sous_piece": "",
                "expert_doc_id": "99-9999",
            }),
        )

        preflight = self.ns["documentary_deletion_preflight"](
            self.ns["build_document_maintenance_rows"](records, "AFF-A", str(self.root), self.cfg),
            [self.ns["documentary_maintenance_row_key"](skipped_row)],
            "scorie skipped",
        )
        self.assertTrue(preflight["ok"], preflight)
        self.ns["apply_documentary_logical_deletion"](
            str(self.root),
            self.cfg,
            preflight["rows"],
            "scorie skipped",
            rebuild_states=False,
        )

        rebuilt_rows = self.ns["build_documents_registry"](records, "AFF-A", str(self.root), self.cfg)
        rebuilt_ids = {row["expert_doc_id"] for row in rebuilt_rows}
        rebuilt_fingerprints = {row["registry_fingerprint"] for row in rebuilt_rows}
        self.assertNotIn(generated_id, rebuilt_ids)
        self.assertNotIn(fingerprint, rebuilt_fingerprints)
        self.assertEqual({"01-001"}, rebuilt_ids)


if __name__ == "__main__":
    unittest.main()
