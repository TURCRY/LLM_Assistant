from __future__ import annotations

import ast
import csv
import hashlib
import json
import os
import re
import shutil
import time
import traceback
import unittest
import uuid
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
ACTION_OPTIONS = ["--reset-vlm", "1", "--vlm-strict", "1"]
WEAK_RERUN_OPTIONS = ["--rerun-weak", "1"]


def _load_annotation_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    names = {
        "_atomic_write_json",
        "_valid_file_path",
        "_infos_declared_path",
        "_annotation_localize_pcfixe_job_path",
        "_annotation_prepare_batch_runtime_infos",
        "_annotation_resolve_batch_submission_notice",
        "_annotation_cache_prune",
        "_annotation_perf_begin",
        "_annotation_perf_end",
        "_annotation_list_json_candidates",
        "_annotation_jobs_roots",
        "_annotation_job_details",
        "_annotation_latest_completed_batch",
        "_annotation_has_newer_blocking_batch",
        "_annotation_latest_non_completed_batch",
        "_annotation_report_preflight",
        "_annotation_report_ui_state",
        "_annotation_batch_action_state",
        "_path_accessible_quick",
        "_csv_kind_from_path",
        "_read_semicolon_csv",
        "_csv_schema_ok",
        "_csv_row_count",
        "_csv_duplicate_values",
        "_csv_missing_key_rows",
        "_photo_key_basename",
        "_annotation_csv_join_audit",
        "_annotation_parse_time_value",
        "_annotation_detail_sort_time",
        "_annotation_verified_batch_from_stamp",
        "_annotation_file_profile",
        "_photo_report_job_preview",
        "submit_annotation_photos_batch_job",
        "submit_annotation_photos_batch_publish_retry_job",
    }
    module = ast.Module(
        body=[node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names],
        type_ignores=[],
    )
    ast.fix_missing_locations(module)
    ns = {
        "Path": Path,
        "re": re,
        "csv": csv,
        "json": json,
        "os": os,
        "time": time,
        "uuid": uuid,
        "datetime": datetime,
        "PHOTO_BATCH_ACTIONS": {
            "initial": {
                "label": "initial",
                "status_label": "initial",
                "profile": "vlm_strict",
                "options": ACTION_OPTIONS,
            },
            "weak_rerun": {
                "label": "weak",
                "status_label": "weak",
                "profile": "rerun_weak",
                "options": WEAK_RERUN_OPTIONS,
            },
        },
        "PHOTO_BATCH_PUBLISH_RETRY_KEY": "publish_retry",
        "PHOTO_BATCH_PUBLISH_RETRY_LABEL": "reprise publication NAS",
        "ANNOTATION_JOBS_FOLDERS": ("done", "failed", "running", "queued", "work"),
        "ANNOTATION_RESOURCE_MTIME_TOLERANCE_SECONDS": 5.0,
        "ANNOTATION_JOBS_DIAGNOSTIC_TTL_SECONDS": 45.0,
        "ANNOTATION_FILE_PROFILE_TTL_SECONDS": 45.0,
        "ANNOTATION_JOB_SCAN_LIMIT_LIGHT": 20,
        "ANNOTATION_JOB_SCAN_LIMIT_DETAILED": 50,
        "ANNOTATION_UNC_LIST_TIMEOUT_SECONDS": 4.0,
        "ANNOTATION_SLOW_BLOCK_SECONDS": 15.0,
        "ANNOTATION_JOBS_SMB_TCP_TIMEOUT_SECONDS": 0.75,
        "PCFIXE_SMB_TEST_TIMEOUT_SECONDS": 0.1,
        "ANNOTATION_LOCAL_PATH_TIMEOUT_SECONDS": 0.20,
        "ANNOTATION_UNC_RESOURCE_TIMEOUT_SECONDS": 3.0,
        "ANNOTATION_UNC_RESOURCE_RETRY_TIMEOUT_SECONDS": 0.75,
        "NAS_AFFAIRES_ROOT": Path(r"C:\nas"),
        "AFFAIRES_ROOT": r"C:\Affaires",
        "PCFIXE_AFFAIRES_ROOT": Path(r"C:\Affaires"),
        "PHOTO_REPORT_JOB_SUPPORTED": True,
        "PHOTOS_CSV_REPORT_KEY_COLUMNS": ("photo_rel_native", "nom_fichier_image"),
        "PHOTOS_BATCH_REQUIRED_COLUMNS": ("photo_rel_native", "batch_status", "batch_id", "batch_ts"),
        "_ANNOTATION_JOB_DETAILS_CACHE": {},
        "_ANNOTATION_FILE_PROFILE_CACHE": {},
    }

    def load_json(path, default):
        try:
            return json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception:
            return default

    def sha256_file(path):
        h = hashlib.sha256()
        with Path(path).open("rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()

    ns.update(
        {
            "load_json": load_json,
            "traceback": traceback,
            "_sha256_file": sha256_file,
            "_path_accessible_quick": lambda path: (Path(path).exists(), ""),
            "_test_path_with_timeout": lambda path: (Path(path).exists(), ""),
            "_unc_host": lambda path: "",
            "_safe_job_token": lambda value: re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "").strip()).strip("_") or "na",
            "_job_failure_details": lambda job_id: {},
            "get_pcfixe_jobs_queued_dir": lambda: ns["_queue_dir"],
            "_assert_annotation_pcfixe_resources_ready": lambda *args, **kwargs: {},
            "preflight_pcfixe_target_dir": lambda *args, **kwargs: {},
            "_annotation_canonical_paths": lambda affaire, captation: ns["_paths"],
            "get_pcfixe_affaires_root": lambda: Path(r"\\10.0.1.10\Affaires"),
            "_annotation_tcp_port_open": lambda host, port: (True, "ok"),
            "_annotation_pcfixe_job_hosts": lambda: ["10.0.1.10", "192.168.0.120", "192.168.0.155"],
            "_pcfixe_vpn_active": lambda: True,
        }
    )
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns


class AnnotationWordBatchManifestTests(unittest.TestCase):
    def setUp(self):
        temp_root = Path(os.environ.get("LLM_ASSISTANT_TEST_TMP") or r"C:\CodexWorkspace\.tmp_llm_assistant_tests")
        if not temp_root.parent.exists():
            temp_root = APP_PATH.parent / "temp"
        temp_root.mkdir(exist_ok=True)
        self.base = temp_root / f"case_{uuid.uuid4().hex}"
        self.base.mkdir(parents=True, exist_ok=False)
        self.ns = _load_annotation_functions()
        self.ns["_queue_dir"] = self.base / "queue"
        self.ns["_queue_dir"].mkdir()
        self.jobs = self.base / "_jobs"
        for folder in ("done", "failed", "running", "queued", "work", "logs"):
            (self.jobs / folder).mkdir(parents=True, exist_ok=True)
        self.photos = self.base / "nas" / "photos.csv"
        self.batch = self.base / "nas" / "photos_batch.csv"
        self.infos = self.base / "nas" / "infos_projet.json"
        self.photos.parent.mkdir(parents=True, exist_ok=True)
        self.photos.write_text("photo_rel_native;nom_fichier_image\np1;P1.JPG\n", encoding="utf-8")
        self.batch.write_text("photo_rel_native;batch_status;batch_id;batch_ts\np1;OK;b1;2026-07-24T12:00:00\n", encoding="utf-8")
        self.infos.write_text("{}", encoding="utf-8")
        self.pcfixe_root = self.base / "pcfixe_affaires"
        self.pcfixe_root.mkdir(parents=True, exist_ok=True)
        self.ns["PCFIXE_AFFAIRES_ROOT"] = self.pcfixe_root
        self.ns["NAS_AFFAIRES_ROOT"] = Path(r"\\192.168.1.20\Affaires")
        self.paths = {
            "nas_photos": self.photos,
            "nas_photos_batch": self.batch,
            "nas_infos": self.infos,
            "nas_report_dir": self.base / "nas" / "compte_rendu_LLM",
            "nas_report_out_dir": self.base / "nas" / "compte_rendu_LLM" / "out",
            "pcfixe_trans_dir": self.pcfixe_root / "2025-J47" / "AF_Expert_ASR" / "transcriptions" / "cap",
            "pcfixe_unc_trans_dir": self.base / "pcfixe_unc" / "2025-J47" / "AF_Expert_ASR" / "transcriptions" / "cap",
            "pcfixe_infos": self.pcfixe_root / "2025-J47" / "AF_Expert_ASR" / "transcriptions" / "cap" / "infos_projet.json",
            "pcfixe_unc_infos": self.base / "pcfixe_unc" / "2025-J47" / "AF_Expert_ASR" / "transcriptions" / "cap" / "infos_projet.json",
            "pcfixe_report_dir": self.base / "pc" / "compte_rendu_LLM",
            "pcfixe_photos": self.pcfixe_root / "2025-J47" / "AE_Expert_captations" / "cap" / "photos" / "photos.csv",
            "pcfixe_photos_batch": self.pcfixe_root / "2025-J47" / "AE_Expert_captations" / "cap" / "photos" / "photos_batch.csv",
        }
        self.paths["pcfixe_photos"].parent.mkdir(parents=True, exist_ok=True)
        self.paths["pcfixe_trans_dir"].mkdir(parents=True, exist_ok=True)
        self.paths["pcfixe_unc_trans_dir"].mkdir(parents=True, exist_ok=True)
        self.paths["pcfixe_photos"].write_text("photo_rel_native;nom_fichier_image\np1;P1.JPG\n", encoding="utf-8")
        self.paths["pcfixe_photos_batch"].write_text("photo_rel_native;batch_status;batch_id;batch_ts\np1;OK;b1;2026-07-24T12:00:00\n", encoding="utf-8")
        self.ns["_paths"] = self.paths
        self.ns["_annotation_existing_infos_path"] = lambda paths: self.infos
        self.original_annotation_jobs_roots = self.ns["_annotation_jobs_roots"]
        self.ns["_annotation_jobs_roots"] = lambda *args, **kwargs: ([{"root": str(self.jobs), "label": "test", "source": "pcfixe"}], [])

    def tearDown(self):
        shutil.rmtree(self.base, ignore_errors=True)

    def _hash(self, path):
        h = hashlib.sha256()
        h.update(Path(path).read_bytes())
        return h.hexdigest()

    def _manifest(self, job_id="annotation_2025-J47_cap_vlm_20260724_120000_a", **overrides):
        data = {
            "job_id": job_id,
            "type": "annotation_photos_batch",
            "status": "done",
            "affaire": "2025-J47",
            "captation": "cap",
            "id_affaire": "2025-J47",
            "id_captation": "cap",
            "dry_run": False,
            "options": list(ACTION_OPTIONS),
            "exit_code": 0,
            "output_verified": True,
            "output_verified_local": True,
            "nas_publish_attempted": True,
            "nas_publish_succeeded": True,
            "nas_publish_error": "",
            "photos_csv_path_used": str(self.photos),
            "photos_csv_sha256_used": self._hash(self.photos),
            "photos_batch_csv_path": str(self.batch),
            "photos_batch_csv_sha256": self._hash(self.batch),
            "photos_batch_csv_local_sha256": self._hash(self.batch),
            "photos_batch_csv_nas_sha256": self._hash(self.batch),
            "finished_at": "2026-07-24T12:00:00",
        }
        data.update(overrides)
        return data

    def _publish_pending_manifest(self, job_id="annotation_2025-J47_cap_vlm_20260724_130000_publish_pending", **overrides):
        data = self._manifest(
            job_id=job_id,
            local_done=True,
            publish_pending=True,
            publish_error_code="NAS_PUBLISH_FAILED",
            nas_publish_succeeded=False,
            nas_publish_error="NAS indisponible",
        )
        data.update(overrides)
        return data

    def _publish_retry_manifest(self, source_manifest, status="done", folder="done", **overrides):
        data = {
            "job_id": f"annotation_publish_retry_2025-J47_cap_20260724_140000_{status}",
            "type": "annotation_photos_batch_publish_retry",
            "status": status,
            "affaire": "2025-J47",
            "captation": "cap",
            "id_affaire": "2025-J47",
            "id_captation": "cap",
            "source_manifest": str(source_manifest),
            "exit_code": 0,
            "publish_pending": False,
            "nas_publish_succeeded": True,
            "finished_at": "2026-07-24T14:00:00",
        }
        if status != "done":
            data["publish_pending"] = True
            data["nas_publish_succeeded"] = False
            data.pop("finished_at", None)
        data.update(overrides)
        return self._write_job(data, folder=folder)

    def _write_job(self, data, folder="done", as_log=False):
        job_id = data["job_id"]
        name = f"{job_id}.manifest.json" if as_log else f"{job_id}.json"
        path = self.jobs / ("logs" if as_log else folder) / name
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return path

    def _write_stamp(self, **overrides):
        data = {
            "batch_id": "b1",
            "job_id": "manual_annotation_migration_2025-J47_cap_20260724_150000_stamp",
            "affaire": "2025-J47",
            "captation": "cap",
            "id_affaire": "2025-J47",
            "id_captation": "cap",
            "photos_csv_sha256_used": self._hash(self.photos),
            "photos_batch_csv_sha256": self._hash(self.batch),
            "last_batch_ts": "2026-07-24 15:00:00",
            "nas_publish_succeeded": True,
            "publish_pending": False,
            "nas_publish_method": "ssh_fallback",
        }
        data.update(overrides)
        path = self.batch.with_name(self.batch.name + ".stamp")
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.utime(self.photos, (2998, 2998))
        os.utime(self.batch, (2998, 2998))
        os.utime(path, (3000, 3000))
        return path

    def _details(self):
        return self.ns["_annotation_job_details"](
            "2025-J47",
            "cap",
            paths=self.paths,
            include_diagnostics=True,
            force_refresh=True,
        )

    def _details_level(self, detail_level):
        return self.ns["_annotation_job_details"](
            "2025-J47",
            "cap",
            paths=self.paths,
            include_diagnostics=True,
            detail_level=detail_level,
            force_refresh=True,
        )

    def _latest(self):
        details, _ = self._details()
        return self.ns["_annotation_latest_completed_batch"](details)

    def _audit(self):
        profile = self.ns["_annotation_file_profile"]
        return {
            "resources": {
                "photos.csv": {"profiles": {"nas": profile(self.photos, csv_expected=True)}, "state": "identique"},
                "photos_batch.csv": {"profiles": {"nas": profile(self.batch, csv_expected=True)}, "state": "identique"},
            },
            "join_audit": self.ns["_annotation_csv_join_audit"](self.photos, self.batch),
        }

    def test_modern_full_manifest_accepted(self):
        self._write_job(self._manifest())
        _, latest = self._latest()
        self.assertEqual(latest["batch_manifest_kind"], "modern")
        self.assertEqual(latest["output_verified_local"], "true")
        self.assertEqual(latest["nas_publish_succeeded"], "true")

    def test_photos_batch_schema_accepts_photo_rel_native_without_nom_fichier_image(self):
        ok, detail = self.ns["_csv_schema_ok"](self.batch, csv_kind="photos_batch")
        self.assertTrue(ok, detail)
        self.assertNotIn("nom_fichier_image", detail)
        profile = self.ns["_annotation_file_profile"](self.batch, csv_expected=True)
        self.assertEqual(profile["schéma"], "cohérent")

    def test_unc_resource_timeout_allows_response_after_100ms_before_configured_limit(self):
        calls = []

        def slow_but_ok(path, timeout_seconds=0.1):
            calls.append(timeout_seconds)
            self.assertGreaterEqual(timeout_seconds, self.ns["ANNOTATION_UNC_RESOURCE_TIMEOUT_SECONDS"])
            return True, ""

        self.ns["_test_path_with_timeout"] = slow_but_ok
        ok, detail = self.ns["_path_accessible_quick"](Path(r"\\nas\Affaires\photos.csv"))

        self.assertTrue(ok)
        self.assertEqual(detail, "")
        self.assertEqual(calls, [self.ns["ANNOTATION_UNC_RESOURCE_TIMEOUT_SECONDS"]])

    def test_unc_resource_true_timeout_is_reported_after_retry(self):
        calls = []

        def always_timeout(path, timeout_seconds=0.1):
            calls.append(timeout_seconds)
            return False, f"Timeout après {int(timeout_seconds * 1000)} ms"

        self.ns["_test_path_with_timeout"] = always_timeout
        ok, detail = self.ns["_path_accessible_quick"](Path(r"\\nas\Affaires\photos.csv"))

        self.assertFalse(ok)
        self.assertIn("Timeout", detail)
        self.assertEqual(
            calls,
            [
                self.ns["ANNOTATION_UNC_RESOURCE_TIMEOUT_SECONDS"],
                self.ns["ANNOTATION_UNC_RESOURCE_RETRY_TIMEOUT_SECONDS"],
            ],
        )

    def test_unc_resource_access_denied_is_not_retried_as_timeout(self):
        calls = []

        def access_denied(path, timeout_seconds=0.1):
            calls.append(timeout_seconds)
            return False, "accès refusé"

        self.ns["_test_path_with_timeout"] = access_denied
        ok, detail = self.ns["_path_accessible_quick"](Path(r"\\nas\Affaires\photos.csv"))

        self.assertFalse(ok)
        self.assertIn("accès refusé", detail)
        self.assertEqual(calls, [self.ns["ANNOTATION_UNC_RESOURCE_TIMEOUT_SECONDS"]])

    def test_csv_row_count_uses_records_not_physical_lines_for_multiline_asr_text(self):
        multiline = self.base / "transcription.csv"
        with multiline.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["photo_rel_native", "dictee_asr_text"], delimiter=";")
            writer.writeheader()
            for idx in range(62):
                writer.writerow({
                    "photo_rel_native": f"photos/P{idx:03}.JPG",
                    "dictee_asr_text": f"ligne A {idx}\nligne B {idx}",
                })
        physical_lines = len(multiline.read_text(encoding="utf-8-sig").splitlines())
        self.assertGreater(physical_lines, 63)
        self.assertEqual(self.ns["_csv_row_count"](multiline), "62")

    def test_join_audit_uses_batch_photo_rel_native_and_rejects_duplicates_or_missing_keys(self):
        audit = self.ns["_annotation_csv_join_audit"](self.photos, self.batch)
        self.assertTrue(audit["ok"], audit)
        self.assertEqual(audit["batch_key"], "photo_rel_native")

        self.batch.write_text(
            "photo_rel_native;batch_status;batch_id;batch_ts\n"
            "p1;OK;b1;2026-07-24T12:00:00\n"
            "p1;OK;b1;2026-07-24T12:00:00\n",
            encoding="utf-8",
        )
        audit = self.ns["_annotation_csv_join_audit"](self.photos, self.batch)
        self.assertFalse(audit["ok"])
        self.assertIn("dupliqué", " | ".join(audit["reasons"]))

        self.batch.write_text(
            "photo_rel_native;batch_status;batch_id;batch_ts\n"
            ";OK;b1;2026-07-24T12:00:00\n",
            encoding="utf-8",
        )
        audit = self.ns["_annotation_csv_join_audit"](self.photos, self.batch)
        self.assertFalse(audit["ok"])
        self.assertIn("photo_rel_native absent", " | ".join(audit["reasons"]))

    def test_old_failed_manifest_does_not_block_newer_verified_output(self):
        failed_path = self._write_job(
            {
                "job_id": "annotation_2025-J47_cap_vlm_20260724_120000_initial_failed",
                "type": "annotation_photos_batch",
                "status": "failed",
                "affaire": "2025-J47",
                "captation": "cap",
                "options": list(ACTION_OPTIONS),
            },
            folder="failed",
        )
        done_path = self._write_job(
            self._manifest(job_id="annotation_2025-J47_cap_vlm_20260724_130000_manual_verified"),
            folder="done",
        )
        os.utime(failed_path, (1000, 1000))
        os.utime(done_path, (2000, 2000))
        os.utime(self.photos, (1500, 1500))
        os.utime(self.batch, (2100, 2100))

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertTrue(preflight["ok"], preflight["reasons"])
        self.assertTrue(state["available"], state)
        self.assertNotIn("job batch plus récent échoué", " | ".join(preflight["reasons"]))

    def test_valid_stamp_without_jobs_manifest_is_completed_batch(self):
        stamp_path = self._write_stamp()
        os.utime(stamp_path, (3000, 3000))
        details, diagnostics = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertEqual(details["initial"]["status"], "completed")
        self.assertEqual(details["initial"]["batch_manifest_kind"], "verified_stamp")
        self.assertEqual(details["initial"]["job_id"], "manual_annotation_migration_2025-J47_cap_20260724_150000_stamp")
        self.assertTrue(preflight["ok"], preflight["reasons"])
        self.assertTrue(state["available"], state)
        self.assertEqual(state["status_code"], "VERIFIED_STAMP")
        self.assertTrue(any(row.get("source") == "verified_stamp" for row in diagnostics["recognized"]))

    def test_old_failed_then_newer_verified_stamp_allows_word(self):
        failed_path = self._write_job(
            {
                "job_id": "annotation_2025-J47_cap_vlm_20260724_120000_initial_failed",
                "type": "annotation_photos_batch",
                "status": "failed",
                "affaire": "2025-J47",
                "captation": "cap",
                "options": list(ACTION_OPTIONS),
            },
            folder="failed",
        )
        stamp_path = self._write_stamp()
        os.utime(failed_path, (1000, 1000))
        os.utime(stamp_path, (3000, 3000))

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertEqual(details["initial"]["batch_manifest_kind"], "verified_stamp")
        self.assertTrue(preflight["ok"], preflight["reasons"])
        self.assertTrue(state["available"], state)

    def test_stamp_with_bad_hash_is_rejected(self):
        self._write_stamp(photos_batch_csv_sha256="0" * 64)
        detail, reasons = self.ns["_annotation_verified_batch_from_stamp"](
            id_affaire="2025-J47",
            id_captation="cap",
            paths=self.paths,
            existing_details={},
        )

        self.assertFalse(detail)
        self.assertIn("hash photos_batch.csv différent", " | ".join(reasons))

    def test_stamp_without_job_id_is_rejected(self):
        self._write_stamp(job_id="")
        detail, reasons = self.ns["_annotation_verified_batch_from_stamp"](
            id_affaire="2025-J47",
            id_captation="cap",
            paths=self.paths,
            existing_details={},
        )

        self.assertFalse(detail)
        self.assertIn("job_id vide", " | ".join(reasons))

    def test_stamp_with_invalid_csv_join_is_rejected(self):
        self.batch.write_text(
            "photo_rel_native;batch_status;batch_id;batch_ts\n"
            "p1;OK;b1;2026-07-24T12:00:00\n"
            "p1;OK;b1;2026-07-24T12:00:00\n",
            encoding="utf-8",
        )
        self._write_stamp(photos_batch_csv_sha256=self._hash(self.batch))
        detail, reasons = self.ns["_annotation_verified_batch_from_stamp"](
            id_affaire="2025-J47",
            id_captation="cap",
            paths=self.paths,
            existing_details={},
        )

        self.assertFalse(detail)
        self.assertIn("dupliqué", " | ".join(reasons))

    def test_valid_stamp_but_newer_publish_pending_blocks_word(self):
        stamp_path = self._write_stamp()
        pending = self._write_job(
            self._publish_pending_manifest(job_id="annotation_2025-J47_cap_vlm_20260724_160000_publish_pending"),
            folder="done",
        )
        os.utime(stamp_path, (3000, 3000))
        os.utime(pending, (4000, 4000))

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertEqual(details["initial"]["publish_pending"], "true")
        self.assertFalse(state["available"])
        self.assertIn("publication NAS en attente", " | ".join(preflight["reasons"]))

    def test_word_button_unlocks_only_with_complete_stamp_proof(self):
        details_without_stamp, _ = self._details()
        preflight_without_stamp = self.ns["_annotation_report_preflight"](self._audit(), details_without_stamp, self.paths)
        state_without_stamp = self.ns["_annotation_report_ui_state"](details_without_stamp, preflight_without_stamp, "provisoire", True)
        self.assertFalse(state_without_stamp["available"])

        self._write_stamp()
        self.ns["_ANNOTATION_JOB_DETAILS_CACHE"].clear()
        details_with_stamp, _ = self._details()
        preflight_with_stamp = self.ns["_annotation_report_preflight"](self._audit(), details_with_stamp, self.paths)
        state_with_stamp = self.ns["_annotation_report_ui_state"](details_with_stamp, preflight_with_stamp, "provisoire", True)

        self.assertTrue(preflight_with_stamp["ok"], preflight_with_stamp["reasons"])
        self.assertTrue(state_with_stamp["available"], state_with_stamp)
        self.assertFalse(state_with_stamp["disable_word"])

    def test_verified_stamp_covers_stale_ui_divergence_when_nas_hashes_match(self):
        self._write_stamp()
        details, _ = self._details()
        profile = self.ns["_annotation_file_profile"]
        audit = {
            "resources": {
                "photos.csv": {
                    "profiles": {"nas": profile(self.photos, csv_expected=True)},
                    "state": "divergence",
                },
                "photos_batch.csv": {
                    "profiles": {"nas": profile(self.batch, csv_expected=True)},
                    "state": "divergence",
                },
            },
            "join_audit": self.ns["_annotation_csv_join_audit"](self.photos, self.batch),
        }
        preflight = self.ns["_annotation_report_preflight"](audit, details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertTrue(preflight["ok"], preflight["reasons"])
        self.assertTrue(state["available"], state)

    def test_modern_publish_guards_refuse_manifest(self):
        cases = [
            ("nas_publish_succeeded", False, "publication NAS échouée"),
            ("nas_publish_attempted", False, "publication NAS non tentée"),
            ("nas_publish_error", "boom", "nas_publish_error non vide"),
            ("output_verified_local", False, "sortie locale non vérifiée"),
            ("dry_run", True, "job dry-run"),
        ]
        for key, value, expected_reason in cases:
            with self.subTest(key=key):
                for child in self.jobs.rglob("*.json"):
                    child.unlink()
                self._write_job(self._manifest(**{key: value}))
                details, diagnostics = self._details()
                self.assertFalse(self.ns["_annotation_latest_completed_batch"](details)[1])
                self.assertIn(expected_reason, " | ".join(item["reason"] for item in diagnostics["ignored"]))

    def test_hash_divergences_refuse_manifest(self):
        cases = [
            ("photos_batch_csv_nas_sha256", "0" * 64, "hash photos_batch.csv NAS différent du manifest"),
            ("photos_csv_sha256_used", "1" * 64, "hash photos.csv source différent du NAS actuel"),
        ]
        for key, value, expected_reason in cases:
            with self.subTest(key=key):
                for child in self.jobs.rglob("*.json"):
                    child.unlink()
                self._write_job(self._manifest(**{key: value}))
                details, diagnostics = self._details()
                self.assertFalse(self.ns["_annotation_latest_completed_batch"](details)[1])
                self.assertIn(expected_reason, " | ".join(item["reason"] for item in diagnostics["ignored"]))

    def test_legacy_accepted_with_warning(self):
        legacy = {
            "job_id": "annotation_2025-J47_cap_vlm_20260724_110000_legacy",
            "type": "annotation_photos_batch",
            "status": "done",
            "affaire": "2025-J47",
            "captation": "cap",
            "options": list(ACTION_OPTIONS),
            "exit_code": 0,
            "photos_batch_csv_path": str(self.batch),
            "finished_at": "2026-07-24T11:00:00",
        }
        self._write_job(legacy)
        key, latest = self._latest()
        self.assertEqual(key, "initial")
        self.assertEqual(latest["batch_manifest_kind"], "legacy")
        preflight = self.ns["_annotation_report_preflight"](self._audit(), {"initial": latest}, self.paths)
        self.assertIn(
            "Batch historique — sortie vérifiée a posteriori, preuve de publication NAS incomplète.",
            preflight["warnings"],
        )

    def test_modern_preferred_over_legacy_and_word_job_contains_publish_state(self):
        self._write_job({
            "job_id": "annotation_2025-J47_cap_vlm_20260724_130000_legacy",
            "type": "annotation_photos_batch",
            "status": "done",
            "affaire": "2025-J47",
            "captation": "cap",
            "options": list(ACTION_OPTIONS),
            "exit_code": 0,
            "photos_batch_csv_path": str(self.batch),
        })
        modern = self._manifest(job_id="annotation_2025-J47_cap_vlm_20260724_120000_modern")
        self._write_job(modern)
        _, latest = self._latest()
        self.assertEqual(latest["job_id"], modern["job_id"])
        preview = self.ns["_photo_report_job_preview"](
            id_affaire="2025-J47",
            id_captation="cap",
            infos_pcfixe=self.infos,
            paths=self.paths,
            audit=self._audit(),
            latest_batch=latest,
            mode="provisoire",
            only_retenue=False,
            dry_run=True,
        )
        self.assertTrue(preview["job"]["batch_nas_publish_succeeded"])
        self.assertTrue(preview["job"]["batch_output_verified"])

    def test_no_batch_preview_is_unavailable_without_valueerror(self):
        preview = self.ns["_photo_report_job_preview"](
            id_affaire="2025-J47",
            id_captation="cap",
            infos_pcfixe=self.infos,
            paths=self.paths,
            audit=self._audit(),
            latest_batch={},
            mode="provisoire",
            only_retenue=False,
            dry_run=True,
        )
        self.assertFalse(preview["available"])
        self.assertEqual(preview["reason_code"], "NO_COMPLETED_BATCH")
        self.assertIsNone(preview["job"])

    def test_report_ui_state_without_batch_disables_word(self):
        state = self.ns["_annotation_report_ui_state"](
            job_details={"initial": {"status": "absent", "job_id": ""}},
            report_preflight={"latest_batch": {}, "reasons": [], "warnings": []},
            report_mode="provisoire",
            gtp_present=True,
        )
        self.assertFalse(state["available"])
        self.assertTrue(state["disable_word"])
        self.assertEqual(state["status_code"], "NO_COMPLETED_BATCH")

    def test_batch_action_state_allows_initial_when_absent_and_blocks_when_queued(self):
        absent = self.ns["_annotation_batch_action_state"]({}, "initial")
        queued = self.ns["_annotation_batch_action_state"]({"initial": {"status": "queued", "job_id": "job123"}}, "initial")
        running = self.ns["_annotation_batch_action_state"]({"initial": {"status": "running", "job_id": "job124"}}, "initial")
        failed = self.ns["_annotation_batch_action_state"]({"initial": {"status": "failed", "job_id": "job125"}}, "initial")
        self.assertTrue(absent["can_submit"])
        self.assertEqual(absent["status_code"], "ABSENT")
        self.assertFalse(queued["can_submit"])
        self.assertEqual(queued["status_code"], "QUEUED")
        self.assertFalse(running["can_submit"])
        self.assertEqual(running["status_code"], "RUNNING")
        self.assertTrue(failed["can_submit"])
        self.assertEqual(failed["status_code"], "FAILED")

    def test_batch_submit_writes_single_json(self):
        result = self.ns["submit_annotation_photos_batch_job"](
            id_affaire="2025-J47",
            id_captation="cap",
            infos_pcfixe=self.infos,
            action_key="initial",
            dry_run=False,
        )
        queued_files = sorted(self.ns["_queue_dir"].glob("*.json"))
        self.assertEqual(result["status"], "queued")
        self.assertEqual(len(queued_files), 1)
        self.assertEqual(queued_files[0].name, f"{result['job_id']}.json")
        runtime_infos = self.paths["pcfixe_unc_trans_dir"] / "_runtime_jobs" / f"infos_projet_runtime_{result['job_id']}.json"
        self.assertTrue(runtime_infos.is_file())

    def test_runtime_infos_localize_unc_paths_to_pcfixe_mirror(self):
        self.infos.write_text(
            json.dumps(
                {
                    "fichier_photos": r"\\192.168.1.20\Affaires\2025-J47\AE_Expert_captations\cap\photos\photos.csv",
                    "fichier_photos_batch": r"\\192.168.1.20\Affaires\2025-J47\AE_Expert_captations\cap\photos\photos_batch.csv",
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        result = self.ns["submit_annotation_photos_batch_job"](
            id_affaire="2025-J47",
            id_captation="cap",
            infos_pcfixe=self.infos,
            action_key="initial",
            dry_run=False,
        )
        queued_job = json.loads((self.ns["_queue_dir"] / f"{result['job_id']}.json").read_text(encoding="utf-8"))
        self.assertEqual(queued_job["fichier_photos"], str(self.paths["pcfixe_photos"]))
        self.assertEqual(queued_job["fichier_photos_batch"], str(self.paths["pcfixe_photos_batch"]))
        runtime_infos = json.loads(
            (self.paths["pcfixe_unc_trans_dir"] / "_runtime_jobs" / f"infos_projet_runtime_{result['job_id']}.json").read_text(encoding="utf-8")
        )
        self.assertEqual(runtime_infos["pcfixe"]["fichier_photos"], str(self.paths["pcfixe_photos"]))
        self.assertEqual(runtime_infos["pcfixe"]["fichier_photos_batch"], str(self.paths["pcfixe_photos_batch"]))

    def test_localize_pcfixe_job_path_accepts_volume1_and_rejects_outside_roots(self):
        localized = self.ns["_annotation_localize_pcfixe_job_path"](
            "/volume1/Affaires/2025-J47/AE_Expert_captations/cap/photos/photos.csv",
            field_name="fichier_photos",
        )
        self.assertEqual(localized, self.paths["pcfixe_photos"])
        with self.assertRaises(ValueError):
            self.ns["_annotation_localize_pcfixe_job_path"](
                r"D:\temp\photos.csv",
                field_name="fichier_photos",
            )

    def test_failed_or_running_newer_blocks_explicitly(self):
        done_path = self._write_job(self._manifest(job_id="annotation_2025-J47_cap_vlm_20260724_120000_done"))
        running_path = self._write_job(
            {
                "job_id": "annotation_2025-J47_cap_vlm_20260724_130000_running",
                "type": "annotation_photos_batch",
                "status": "running",
                "affaire": "2025-J47",
                "captation": "cap",
                "options": list(ACTION_OPTIONS),
            },
            folder="running",
        )
        os.utime(done_path, (1000, 1000))
        os.utime(running_path, (2000, 2000))
        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        self.assertIn("un job batch plus récent est en cours : résultat actuel ambigu", preflight["reasons"])

    def test_done_json_is_associated_with_log_manifest_and_duplicates_by_job_id_are_deduped(self):
        job_id = "annotation_2025-J47_cap_vlm_20260724_120000_assoc"
        self._write_job({"job_id": job_id, "type": "annotation_photos_batch", "status": "done", "affaire": "2025-J47", "captation": "cap", "options": list(ACTION_OPTIONS)}, folder="done")
        self._write_job(self._manifest(job_id=job_id), as_log=True)
        self._write_job(self._manifest(job_id=job_id), folder="done")
        details, _ = self._details_level("detailed")
        completed = [d for d in details.values() if d.get("status") == "completed" and d.get("job_id") == job_id]
        self.assertEqual(len(completed), 1)
        self.assertEqual(completed[0]["manifest_path"], str(self.jobs / "logs" / f"{job_id}.manifest.json"))

    def test_failed_status_is_visible_in_light_scan_and_stale_notice_is_reconciled(self):
        job_id = "annotation_2025-J47_cap_vlm_20260724_140000_failed"
        self._write_job(
            {
                "job_id": job_id,
                "type": "annotation_photos_batch",
                "status": "failed",
                "affaire": "2025-J47",
                "captation": "cap",
                "options": list(ACTION_OPTIONS),
            },
            folder="failed",
        )
        details, _ = self._details_level("light")
        self.assertEqual(details["initial"]["status"], "failed")
        self.assertEqual(details["initial"]["job_id"], job_id)
        notice = self.ns["_annotation_resolve_batch_submission_notice"](
            {
                "id_affaire": "2025-J47",
                "id_captation": "cap",
                "action_key": "initial",
                "action_label": "initial",
                "job_id": job_id,
                "job_path": "queued.json",
                "status": "queued",
            },
            details,
            id_affaire="2025-J47",
            id_captation="cap",
        )
        self.assertEqual(notice["status"], "failed")

    def test_initial_failed_then_weak_rerun_done_allows_word(self):
        self._write_job(
            {
                "job_id": "annotation_2025-J47_cap_vlm_20260724_120000_initial_failed",
                "type": "annotation_photos_batch",
                "status": "failed",
                "affaire": "2025-J47",
                "captation": "cap",
                "options": list(ACTION_OPTIONS),
            },
            folder="failed",
        )
        self._write_job(
            self._manifest(
                job_id="annotation_2025-J47_cap_rerun_weak_20260724_130000_done",
                options=list(WEAK_RERUN_OPTIONS),
            ),
            folder="done",
        )

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertTrue(preflight["ok"], preflight["reasons"])
        self.assertTrue(state["available"], state)
        self.assertEqual(preflight["latest_batch_key"], "weak_rerun")

    def test_initial_failed_then_weak_rerun_no_op_allows_word(self):
        self._write_job(
            {
                "job_id": "annotation_2025-J47_cap_vlm_20260724_120000_initial_failed",
                "type": "annotation_photos_batch",
                "status": "failed",
                "affaire": "2025-J47",
                "captation": "cap",
                "options": list(ACTION_OPTIONS),
            },
            folder="failed",
        )
        self._write_job(
            self._manifest(
                job_id="annotation_2025-J47_cap_rerun_weak_20260724_130000_noop",
                options=list(WEAK_RERUN_OPTIONS),
                output_verified=False,
                output_verified_local=False,
                nas_publish_attempted=False,
                nas_publish_succeeded=False,
                photos_csv_sha256_used="",
                photos_batch_csv_sha256="",
                photos_batch_csv_local_sha256="",
                photos_batch_csv_nas_sha256="",
                no_op=True,
                no_op_reason="NO_WEAK_ROWS",
            ),
            folder="done",
        )

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertTrue(preflight["ok"], preflight["reasons"])
        self.assertTrue(state["available"], state)
        self.assertIn("Aucune ligne WEAK", " | ".join(preflight["warnings"]))
        self.assertEqual(details["weak_rerun"]["no_op"], "true")

    def test_weak_rerun_queued_or_running_blocks_word(self):
        for status, folder, expected_code in (
            ("queued", "queued", "WEAK_RERUN_QUEUED"),
            ("running", "running", "WEAK_RERUN_RUNNING"),
        ):
            with self.subTest(status=status):
                for child in self.jobs.rglob("*.json"):
                    child.unlink()
                self._write_job(
                    {
                        "job_id": f"annotation_2025-J47_cap_rerun_weak_20260724_130000_{status}",
                        "type": "annotation_photos_batch",
                        "status": status,
                        "affaire": "2025-J47",
                        "captation": "cap",
                        "options": list(WEAK_RERUN_OPTIONS),
                    },
                    folder=folder,
                )

                details, _ = self._details()
                preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
                state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

                self.assertFalse(preflight["ok"])
                self.assertFalse(state["available"])
                self.assertEqual(state["status_code"], expected_code)

    def test_weak_rerun_failed_blocks_word_and_allows_retry(self):
        self._write_job(
            {
                "job_id": "annotation_2025-J47_cap_rerun_weak_20260724_130000_failed",
                "type": "annotation_photos_batch",
                "status": "failed",
                "affaire": "2025-J47",
                "captation": "cap",
                "options": list(WEAK_RERUN_OPTIONS),
            },
            folder="failed",
        )

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)
        action_state = self.ns["_annotation_batch_action_state"](details, "weak_rerun")

        self.assertFalse(preflight["ok"])
        self.assertEqual(state["status_code"], "WEAK_RERUN_FAILED")
        self.assertTrue(action_state["can_submit"])

    def test_stale_session_queued_is_reconciled_with_real_failed_or_done(self):
        for real_status, folder, expected_status in (
            ("failed", "failed", "failed"),
            ("done", "done", "completed"),
        ):
            with self.subTest(real_status=real_status):
                for child in self.jobs.rglob("*.json"):
                    child.unlink()
                job_id = f"annotation_2025-J47_cap_rerun_weak_20260724_130000_{real_status}"
                manifest = (
                    self._manifest(job_id=job_id, options=list(WEAK_RERUN_OPTIONS))
                    if real_status == "done"
                    else {
                        "job_id": job_id,
                        "type": "annotation_photos_batch",
                        "status": "failed",
                        "affaire": "2025-J47",
                        "captation": "cap",
                        "options": list(WEAK_RERUN_OPTIONS),
                    }
                )
                self._write_job(manifest, folder=folder)
                details, _ = self._details()
                notice = self.ns["_annotation_resolve_batch_submission_notice"](
                    {
                        "id_affaire": "2025-J47",
                        "id_captation": "cap",
                        "action_key": "weak_rerun",
                        "action_label": "weak",
                        "job_id": job_id,
                        "job_path": "queued.json",
                        "status": "queued",
                    },
                    details,
                    id_affaire="2025-J47",
                    id_captation="cap",
                )
                self.assertEqual(details["weak_rerun"]["status"], expected_status)
                self.assertEqual(notice["status"], expected_status)

    def test_local_done_publish_pending_is_not_failed_and_blocks_word(self):
        pending_path = self._write_job(self._publish_pending_manifest(), folder="done")

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertEqual(details["initial"]["status"], "completed")
        self.assertEqual(details["initial"]["publish_pending"], "true")
        self.assertEqual(details["initial"]["nas_publish_succeeded"], "false")
        self.assertIn("Batch local validé", details["initial"]["verification_note"])
        self.assertFalse(preflight["ok"])
        self.assertIn("publication NAS en attente", " | ".join(preflight["reasons"]))
        self.assertEqual(state["status_code"], "BATCH_LOCAL_DONE_PUBLISH_PENDING")
        self.assertEqual(details["initial"]["manifest_path"], str(pending_path))

    def test_publish_retry_queued_running_failed_block_word(self):
        pending_path = self._write_job(self._publish_pending_manifest(), folder="done")
        for status, folder, expected_code in (
            ("queued", "queued", "PUBLISH_RETRY_QUEUED"),
            ("running", "running", "PUBLISH_RETRY_RUNNING"),
            ("failed", "failed", "PUBLISH_RETRY_FAILED"),
        ):
            with self.subTest(status=status):
                for child in self.jobs.rglob("annotation_publish_retry_*.json"):
                    child.unlink()
                self._publish_retry_manifest(pending_path, status=status, folder=folder)

                details, _ = self._details()
                preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
                state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

                self.assertFalse(state["available"])
                self.assertEqual(state["status_code"], expected_code)

    def test_word_allowed_after_publish_retry_done(self):
        pending_path = self._write_job(self._publish_pending_manifest(), folder="done")
        self._publish_retry_manifest(pending_path, status="done", folder="done")

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertEqual(details["publish_retry"]["status"], "completed")
        self.assertEqual(details["publish_retry"]["publish_pending"], "false")
        self.assertEqual(details["publish_retry"]["nas_publish_succeeded"], "true")
        self.assertTrue(preflight["ok"], preflight["reasons"])
        self.assertTrue(state["available"], state)
        self.assertEqual(preflight["latest_batch_key"], "publish_retry")

    def test_initial_failed_replaced_by_local_validated_pending_publish(self):
        self._write_job(
            {
                "job_id": "annotation_2025-J47_cap_vlm_20260724_120000_initial_failed",
                "type": "annotation_photos_batch",
                "status": "failed",
                "affaire": "2025-J47",
                "captation": "cap",
                "options": list(ACTION_OPTIONS),
            },
            folder="failed",
        )
        self._write_job(
            self._publish_pending_manifest(job_id="annotation_2025-J47_cap_vlm_20260724_130000_publish_pending"),
            folder="done",
        )

        details, _ = self._details()
        preflight = self.ns["_annotation_report_preflight"](self._audit(), details, self.paths)
        state = self.ns["_annotation_report_ui_state"](details, preflight, "provisoire", True)

        self.assertEqual(details["initial"]["status"], "completed")
        self.assertEqual(state["status_code"], "BATCH_LOCAL_DONE_PUBLISH_PENDING")
        self.assertNotIn("job batch plus récent échoué", " | ".join(preflight["reasons"]))

    def test_publish_retry_submit_writes_source_manifest_job(self):
        source_manifest = self.jobs / "done" / "source.manifest.json"
        source_manifest.write_text("{}", encoding="utf-8")
        result = self.ns["submit_annotation_photos_batch_publish_retry_job"](
            id_affaire="2025-J47",
            id_captation="cap",
            source_manifest=source_manifest,
            dry_run=False,
        )

        queued_job = json.loads((self.ns["_queue_dir"] / f"{result['job_id']}.json").read_text(encoding="utf-8"))
        self.assertEqual(result["status"], "queued")
        self.assertEqual(queued_job["type"], "annotation_photos_batch_publish_retry")
        self.assertEqual(queued_job["source_manifest"], str(source_manifest))

    def test_light_render_does_not_scan_logs(self):
        job_id = "annotation_2025-J47_cap_vlm_20260724_120000_light"
        self._write_job(self._manifest(job_id=job_id), folder="done")
        _, diagnostics = self._details_level("light")
        perf_blocks = diagnostics.get("perf_blocks", [])
        self.assertFalse(any(block.get("block") == "scan logs" for block in perf_blocks))

    def test_vpn_root_stops_after_first_accessible_host(self):
        call_order = []

        def fake_timeout(path):
            raw = str(path)
            call_order.append(raw)
            if raw.startswith(r"\\10.0.1.10\Affaires\_jobs"):
                return True, ""
            if raw.startswith(r"\\192.168.0.120\Affaires\_jobs") or raw.startswith(r"\\192.168.0.155\Affaires\_jobs"):
                raise AssertionError("les alias LAN ne doivent pas être testés quand la racine VPN répond")
            return False, "blocked"

        self.ns["_test_path_with_timeout"] = fake_timeout
        self.ns["_unc_host"] = lambda path: "10.0.1.10" if str(path).startswith("\\\\10.0.1.10\\") else ("192.168.0.120" if str(path).startswith("\\\\192.168.0.120\\") else ("192.168.0.155" if str(path).startswith("\\\\192.168.0.155\\") else ""))
        roots, probes = self.original_annotation_jobs_roots(detail_level="light")
        self.assertEqual(len(roots), 1)
        self.assertTrue(str(roots[0]["root"]).startswith(r"\\10.0.1.10\Affaires\_jobs"))
        self.assertTrue(any(raw.startswith(r"\\10.0.1.10\Affaires\_jobs") for raw in call_order))
        self.assertFalse(any(raw.startswith(r"\\192.168.0.120\Affaires\_jobs") for raw in call_order))
        self.assertFalse(any(raw.startswith(r"\\192.168.0.155\Affaires\_jobs") for raw in call_order))
        self.assertEqual(len(probes), 1)

    def test_filename_filter_limits_irrelevant_json_reads(self):
        irrelevant = self.jobs / "done" / "annotation_2024-Z99_other_20260724_120000_noise.json"
        irrelevant.write_text(json.dumps(self._manifest(job_id="annotation_2024-Z99_other_20260724_120000_noise", affaire="2024-Z99", captation="other")), encoding="utf-8")
        self._write_job(self._manifest(job_id="annotation_2025-J47_cap_vlm_20260724_120000_match"), folder="done")
        _, diagnostics = self._details_level("light")
        self.assertEqual(diagnostics["timings"]["json_files_seen"], 1)


if __name__ == "__main__":
    unittest.main()
