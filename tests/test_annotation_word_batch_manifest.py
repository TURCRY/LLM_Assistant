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


def _load_annotation_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    names = {
        "_valid_file_path",
        "_infos_declared_path",
        "_annotation_cache_prune",
        "_annotation_perf_begin",
        "_annotation_perf_end",
        "_annotation_list_json_candidates",
        "_annotation_jobs_roots",
        "_annotation_job_details",
        "_annotation_latest_completed_batch",
        "_annotation_has_newer_blocking_batch",
        "_annotation_report_preflight",
        "_annotation_report_ui_state",
        "_annotation_batch_action_state",
        "_photo_report_job_preview",
        "submit_annotation_photos_batch_job",
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
            }
        },
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
        "NAS_AFFAIRES_ROOT": Path(r"C:\nas"),
        "AFFAIRES_ROOT": r"C:\Affaires",
        "PHOTO_REPORT_JOB_SUPPORTED": True,
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

    def file_profile(path, *, csv_expected=False):
        path = Path(path)
        return {
            "présent": "oui" if path.is_file() else "non",
            "sha256": sha256_file(path) if path.is_file() else "",
            "mtime": str(path.stat().st_mtime) if path.exists() else "0",
            "schéma": "cohérent" if path.is_file() else "absent",
        }

    def csv_schema_ok(path, required_columns=("photo_rel_native", "nom_fichier_image")):
        with Path(path).open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            fields = set(reader.fieldnames or [])
        missing = [col for col in required_columns if col not in fields]
        return (not missing), ", ".join(missing)

    ns.update(
        {
            "load_json": load_json,
            "traceback": traceback,
            "_path_accessible_quick": lambda path: (Path(path).exists(), ""),
            "_test_path_with_timeout": lambda path: (Path(path).exists(), ""),
            "_unc_host": lambda path: "",
            "_safe_job_token": lambda value: re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "").strip()).strip("_") or "na",
            "_annotation_file_profile": file_profile,
            "_csv_schema_ok": csv_schema_ok,
            "_job_failure_details": lambda job_id: {},
            "get_pcfixe_jobs_queued_dir": lambda: ns["_queue_dir"],
            "_assert_annotation_pcfixe_resources_ready": lambda *args, **kwargs: {},
            "preflight_pcfixe_target_dir": lambda *args, **kwargs: {},
            "_annotation_canonical_paths": lambda affaire, captation: ns["_paths"],
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
        self.photos.write_text("photo_rel_native,nom_fichier_image\np1,P1.JPG\n", encoding="utf-8")
        self.batch.write_text("photo_rel_native,nom_fichier_image\np1,P1.JPG\n", encoding="utf-8")
        self.infos.write_text("{}", encoding="utf-8")
        self.paths = {
            "nas_photos": self.photos,
            "nas_photos_batch": self.batch,
            "nas_infos": self.infos,
            "nas_report_dir": self.base / "nas" / "compte_rendu_LLM",
            "nas_report_out_dir": self.base / "nas" / "compte_rendu_LLM" / "out",
            "pcfixe_report_dir": self.base / "pc" / "compte_rendu_LLM",
            "pcfixe_photos": self.base / "pc" / "photos.csv",
            "pcfixe_photos_batch": self.base / "pc" / "photos_batch.csv",
        }
        self.ns["_paths"] = self.paths
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

    def _write_job(self, data, folder="done", as_log=False):
        job_id = data["job_id"]
        name = f"{job_id}.manifest.json" if as_log else f"{job_id}.json"
        path = self.jobs / ("logs" if as_log else folder) / name
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
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
            }
        }

    def test_modern_full_manifest_accepted(self):
        self._write_job(self._manifest())
        _, latest = self._latest()
        self.assertEqual(latest["batch_manifest_kind"], "modern")
        self.assertEqual(latest["output_verified_local"], "true")
        self.assertEqual(latest["nas_publish_succeeded"], "true")

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
        self.assertTrue(absent["can_submit"])
        self.assertEqual(absent["status_code"], "ABSENT")
        self.assertFalse(queued["can_submit"])
        self.assertEqual(queued["status_code"], "QUEUED")

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
