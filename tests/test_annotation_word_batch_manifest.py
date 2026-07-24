from __future__ import annotations

import ast
import csv
import hashlib
import json
import os
import re
import shutil
import unittest
import uuid
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
ACTION_OPTIONS = ["--reset-vlm", "1", "--vlm-strict", "1"]


def _load_annotation_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    names = {
        "_annotation_job_details",
        "_annotation_latest_completed_batch",
        "_annotation_has_newer_blocking_batch",
        "_annotation_report_preflight",
        "_photo_report_job_preview",
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
        "PHOTO_REPORT_JOB_SUPPORTED": True,
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
            "_path_accessible_quick": lambda path: (Path(path).exists(), ""),
            "_test_path_with_timeout": lambda path: (Path(path).exists(), ""),
            "_unc_host": lambda path: "",
            "_safe_job_token": lambda value: re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "").strip()).strip("_") or "na",
            "_annotation_file_profile": file_profile,
            "_csv_schema_ok": csv_schema_ok,
            "_job_failure_details": lambda job_id: {},
            "get_pcfixe_jobs_queued_dir": lambda: ns["_queue_dir"],
            "_annotation_canonical_paths": lambda affaire, captation: ns["_paths"],
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
        self.ns["_annotation_jobs_roots"] = lambda: ([{"root": str(self.jobs), "label": "test", "source": "pcfixe"}], [])

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
        return self.ns["_annotation_job_details"]("2025-J47", "cap", paths=self.paths, include_diagnostics=True)

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
        details, _ = self._details()
        completed = [d for d in details.values() if d.get("status") == "completed" and d.get("job_id") == job_id]
        self.assertEqual(len(completed), 1)
        self.assertEqual(completed[0]["manifest_path"], str(self.jobs / "logs" / f"{job_id}.manifest.json"))


if __name__ == "__main__":
    unittest.main()
