from __future__ import annotations

import ast
import hashlib
import json
import os
import shutil
import tempfile
import unittest
import uuid
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
FUNCTIONS = {
    "prepare_audit_reunion_quality_infos",
    "submit_audit_reunion_quality_job",
}


def _load_namespace(pc_root: Path, nas_root: Path, queue: Path):
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    body = [
        node for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS
    ]
    namespace = {
        "Path": Path,
        "hashlib": hashlib,
        "json": json,
        "os": os,
        "shutil": shutil,
        "uuid": uuid,
        "datetime": datetime,
        "NAS_AFFAIRES_ROOT": nas_root,
        "PCFIXE_AFFAIRES_ROOT": Path(r"C:\Affaires"),
        "get_pcfixe_affaires_root": lambda: pc_root,
        "get_pcfixe_jobs_queued_dir": lambda: queue,
        "preflight_pcfixe_target_dir": lambda path: Path(path).mkdir(parents=True, exist_ok=True),
        "_nas_affaires_path_to_volume1": lambda value: str(value).replace(str(nas_root), "/volume1/Affaires").replace("\\", "/"),
        "_is_volume1_affaires_path": lambda value: str(value).startswith("/volume1/Affaires/"),
    }
    exec(compile(ast.Module(body=body, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


class AuditReunionQualityPreflightTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.pc_root = root / "pcfixe" / "Affaires"
        self.nas_root = root / "nas" / "Affaires"
        self.queue = self.pc_root / "_jobs" / "queued"
        self.rel = Path("2026-A60/AF_Expert_ASR/transcriptions/accedit-2026-10-01/infos_projet.json")
        self.ns = _load_namespace(self.pc_root, self.nas_root, self.queue)

    def tearDown(self):
        self.temp.cleanup()

    def _call(self):
        return self.ns["prepare_audit_reunion_quality_infos"](
            infos_path=self.nas_root / self.rel,
            id_affaire="2026-A60",
            id_captation="accedit-2026-10-01",
        )

    @staticmethod
    def _write(path: Path, payload: bytes):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload)

    def test_present_pcfixe_is_used_without_nas_copy(self):
        self._write(self.pc_root / self.rel, b'{"source":"pcfixe"}')
        self._write(self.nas_root / self.rel, b'{"source":"nas"}')

        result = self._call()

        self.assertEqual(result["source"], "pcfixe")
        self.assertFalse(result["synchronized"])
        self.assertEqual((self.pc_root / self.rel).read_bytes(), b'{"source":"pcfixe"}')

    def test_absent_pcfixe_present_nas_is_synchronized(self):
        payload = b'{"id_affaire":"2026-A60"}'
        self._write(self.nas_root / self.rel, payload)

        result = self._call()

        self.assertEqual(result["source"], "nas")
        self.assertTrue(result["synchronized"])
        self.assertEqual((self.pc_root / self.rel).read_bytes(), payload)

    def test_absent_both_raises_before_queue_deposit(self):
        with self.assertRaisesRegex(FileNotFoundError, "PC fixe puis NAS"):
            self._call()
        self.assertFalse(self.queue.exists())

    def test_hashes_are_identical_after_synchronization(self):
        payload = b'{"id_affaire":"2026-A60","accent":"reunion"}'
        self._write(self.nas_root / self.rel, payload)

        result = self._call()

        expected = hashlib.sha256(payload).hexdigest()
        self.assertEqual(result["sha256_source"], expected)
        self.assertEqual(result["sha256_target"], expected)

    def test_job_receives_explicit_pcfixe_runtime_path(self):
        self._write(self.nas_root / self.rel, b"{}")
        run_dir = self.nas_root / "2026-A60/BE_Traitement_captations/accedit-2026-10-01/compte_rendu_LLM/out/job_test"
        run_dir.mkdir(parents=True)

        result = self.ns["submit_audit_reunion_quality_job"](
            job_dir=run_dir,
            infos_path=self.nas_root / self.rel,
            id_affaire="2026-A60",
            id_captation="accedit-2026-10-01",
        )

        self.assertEqual(
            result["job"]["infos_projet"],
            r"C:\Affaires\2026-A60\AF_Expert_ASR\transcriptions\accedit-2026-10-01\infos_projet.json",
        )
        queued = json.loads(Path(result["job_path"]).read_text(encoding="utf-8"))
        self.assertEqual(queued["infos_projet"], result["job"]["infos_projet"])


if __name__ == "__main__":
    unittest.main()
