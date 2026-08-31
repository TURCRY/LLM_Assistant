"""Tests ciblés : écriture du backend LLM du batch photos dans config_llm.json.

Source d'autorité : config_llm.json["llm_backend"] (lu par batch_all_photos_pcfixe.py).
infos_projet.json["llm_backend"] est conservé uniquement pour compatibilité/affichage UI.
"""

from __future__ import annotations

import ast
import hashlib
import json
import os
import shutil
import unittest
import uuid
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


_FUNCTION_NAMES = {
    "_normalize_cr_llm_backend",
    "_patch_captation_llm_backend",
    "load_json",
    "_atomic_write_json",
    "_sha256_file",
    "_annotation_batch_resource_plan",
    "_annotation_nas_required_resources",
    "_annotation_pcfixe_required_resources",
    "_annotation_laptop_config_resources",
}


def _load_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    module = ast.Module(
        body=[node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in _FUNCTION_NAMES],
        type_ignores=[],
    )
    ast.fix_missing_locations(module)
    ns = {
        "Path": Path,
        "json": json,
        "os": os,
        "hashlib": hashlib,
        "CR_LLM_BACKENDS": {"openai", "local"},
    }
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns


def _write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _default_config(backend: str) -> dict:
    return {
        "openai_api_key": "API_OPENAI_PLACEHOLDER",
        "model": "gpt-4o-mini",
        "asr_model": "gpt-4o-mini-transcribe",
        "temperature": 0.25,
        "max_tokens": 120,
        "retries": 3,
        "timeout": 180,
        "llm_backend": backend,
        "local_llm": {
            "base_url": "",
            "api_key": "LOCAL_LLM_API_KEY_PLACEHOLDER",
            "model": "Qwen_2_5_14B",
            "timeout": 30,
        },
        "batch": {
            "max_tokens": {"libelle": 300, "commentaire": 400},
            "temperature": {"libelle": 0.1, "commentaire": 0.1},
        },
    }


def _sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


class AnnotationLlmBackendTests(unittest.TestCase):
    def setUp(self):
        temp_root = Path(os.environ.get("LLM_ASSISTANT_TEST_TMP") or r"C:\CodexWorkspace\.tmp_llm_assistant_tests")
        if not temp_root.parent.exists():
            temp_root = APP_PATH.parent / "temp"
        temp_root.mkdir(exist_ok=True)
        self.base = temp_root / f"llm_backend_{uuid.uuid4().hex}"
        self.base.mkdir(parents=True, exist_ok=False)
        self.ns = _load_functions()

    def tearDown(self):
        shutil.rmtree(self.base, ignore_errors=True)

    def _config_path(self) -> Path:
        return self.base / "trans" / "config_llm.json"

    def test_patch_backend_openai_writes_config_llm_and_syncs_infos(self):
        patch = self.ns["_patch_captation_llm_backend"]
        config_path = self._config_path()
        infos_path = config_path.with_name("infos_projet.json")
        _write_json(config_path, _default_config(backend="local"))
        _write_json(infos_path, {"id_affaire": "2025-J47", "llm_backend": "local"})

        result = patch(config_llm_path=config_path, backend="openai", infos_path=infos_path)

        self.assertEqual(result["backend"], "openai")
        self.assertEqual(result["config_llm_old"], "local")
        self.assertTrue(result["config_llm_written"])
        self.assertTrue(result["infos_written"])
        self.assertEqual(result["infos_sync_error"], "")

        config = json.loads(config_path.read_text(encoding="utf-8"))
        self.assertEqual(config["llm_backend"], "openai")
        # préservation des autres champs de config_llm.json
        self.assertEqual(config["model"], "gpt-4o-mini")
        self.assertEqual(config["asr_model"], "gpt-4o-mini-transcribe")
        self.assertEqual(config["temperature"], 0.25)
        self.assertEqual(config["max_tokens"], 120)
        self.assertEqual(config["retries"], 3)
        self.assertEqual(config["timeout"], 180)
        self.assertEqual(config["local_llm"]["model"], "Qwen_2_5_14B")
        self.assertEqual(config["local_llm"]["base_url"], "")
        self.assertEqual(config["local_llm"]["timeout"], 30)
        self.assertEqual(config["batch"]["max_tokens"]["commentaire"], 400)

        infos = json.loads(infos_path.read_text(encoding="utf-8"))
        self.assertEqual(infos["llm_backend"], "openai")
        self.assertEqual(infos["id_affaire"], "2025-J47")

    def test_patch_backend_local_writes_config_llm(self):
        patch = self.ns["_patch_captation_llm_backend"]
        config_path = self._config_path()
        _write_json(config_path, _default_config(backend="openai"))

        result = patch(config_llm_path=config_path, backend="local")

        self.assertEqual(result["config_llm_old"], "openai")
        self.assertTrue(result["config_llm_written"])
        config = json.loads(config_path.read_text(encoding="utf-8"))
        self.assertEqual(config["llm_backend"], "local")
        self.assertEqual(config["local_llm"]["model"], "Qwen_2_5_14B")

    def test_patch_backend_no_rewrite_when_unchanged(self):
        patch = self.ns["_patch_captation_llm_backend"]
        config_path = self._config_path()
        infos_path = config_path.with_name("infos_projet.json")
        _write_json(config_path, _default_config(backend="openai"))
        _write_json(infos_path, {"llm_backend": "openai"})

        config_before = config_path.read_bytes()
        infos_before = infos_path.read_bytes()

        result = patch(config_llm_path=config_path, backend="openai", infos_path=infos_path)

        self.assertFalse(result["config_llm_written"])
        self.assertFalse(result["infos_written"])
        self.assertEqual(config_path.read_bytes(), config_before)
        self.assertEqual(infos_path.read_bytes(), infos_before)

    def test_patch_backend_normalizes_case_and_rejects_invalid(self):
        patch = self.ns["_patch_captation_llm_backend"]
        config_path = self._config_path()
        _write_json(config_path, _default_config(backend="local"))

        result = patch(config_llm_path=config_path, backend="OPENAI")
        self.assertEqual(result["backend"], "openai")
        config = json.loads(config_path.read_text(encoding="utf-8"))
        self.assertEqual(config["llm_backend"], "openai")

        with self.assertRaises(ValueError):
            patch(config_llm_path=config_path, backend="bogus")

    def test_patch_backend_skips_infos_when_missing(self):
        patch = self.ns["_patch_captation_llm_backend"]
        config_path = self._config_path()
        _write_json(config_path, _default_config(backend="local"))
        missing_infos = self.base / "absent" / "infos_projet.json"

        result = patch(config_llm_path=config_path, backend="openai", infos_path=missing_infos)

        self.assertTrue(result["config_llm_written"])
        self.assertFalse(result["infos_written"])
        self.assertIn("absent", result["infos_sync_error"])
        self.assertFalse(missing_infos.exists())

    def test_required_resources_hash_matches_patched_config(self):
        patch = self.ns["_patch_captation_llm_backend"]
        plan = self.ns["_annotation_batch_resource_plan"]
        sha256_file = self.ns["_sha256_file"]

        nas = self.base / "nas"
        pc = self.base / "pc"
        pc_unc = self.base / "pc_unc"
        laptop = self.base / "laptop"
        for directory in (nas, pc, pc_unc, laptop):
            directory.mkdir(parents=True, exist_ok=True)
        paths = {
            "nas_trans_dir": nas,
            "nas_infos": nas / "infos_projet.json",
            "nas_photos": nas / "photos.csv",
            "nas_photos_batch": nas / "photos_batch.csv",
            "laptop_trans_dir": laptop,
            "pcfixe_trans_dir": pc,
            "pcfixe_infos": pc / "infos_projet.json",
            "pcfixe_photos": pc / "photos.csv",
            "pcfixe_photos_batch": pc / "photos_batch.csv",
            "pcfixe_unc_trans_dir": pc_unc,
            "pcfixe_unc_infos": pc_unc / "infos_projet.json",
            "pcfixe_unc_photos": pc_unc / "photos.csv",
            "pcfixe_unc_photos_batch": pc_unc / "photos_batch.csv",
        }
        for value in paths.values():
            if value.suffix:
                value.write_text("fixture", encoding="utf-8")

        config_path = nas / "config_llm.json"
        _write_json(config_path, _default_config(backend="local"))

        result = patch(config_llm_path=config_path, backend="openai", infos_path=paths["nas_infos"])
        self.assertTrue(result["config_llm_written"])

        resources = plan(paths)
        config_row = next(row for row in resources if row["logical_name"] == "config_llm.json")
        self.assertEqual(config_row["nas_source"], str(config_path))
        self.assertEqual(config_row["sha256"], sha256_file(config_path))
        self.assertEqual(config_row["sha256"], _sha256_bytes(config_path.read_bytes()))
        referenced = json.loads(config_path.read_text(encoding="utf-8"))
        self.assertEqual(referenced["llm_backend"], "openai")


if __name__ == "__main__":
    unittest.main()
