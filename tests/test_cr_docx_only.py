from __future__ import annotations

import ast
import json
import os
import shutil
import unittest
import uuid
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_functions(nas_root: Path, queue_dir: Path):
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    names = {
        "_quote_nas_arg",
        "_compte_rendu_nas_infos_path",
        "_compte_rendu_nas_out_dir",
        "_nas_affaires_path_to_volume1",
        "_volume1_affaires_path_to_unc",
        "_is_volume1_affaires_path",
        "_compte_rendu_existing_run_basename",
        "_validate_compte_rendu_existing_run_contract",
        "_compte_rendu_existing_run_contract",
        "_compte_rendu_nas_command",
        "discover_compte_rendu_run_dirs",
        "inspect_compte_rendu_docx_run",
        "discover_compte_rendu_docx_runs",
        "submit_compte_rendu_job",
    }
    module = ast.Module(
        body=[node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names],
        type_ignores=[],
    )
    ast.fix_missing_locations(module)

    def load_json(path, default):
        try:
            return json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception:
            return default

    namespace = {
        "Path": Path,
        "re": __import__("re"),
        "os": os,
        "json": json,
        "uuid": uuid,
        "datetime": datetime,
        "NAS_AFFAIRES_ROOT": nas_root,
        "load_json": load_json,
        "_cr_existing_routing": lambda infos: {},
        "get_pcfixe_jobs_queued_dir": lambda: queue_dir,
        "_patch_cr_llm_routing_copies": lambda **kwargs: [],
        "preflight_pcfixe_target_dir": lambda path: None,
    }
    exec(compile(module, str(APP_PATH), "exec"), namespace)
    return namespace


class CompteRenduDocxOnlyTests(unittest.TestCase):
    def setUp(self):
        temp_root = Path(os.environ.get("LLM_ASSISTANT_TEST_TMP") or r"C:\CodexWorkspace\.tmp_llm_assistant_tests")
        temp_root.mkdir(exist_ok=True)
        self.base = temp_root / f"cr_docx_only_{uuid.uuid4().hex}"
        self.base.mkdir()
        self.nas_root = self.base / "Affaires"
        self.queue = self.base / "queue"
        self.queue.mkdir()
        self.ns = _load_functions(self.nas_root, self.queue)
        self.affaire = "2026-A60"
        self.captation = "accedit-2026-06-17"
        self.out_dir = (
            self.nas_root
            / self.affaire
            / "BE_Traitement_captations"
            / self.captation
            / "compte_rendu_LLM"
            / "out"
        )
        self.run_name = "job_20260805_141904_811447936_29008"
        self.run_dir = self.out_dir / self.run_name
        self.run_dir.mkdir(parents=True)
        (self.run_dir / "global_final.json").write_text("{}", encoding="utf-8")
        (self.run_dir / "global_by_sujet.json").write_text("{}", encoding="utf-8")
        (self.run_dir / "audit_reunion_quality.json").write_text("{}", encoding="utf-8")
        (self.run_dir / "pipeline_qa_status.json").write_text('{"ok": true}', encoding="utf-8")
        self.infos = self.nas_root / self.affaire / "AF_Expert_ASR" / "transcriptions" / self.captation / "infos_projet.json"
        self.infos.parent.mkdir(parents=True)
        self.infos.write_text("{}", encoding="utf-8")
        self.nas_run = (
            f"/volume1/Affaires/{self.affaire}/BE_Traitement_captations/"
            f"{self.captation}/compte_rendu_LLM/out/{self.run_name}"
        )

    def tearDown(self):
        shutil.rmtree(self.base, ignore_errors=True)

    def test_absolute_nas_run_is_normalized_in_manifest_and_command(self):
        profiles = self.ns["discover_compte_rendu_docx_runs"](
            self.out_dir,
            id_affaire=self.affaire,
            id_captation=self.captation,
        )
        self.assertEqual(profiles[0]["nas_path"], self.nas_run)
        self.assertEqual(profiles[0]["existing_run"], self.run_name)
        self.assertTrue(profiles[0]["audited"])

        result = self.ns["submit_compte_rendu_job"](
            infos_path=self.infos,
            id_affaire=self.affaire,
            id_captation=self.captation,
            docx_only=True,
            existing_run=profiles[0]["nas_path"],
            mirror_pc=True,
            dry_run=False,
        )

        manifest = json.loads(Path(result["job_path"]).read_text(encoding="utf-8"))
        self.assertTrue(manifest["docx_only"])
        self.assertEqual(manifest["existing_run"], self.run_name)
        self.assertEqual(manifest["existing_run_path"], self.nas_run)
        self.assertIn("--docx-only", result["nas_command"])
        self.assertIn("--existing-run", result["nas_command"])
        self.assertIn(self.run_name, result["nas_command"])
        self.assertNotIn(f"--existing-run' '{self.nas_run}", result["nas_command"])
        self.assertIn("--mirror-pc", result["nas_command"])

    def test_basename_run_is_accepted_and_manifest_has_no_path(self):
        result = self.ns["submit_compte_rendu_job"](
            infos_path=self.infos,
            id_affaire=self.affaire,
            id_captation=self.captation,
            docx_only=True,
            existing_run=self.run_name,
            mirror_pc=True,
            dry_run=False,
        )

        manifest = json.loads(Path(result["job_path"]).read_text(encoding="utf-8"))
        self.assertEqual(manifest["existing_run"], self.run_name)
        self.assertEqual(manifest["existing_run_path"], self.nas_run)
        self.assertIn("--existing-run", result["nas_command"])
        self.assertIn(self.run_name, result["nas_command"])
        self.assertNotIn(self.nas_run, result["nas_command"])

    def test_docx_only_without_run_is_blocked_before_manifest_write(self):
        with self.assertRaisesRegex(ValueError, "existing_run valide"):
            self.ns["submit_compte_rendu_job"](
                infos_path=self.infos,
                id_affaire=self.affaire,
                id_captation=self.captation,
                docx_only=True,
                existing_run="",
                dry_run=False,
            )
        self.assertEqual(list(self.queue.glob("*.json")), [])

    def test_contract_rejects_absolute_path_and_invalid_basename(self):
        with self.assertRaisesRegex(ValueError, "pas un chemin"):
            self.ns["_validate_compte_rendu_existing_run_contract"](self.nas_run)
        with self.assertRaisesRegex(ValueError, "commencer par job_"):
            self.ns["_validate_compte_rendu_existing_run_contract"]("run_20260805")
        with self.assertRaisesRegex(ValueError, "invalide"):
            self.ns["_validate_compte_rendu_existing_run_contract"]("job_bad value")
        with self.assertRaisesRegex(ValueError, "pas un chemin"):
            self.ns["_validate_compte_rendu_existing_run_contract"](self.run_name + "\\nested")

    def test_run_from_another_captation_is_rejected(self):
        wrong_run = self.nas_run.replace(self.captation, "accedit-2026-05-21")
        profile = self.ns["inspect_compte_rendu_docx_run"](
            wrong_run,
            id_affaire=self.affaire,
            id_captation=self.captation,
        )
        self.assertFalse(profile["valid"])
        with self.assertRaisesRegex(ValueError, "Run DOCX invalide"):
            self.ns["submit_compte_rendu_job"](
                infos_path=self.infos,
                id_affaire=self.affaire,
                id_captation=self.captation,
                docx_only=True,
                existing_run=wrong_run,
                dry_run=True,
            )

    def test_run_from_another_affaire_is_rejected(self):
        wrong_run = self.nas_run.replace(self.affaire, "2026-B99")
        profile = self.ns["inspect_compte_rendu_docx_run"](
            wrong_run,
            id_affaire=self.affaire,
            id_captation=self.captation,
        )
        self.assertFalse(profile["valid"])
        with self.assertRaisesRegex(ValueError, "Run DOCX invalide"):
            self.ns["submit_compte_rendu_job"](
                infos_path=self.infos,
                id_affaire=self.affaire,
                id_captation=self.captation,
                docx_only=True,
                existing_run=wrong_run,
                dry_run=True,
            )

    def test_real_case_contract_value(self):
        contract = self.ns["_compte_rendu_existing_run_contract"](self.nas_run)
        self.assertEqual(contract, "job_20260805_141904_811447936_29008")

    def test_ui_button_is_disabled_when_docx_run_is_invalid(self):
        source = APP_PATH.read_text(encoding="utf-8-sig")
        self.assertIn("docx_submission_blocked = cr_docx_only and not docx_run_profile", source)
        self.assertIn("disabled=not cr_spooler_available or docx_submission_blocked", source)


if __name__ == "__main__":
    unittest.main()
