from __future__ import annotations

import ast
import json
import os
import re
import shutil
import unittest
import uuid
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"

REFERE_HELPERS = {
    "_compte_rendu_nas_context_path",
    "_compte_rendu_mirror_affaire_bases",
    "_compte_rendu_out_dir_candidates",
    "_compte_rendu_trans_file_candidates",
    "_compte_rendu_any_path_to_volume1",
    "_compte_rendu_path_is_dir",
    "_compte_rendu_path_has_entries",
    "_compte_rendu_read_json_file",
    "_compte_rendu_any_file_exists",
    "inspect_compte_rendu_refere_source_run",
    "discover_compte_rendu_refere_source_runs",
    "resolve_compte_rendu_refere_source_run",
    "preflight_compte_rendu_refere_preventif",
    "build_compte_rendu_refere_preventif_job",
    "submit_compte_rendu_refere_preventif_job",
    "find_compte_rendu_refere_preventif_job_status",
}

SUPPORT_NAMES = {
    "_compte_rendu_nas_infos_path",
    "_compte_rendu_existing_run_basename",
    "discover_compte_rendu_run_dirs",
    "_audit_quality_read_text",
    "_audit_quality_read_json",
    "_audit_quality_find_log",
    "_audit_quality_manifest_value",
    "_safe_job_token",
}


def _load_functions(nas_root: Path, laptop_root: Path, jobs_root: Path):
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted = REFERE_HELPERS | SUPPORT_NAMES
    module = ast.Module(
        body=[node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in wanted],
        type_ignores=[],
    )
    ast.fix_missing_locations(module)
    namespace = {
        "Path": Path,
        "os": os,
        "re": re,
        "json": json,
        "uuid": uuid,
        "datetime": datetime,
        "NAS_AFFAIRES_ROOT": nas_root,
        "AFFAIRES_ROOT": laptop_root,
        "PCFIXE_AFFAIRES_ROOT": laptop_root,
        "AUDIT_REUNION_QUALITY_ARTIFACTS": (
            "global.json",
            "global_by_sujet.json",
            "global_final.json",
            "global_meeting.json",
            "debrief.json",
        ),
        "get_pcfixe_jobs_queued_dir": lambda: jobs_root / "queued",
        "get_pcfixe_jobs_root": lambda: jobs_root,
        "preflight_pcfixe_target_dir": lambda path: None,
    }
    exec(compile(module, str(APP_PATH), "exec"), namespace)
    return namespace


def _write_run(
    root: Path,
    affaire: str,
    captation: str,
    run_name: str,
    *,
    segments: bool = True,
    compacts: bool = True,
    qa: bool | None = True,
    metadata_ok: bool = True,
) -> Path:
    run_dir = (
        root
        / affaire
        / "BE_Traitement_captations"
        / captation
        / "compte_rendu_LLM"
        / "out"
        / run_name
    )
    run_dir.mkdir(parents=True, exist_ok=True)
    if segments:
        (run_dir / "segments").mkdir(exist_ok=True)
        (run_dir / "segments" / "segment_001.json").write_text("{}", encoding="utf-8")
    if compacts:
        (run_dir / "pass2E_sujets_compact").mkdir(exist_ok=True)
        (run_dir / "pass2E_sujets_compact" / "sujet_001_compact.json").write_text("{}", encoding="utf-8")
    if qa is not None:
        (run_dir / "pipeline_qa_status.json").write_text(
            json.dumps({"ok": bool(qa)}), encoding="utf-8"
        )
    if metadata_ok:
        logs = run_dir / "logs"
        logs.mkdir(exist_ok=True)
        (logs / "run_metadata_20260903_152405.json").write_text(
            json.dumps({"id_affaire": affaire, "id_captation": captation}),
            encoding="utf-8",
        )
    return run_dir


def _volume1_expected(affaire: str, captation: str, run_name: str, sub: str) -> str:
    return (
        f"/volume1/Affaires/{affaire}/BE_Traitement_captations/"
        f"{captation}/compte_rendu_LLM/out/{run_name}/{sub}"
    )


class CompteRenduReferePreventifTests(unittest.TestCase):
    def setUp(self):
        temp_root = Path(os.environ.get("LLM_ASSISTANT_TEST_TMP") or r"C:\CodexWorkspace\.tmp_llm_assistant_tests")
        temp_root.mkdir(exist_ok=True)
        self.base = temp_root / f"cr_refere_{uuid.uuid4().hex}"
        self.base.mkdir()
        self.nas_root = self.base / "NasAffaires"
        self.laptop_root = self.base / "LaptopAffaires"
        self.jobs_root = self.base / "jobs"
        for sub in ("queued", "running", "done", "failed", "logs"):
            (self.jobs_root / sub).mkdir(parents=True)
        self.ns = _load_functions(self.nas_root, self.laptop_root, self.jobs_root)
        self.affaire = "2025-J48"
        self.captation = "accedit-2025-11-20"
        self.run_name = "job_20260903_152404_444644044_4568"
        self.run_dir = _write_run(
            self.nas_root, self.affaire, self.captation, self.run_name,
        )
        for filename in ("infos_projet.json", "contexte_general_compte_rendu.json"):
            trans_dir = (
                self.nas_root / self.affaire / "AF_Expert_ASR" / "transcriptions" / self.captation
            )
            trans_dir.mkdir(parents=True, exist_ok=True)
            (trans_dir / filename).write_text("{}", encoding="utf-8")

    def tearDown(self):
        shutil.rmtree(self.base, ignore_errors=True)

    def _resolve(self):
        return self.ns["resolve_compte_rendu_refere_source_run"](self.affaire, self.captation)

    def _preflight(self):
        return self.ns["preflight_compte_rendu_refere_preventif"](self.affaire, self.captation)

    def _build(self, job_id="cr_refere_test_j48", profile=None):
        profile = profile if profile is not None else self._resolve()["profile"]
        return self.ns["build_compte_rendu_refere_preventif_job"](
            id_affaire=self.affaire,
            id_captation=self.captation,
            run_profile=profile,
            model_pass2f="pass2f_remote",
            job_id=job_id,
        )

    def test_newest_valid_run_selected_among_several(self):
        _write_run(
            self.nas_root, self.affaire, self.captation,
            "job_20260902_193906_396946666_23848",
        )
        _write_run(
            self.nas_root, self.affaire, self.captation,
            "job_20260903_084405_251651767_10740",
        )
        resolved = self._resolve()
        self.assertTrue(resolved["ok"])
        self.assertEqual(resolved["profile"]["run_name"], self.run_name)

    def test_run_with_qa_false_is_excluded_and_older_valid_is_used(self):
        _write_run(
            self.nas_root, self.affaire, self.captation,
            "job_20260904_000000_000000000_00000",
            qa=False,
        )
        resolved = self._resolve()
        self.assertTrue(resolved["ok"])
        self.assertEqual(resolved["profile"]["run_name"], self.run_name)

    def test_run_without_segments_or_compacts_is_rejected(self):
        self.assertTrue(self._resolve()["ok"])
        bad_run = (
            self.nas_root / self.affaire / "BE_Traitement_captations" / self.captation
            / "compte_rendu_LLM" / "out" / "job_20260905_000000_000000000_00000"
        )
        bad_run.mkdir(parents=True)
        profile = self.ns["inspect_compte_rendu_refere_source_run"](
            bad_run, id_affaire=self.affaire, id_captation=self.captation,
        )
        self.assertFalse(profile["valid"])
        self.assertTrue(any("segments" in reason for reason in profile["reasons"]))
        self.assertTrue(any("pass2E_sujets_compact" in reason for reason in profile["reasons"]))

    def test_run_with_other_affaire_metadata_is_rejected(self):
        other = _write_run(
            self.nas_root, "2026-A60", self.captation, self.run_name,
        )
        profile = self.ns["inspect_compte_rendu_refere_source_run"](
            other, id_affaire=self.affaire, id_captation=self.captation,
        )
        self.assertFalse(profile["valid"])
        self.assertTrue(any("autre affaire" in reason for reason in profile["reasons"]))

    def test_laptop_mirror_is_used_when_nas_has_no_valid_run(self):
        _write_run(
            self.laptop_root, self.affaire, self.captation,
            "job_20260903_152404_444644044_4568",
        )
        resolved = self._resolve()
        self.assertTrue(resolved["ok"])
        self.assertEqual(resolved["profile"]["run_name"], self.run_name)
        self.assertEqual(
            resolved["profile"]["run_nas"],
            _volume1_expected(self.affaire, self.captation, self.run_name, "").rstrip("/"),
        )

    def test_preflight_reports_missing_context_and_infos(self):
        preflight = self._preflight()
        self.assertTrue(preflight["ok"])
        self.assertEqual(preflight["missing"], [])
        shutil.rmtree(
            self.nas_root / self.affaire / "AF_Expert_ASR", ignore_errors=True
        )
        preflight = self._preflight()
        self.assertFalse(preflight["ok"])
        joined = "; ".join(preflight["missing"])
        self.assertIn("infos_projet.json", joined)
        self.assertIn("contexte_general_compte_rendu.json", joined)

    def test_exact_contract_json_has_only_the_fixed_keys(self):
        job = self._build()
        self.assertEqual(
            set(job),
            {
                "job_id",
                "type",
                "id_affaire",
                "id_captation",
                "infos_projet_path",
                "generalist_segments_dir",
                "pass2e_compacts_dir",
                "contexte_general_compte_rendu_path",
                "model_pass2f",
            },
        )
        self.assertEqual(job["job_id"], "cr_refere_test_j48")
        self.assertEqual(job["type"], "compte_rendu_refere_preventif")
        self.assertEqual(job["id_affaire"], self.affaire)
        self.assertEqual(job["id_captation"], self.captation)
        self.assertEqual(job["model_pass2f"], "pass2f_remote")
        self.assertNotIn("DataRoot", job)
        self.assertNotIn("OutDir", job)
        self.assertNotIn("RefereLotLabel", job)
        self.assertNotIn("nas_command", job)
        self.assertEqual(
            job["infos_projet_path"],
            f"/volume1/Affaires/{self.affaire}/AF_Expert_ASR/transcriptions/"
            f"{self.captation}/infos_projet.json",
        )
        self.assertEqual(
            job["contexte_general_compte_rendu_path"],
            f"/volume1/Affaires/{self.affaire}/AF_Expert_ASR/transcriptions/"
            f"{self.captation}/contexte_general_compte_rendu.json",
        )
        self.assertEqual(
            job["generalist_segments_dir"],
            _volume1_expected(self.affaire, self.captation, self.run_name, "segments"),
        )
        self.assertEqual(
            job["pass2e_compacts_dir"],
            _volume1_expected(self.affaire, self.captation, self.run_name, "pass2E_sujets_compact"),
        )

    def test_real_j48_contract_matches_validated_run_without_hardcoding(self):
        job = self._build()
        self.assertEqual(job["id_affaire"], "2025-J48")
        self.assertEqual(job["id_captation"], "accedit-2025-11-20")
        self.assertIn(f"/out/{self.run_name}/segments", job["generalist_segments_dir"])
        self.assertIn(f"/out/{self.run_name}/pass2E_sujets_compact", job["pass2e_compacts_dir"])
        source = APP_PATH.read_text(encoding="utf-8-sig")
        self.assertNotIn("job_20260903_152404_444644044_4568", source)

    def test_auto_job_id_shape(self):
        job = self._build(job_id="")
        self.assertTrue(job["job_id"].startswith("cr_refere_"))
        self.assertIn(self.affaire, job["job_id"])
        self.assertIn(self.captation, job["job_id"])
        self.assertRegex(job["job_id"], r"_\d{8}_\d{6}_[0-9a-f]{8}$")

    def test_build_rejects_wrong_model_and_invalid_affaire(self):
        with self.assertRaisesRegex(ValueError, "pass2f_remote"):
            self.ns["build_compte_rendu_refere_preventif_job"](
                id_affaire=self.affaire,
                id_captation=self.captation,
                run_profile=self._resolve()["profile"],
                model_pass2f="pass2f_local",
                job_id="x",
            )
        with self.assertRaisesRegex(ValueError, "invalides"):
            self.ns["build_compte_rendu_refere_preventif_job"](
                id_affaire="2025\\J48",
                id_captation=self.captation,
                run_profile=self._resolve()["profile"],
                model_pass2f="pass2f_remote",
                job_id="x",
            )

    def test_dry_run_submission_writes_nothing(self):
        job = self._build()
        result = self.ns["submit_compte_rendu_refere_preventif_job"](job, dry_run=True)
        self.assertEqual(result["status"], "dry-run")
        self.assertEqual(list((self.jobs_root / "queued").glob("*.json")), [])

    def test_atomic_submission_writes_queued_json_and_status_queued(self):
        job = self._build()
        result = self.ns["submit_compte_rendu_refere_preventif_job"](job, dry_run=False)
        self.assertEqual(result["status"], "queued")
        queued_file = self.jobs_root / "queued" / f"{job['job_id']}.json"
        self.assertTrue(queued_file.is_file())
        self.assertEqual(json.loads(queued_file.read_text(encoding="utf-8")), job)
        status = self.ns["find_compte_rendu_refere_preventif_job_status"](job["job_id"])
        self.assertEqual(status["status"], "queued")
        self.assertEqual(status["type"], "compte_rendu_refere_preventif")

    def test_follow_up_done_and_failed_with_manifest(self):
        ns = self.ns
        for job_id, status_dir, manifest_extra in (
            ("cr_refere_done_j48", "done", {"finished_at": "2026-09-10T10:00:00", "exit_code": "0"}),
            (
                "cr_refere_failed_j48",
                "failed",
                {"error_code": "LLM_TIMEOUT", "error_message": "appel LLM en echec"},
            ),
        ):
            job = {"job_id": job_id, "type": "compte_rendu_refere_preventif"}
            (self.jobs_root / status_dir / f"{job_id}.json").write_text(
                json.dumps(job), encoding="utf-8"
            )
            manifest = {
                "job_id": job_id,
                "status": status_dir,
                "started_at": "2026-09-10T09:00:00",
            }
            manifest.update(manifest_extra)
            (self.jobs_root / "logs" / f"{job_id}.manifest.json").write_text(
                json.dumps(manifest), encoding="utf-8"
            )
        done = ns["find_compte_rendu_refere_preventif_job_status"]("cr_refere_done_j48")
        self.assertEqual(done["status"], "done")
        self.assertEqual(done["finished_at"], "2026-09-10T10:00:00")
        self.assertEqual(done["exit_code"], "0")
        failed = ns["find_compte_rendu_refere_preventif_job_status"]("cr_refere_failed_j48")
        self.assertEqual(failed["status"], "failed")
        self.assertEqual(failed["error_code"], "LLM_TIMEOUT")
        self.assertEqual(failed["error_message"], "appel LLM en echec")

    def test_ui_markers_and_generalist_regression(self):
        source = APP_PATH.read_text(encoding="utf-8-sig")
        self.assertIn("Lancer le compte rendu \u2013 R\u00e9f\u00e9r\u00e9 pr\u00e9ventif", source)
        self.assertIn('key="cr_refere_preventif_submit"', source)
        self.assertIn("def submit_compte_rendu_job(", source)
        self.assertIn('"type": "compte_rendu"', source)
        self.assertIn("PHOTO_BATCH_RESET_VALUES = (", source)


if __name__ == "__main__":
    unittest.main()
