"""Garde-fou de depot des jobs compte_rendu_refere_preventif (LLM_Assistant).

Aucun acces reseau, aucun spooler, aucun job reel : la racine _jobs est injectee
dans un dossier temporaire. Les tests couvrent :

- depot reussi -> status="queued" et deposit_verified=True ;
- fichier absent apres ecriture -> erreur explicite ;
- manifest relu avec un job_id different -> erreur explicite ;
- message d'echec enrichi (chemin tente, racine queue, hote SMB, temporaire, exception) ;
- suivi immediat retournant deja failed (consommation rapide par le spooler).
"""

import ast
import json
import os
import re
import tempfile
import unittest
import uuid
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"

FUNCTIONS = {
    "submit_compte_rendu_refere_preventif_job",
    "find_compte_rendu_refere_preventif_job_status",
    "_safe_job_token",
    "_unc_host",
    "_audit_quality_read_text",
    "_audit_quality_read_json",
    "_audit_quality_find_log",
    "_audit_quality_manifest_value",
}


def _load_namespace(jobs_root, *, os_module=None, preflight=None):
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    body = [
        node
        for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS
    ]
    namespace = {
        "Path": Path,
        "os": os_module or os,
        "re": re,
        "json": json,
        "uuid": uuid,
        "datetime": datetime,
        "get_pcfixe_jobs_queued_dir": lambda: jobs_root / "queued",
        "get_pcfixe_jobs_root": lambda: jobs_root,
        "preflight_pcfixe_target_dir": preflight or (lambda path: None),
    }
    exec(compile(ast.Module(body=body, type_ignores=[]), str(APP_PATH), "exec"), namespace)
    return namespace


class _OsShim:
    """Proxy de os dont seul replace() est instrumente."""

    def __init__(self, replace_impl):
        self._replace_impl = replace_impl

    def replace(self, src, dst):
        return self._replace_impl(src, dst)

    def __getattr__(self, name):
        return getattr(os, name)


def _build_job(job_id):
    return {
        "job_id": job_id,
        "type": "compte_rendu_refere_preventif",
        "id_affaire": "2025-J48",
        "id_captation": "accedit-2025-11-20",
        "infos_projet_path": (
            "/volume1/Affaires/2025-J48/AF_Expert_ASR/transcriptions/"
            "accedit-2025-11-20/infos_projet.json"
        ),
        "generalist_segments_dir": (
            "/volume1/Affaires/2025-J48/BE_Traitement_captations/"
            "accedit-2025-11-20/compte_rendu_LLM/out/job_generaliste/segments"
        ),
        "pass2e_compacts_dir": (
            "/volume1/Affaires/2025-J48/BE_Traitement_captations/"
            "accedit-2025-11-20/compte_rendu_LLM/out/job_generaliste/pass2E_sujets_compact"
        ),
        "contexte_general_compte_rendu_path": (
            "/volume1/Affaires/2025-J48/AF_Expert_ASR/transcriptions/"
            "accedit-2025-11-20/contexte_general_compte_rendu.json"
        ),
        "model_pass2f": "pass2f_remote",
    }


class SubmitGuardTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.jobs_root = Path(self.temp.name) / "_jobs"
        (self.jobs_root / "queued").mkdir(parents=True)

    def tearDown(self):
        self.temp.cleanup()

    def _queued_file(self, job_id):
        return self.jobs_root / "queued" / f"{job_id}.json"

    # ---------------------------------------------------- depot nominal
    def test_depot_reussi_status_queued_et_depot_verifie(self):
        ns = _load_namespace(self.jobs_root)
        job = _build_job("cr_refere_j48_ok")

        result = ns["submit_compte_rendu_refere_preventif_job"](job, dry_run=False)

        self.assertEqual(result["status"], "queued")
        self.assertTrue(result["deposit_verified"])
        self.assertEqual(result["queued_root"], str(self.jobs_root / "queued"))
        self.assertEqual(result["job_path"], str(self._queued_file("cr_refere_j48_ok")))

        queued_file = self._queued_file("cr_refere_j48_ok")
        self.assertTrue(queued_file.is_file())
        self.assertEqual(json.loads(queued_file.read_text(encoding="utf-8")), job)
        self.assertEqual(
            list((self.jobs_root / "queued").glob("*.tmp")),
            [],
            "aucun temporaire ne doit subsister apres os.replace",
        )

        status = ns["find_compte_rendu_refere_preventif_job_status"]("cr_refere_j48_ok")
        self.assertEqual(status["status"], "queued")
        self.assertEqual(status["type"], "compte_rendu_refere_preventif")

    def test_dry_run_ne_depose_rien_et_n_est_pas_verifie(self):
        ns = _load_namespace(self.jobs_root)
        job = _build_job("cr_refere_j48_dry")

        result = ns["submit_compte_rendu_refere_preventif_job"](job, dry_run=True)

        self.assertEqual(result["status"], "dry-run")
        self.assertFalse(result["deposit_verified"])
        self.assertEqual(list((self.jobs_root / "queued").glob("*.json")), [])

    # ------------------------------------------------- ecriture non prouvee
    def test_fichier_absent_apres_ecriture_leve_erreur_explicite(self):
        def replace_then_drop(src, dst):
            os.replace(src, dst)
            os.unlink(dst)

        ns = _load_namespace(self.jobs_root, os_module=_OsShim(replace_then_drop))
        job = _build_job("cr_refere_j48_absent")

        with self.assertRaises(RuntimeError) as ctx:
            ns["submit_compte_rendu_refere_preventif_job"](job, dry_run=False)

        message = str(ctx.exception)
        self.assertIn("non verifie", message)
        self.assertIn("fichier absent apres ecriture", message)
        self.assertIn(str(self._queued_file("cr_refere_j48_absent")), message)
        self.assertIn(str(self.jobs_root / "queued"), message)
        self.assertIn(".json.tmp", message)

    def test_manifest_relu_avec_job_id_different_leve_erreur_explicite(self):
        def replace_then_tamper(src, dst):
            os.replace(src, dst)
            data = json.loads(Path(dst).read_text(encoding="utf-8"))
            data["job_id"] = "cr_refere_autre_job"
            Path(dst).write_text(json.dumps(data), encoding="utf-8")

        ns = _load_namespace(self.jobs_root, os_module=_OsShim(replace_then_tamper))
        job = _build_job("cr_refere_j48_tamper")

        with self.assertRaises(RuntimeError) as ctx:
            ns["submit_compte_rendu_refere_preventif_job"](job, dry_run=False)

        message = str(ctx.exception)
        self.assertIn("incoherent", message)
        self.assertIn("job_id relu='cr_refere_autre_job'", message)
        self.assertIn("attendu='cr_refere_j48_tamper'", message)

    def test_echec_depot_message_enrichi(self):
        def preflight_boom(path):
            raise FileNotFoundError(2, "chemin d'acces introuvable")

        ns = _load_namespace(self.jobs_root, preflight=preflight_boom)
        job = _build_job("cr_refere_j48_boom")

        with self.assertRaises(RuntimeError) as ctx:
            ns["submit_compte_rendu_refere_preventif_job"](job, dry_run=False)

        message = str(ctx.exception)
        self.assertIn("chemin tente=", message)
        self.assertIn(str(self._queued_file("cr_refere_j48_boom")), message)
        self.assertIn("racine queue=", message)
        self.assertIn(str(self.jobs_root / "queued"), message)
        self.assertIn("hote SMB=", message)
        self.assertIn("temporaire=", message)
        self.assertIn(".json.tmp", message)
        self.assertIn("exception=FileNotFoundError", message)
        self.assertIn("chemin d'acces introuvable", message)
        self.assertFalse(self._queued_file("cr_refere_j48_boom").exists())

    # ------------------------------- consommation rapide par le spooler
    def test_suivi_immediat_deja_failed_apres_consommation_rapide(self):
        ns = _load_namespace(self.jobs_root)
        job_id = "cr_refere_consomme_vite"
        result = ns["submit_compte_rendu_refere_preventif_job"](
            _build_job(job_id), dry_run=False
        )
        self.assertEqual(result["status"], "queued")
        self.assertTrue(result["deposit_verified"])

        # Le spooler consomme la queue en ~1 s : queued -> failed.
        failed_dir = self.jobs_root / "failed"
        failed_dir.mkdir(parents=True, exist_ok=True)
        os.replace(self._queued_file(job_id), failed_dir / f"{job_id}.json")
        failed_path = failed_dir / f"{job_id}.json"
        manifest = json.loads(failed_path.read_text(encoding="utf-8"))
        manifest["error_code"] = "PROCESS_FAILED"
        manifest["error_message"] = "CSV ASR complet introuvable"
        failed_path.write_text(json.dumps(manifest), encoding="utf-8")

        status = ns["find_compte_rendu_refere_preventif_job_status"](job_id)

        self.assertEqual(status["status"], "failed")
        self.assertEqual(status["error_code"], "PROCESS_FAILED")
        self.assertIn("CSV ASR complet introuvable", status["error_message"])
        self.assertEqual(status["type"], "compte_rendu_refere_preventif")
        self.assertEqual(status["job_path"], str(failed_path))
        self.assertFalse(self._queued_file(job_id).exists())


if __name__ == "__main__":
    unittest.main()
