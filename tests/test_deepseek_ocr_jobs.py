from __future__ import annotations

import ast
import json
import os
import shutil
import types
import unittest
import uuid
from datetime import datetime
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_deepseek_helpers(base: Path):
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted = {
        "_pcfixe_affaires_share_candidates",
        "_safe_deepseek_job_token",
        "build_deepseek_ocr_job_id_value",
        "write_deepseek_ocr_job_manifest",
        "pcfixe_server_path_to_resolved_jobs_unc",
        "find_deepseek_ocr_job_status_in_root",
        "discover_deepseek_ocr_jobs_for_project_in_root",
        "deepseek_ocr_follow_state",
        "deepseek_ocr_should_scan_jobs",
    }
    module = ast.Module(
        body=[node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in wanted],
        type_ignores=[],
    )
    ast.fix_missing_locations(module)

    real_os = os
    replace_calls = []

    def fake_replace(src, dst):
        replace_calls.append((str(src), str(dst)))
        real_os.replace(src, dst)

    fake_os = types.SimpleNamespace(replace=fake_replace)

    def load_json(path, default):
        try:
            return json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception:
            return default

    ns = {
        "Path": Path,
        "re": __import__("re"),
        "json": json,
        "os": fake_os,
        "uuid": uuid,
        "datetime": datetime,
        "unicodedata": __import__("unicodedata"),
        "PCFIXE_AFFAIRES_SHARE_CANDIDATES": (r"\\192.168.0.155\Affaires", r"\\10.0.1.10\Affaires"),
        "_pcfixe_vpn_active": lambda: False,
        "get_pcfixe_jobs_queued_dir": lambda: base / "_jobs" / "queued",
        "get_pcfixe_jobs_root": lambda: base / "_jobs",
        "get_pcfixe_affaires_root": lambda: base,
        "preflight_pcfixe_target_dir": lambda path: None,
        "_norm": lambda value: str(value or "").replace("/", "\\").rstrip("\\/ "),
        "pj": lambda *parts: str(Path(str(parts[0])).joinpath(*[str(p) for p in parts[1:] if str(p)])),
        "load_json": load_json,
        "_replace_calls": replace_calls,
    }
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns


class DeepSeekOcrJobsTests(unittest.TestCase):
    def setUp(self):
        temp_root = Path(os.environ.get("LLM_ASSISTANT_TEST_TMP") or r"C:\CodexWorkspace\.tmp_llm_assistant_tests")
        temp_root.mkdir(exist_ok=True)
        self.base = temp_root / f"deepseek_ocr_{uuid.uuid4().hex}"
        self.base.mkdir()
        self.ns = _load_deepseek_helpers(self.base)

    def tearDown(self):
        shutil.rmtree(self.base, ignore_errors=True)

    def _job(self, job_id: str = "deepseek_ocr_2026-A60_LETTRE_N1_20260806_120000") -> dict:
        return {
            "job_id": job_id,
            "type": "deepseek_ocr",
            "project_id": "2026-A60",
            "source_pdf": r"C:\Affaires\2026-A60\AA_Expert_Admin\Depot_initial\LETTRE N°1 du 29 mai 2026_pages_user_2.pdf",
            "pages": [2],
            "output_dir": r"C:\Affaires\2026-A60\AD_Expert_Traitements\_OCR_Dire_Bordereau\LETTRE N°1 du 29 mai 2026_pages_user_2",
            "png_dir": r"C:\Affaires\2026-A60\AD_Expert_Traitements\_OCR_Dire_Bordereau\_pages_png\LETTRE N°1 du 29 mai 2026_pages_user_2",
            "tile_count": 2,
            "postprocess": True,
            "retry_glitch_pages": True,
            "tile_glitch_pages": True,
            "fallback_tesseract_pages": False,
        }

    def test_vpn_active_prioritizes_vpn_jobs_root_before_lan(self):
        roots = self.ns["_pcfixe_affaires_share_candidates"](vpn_active=True)
        self.assertEqual(roots[0], r"\\10.0.1.10\Affaires")
        self.assertIn(r"\\192.168.0.155\Affaires", roots[1:])

    def test_deepseek_dry_run_branch_does_not_write_manifest(self):
        text = APP_PATH.read_text(encoding="utf-8-sig")
        start = text.index('        if ocr_engine == "DeepSeekOCR avancé":\n            dry_run_docs')
        end = text.index('    if ocr_engine == "DeepSeekOCR avancé":\n        with st.expander("Diagnostic DeepSeekOCR VPN / jobs"', start)
        dry_run_branch = text[start:end]
        self.assertIn("st.session_state[deepseek_state_key] = dry_run_docs", dry_run_branch)
        self.assertNotIn("write_deepseek_ocr_job", dry_run_branch)

    def test_real_submit_creates_manifest_in_queued(self):
        target = self.ns["write_deepseek_ocr_job_manifest"](self._job())
        path = Path(target)
        self.assertEqual(path.parent, self.base / "_jobs" / "queued")
        self.assertTrue(path.is_file())
        manifest = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(manifest["type"], "deepseek_ocr")
        self.assertEqual(manifest["pages"], [2])

    def test_atomic_write_uses_temp_file_then_os_replace(self):
        target = Path(self.ns["write_deepseek_ocr_job_manifest"](self._job()))
        calls = self.ns["_replace_calls"]
        self.assertEqual(len(calls), 1)
        tmp, final = calls[0]
        self.assertTrue(tmp.endswith(".tmp"))
        self.assertEqual(final, str(target))
        self.assertFalse(Path(tmp).exists())

    def test_follow_state_keeps_job_id_when_scan_does_not_find_job(self):
        state = self.ns["deepseek_ocr_follow_state"](
            "deepseek_ocr_2026-A60_missing",
            [],
            {"json_path": "queued.json"},
        )
        self.assertEqual(state["job_id"], "deepseek_ocr_2026-A60_missing")
        self.assertEqual(state["state"], "déposé_non_retrouvé")
        self.assertFalse(state["found"])

    def test_empty_job_id_disables_queue_scan(self):
        self.assertFalse(self.ns["deepseek_ocr_should_scan_jobs"](""))
        self.assertFalse(self.ns["deepseek_ocr_should_scan_jobs"]("   "))
        self.assertTrue(self.ns["deepseek_ocr_should_scan_jobs"]("deepseek_ocr_2026-A60_doc"))

    def test_empty_job_id_ui_disables_check_button_and_avoids_scan(self):
        text = APP_PATH.read_text(encoding="utf-8-sig")
        start = text.index('if not deepseek_ocr_should_scan_jobs(session_last_job):')
        end = text.index("        else:\n            discovered_jobs_all", start)
        empty_job_block = text[start:end]
        self.assertIn("Aucun job DeepSeekOCR n’a encore été déposé pour cette sélection.", empty_job_block)
        self.assertIn('disabled=True', empty_job_block)
        self.assertNotIn("discover_deepseek_ocr_jobs_for_project", empty_job_block)
        self.assertNotIn("find_deepseek_ocr_job_status", empty_job_block)

    def test_dry_run_shows_next_step_and_submit_panel_before_stop(self):
        text = APP_PATH.read_text(encoding="utf-8-sig")
        start = text.index("    analyze_dire_bord_clicked = st.button")
        end = text.index('    if ocr_engine == "DeepSeekOCR avancé" and st.session_state.get(deepseek_state_key):', start)
        dry_run_branch = text[start:end]
        self.assertIn("Préparation terminée. Aucun job OCR n’a encore été déposé sur le PC fixe.", dry_run_branch)
        self.assertIn("Étape suivante : cliquer sur « Déposer le job DeepSeekOCR PC fixe ».", dry_run_branch)
        panel_start = end
        panel_end = text.index('    if ocr_engine == "DeepSeekOCR avancé":\n        st.markdown("#### Suivi du job DeepSeekOCR")', panel_start)
        panel_block = text[panel_start:panel_end]
        self.assertEqual(panel_block.count("render_deepseek_submit_panel()"), 1)
        self.assertIn("if deepseek_stop_after_submit_panel:", panel_block)
        self.assertLess(panel_block.index("render_deepseek_submit_panel()"), panel_block.index("st.stop()"))

    def test_submit_panel_is_rendered_once_per_deepseek_run(self):
        text = APP_PATH.read_text(encoding="utf-8-sig")
        submit_call_count = text.count("render_deepseek_submit_panel(") - text.count("def render_deepseek_submit_panel(")
        self.assertEqual(submit_call_count, 1)
        self.assertEqual(text.count('key=f"submit_deepseek_ocr_job_{project_id}"'), 1)
        self.assertIn("deepseek_stop_after_submit_panel = True", text)

    def test_submit_diagnostic_includes_jobs_root_queue_job_id_and_manifest(self):
        text = APP_PATH.read_text(encoding="utf-8-sig")
        start = text.index('with st.expander("Diagnostic dépôt DeepSeekOCR"')
        end = text.index('                    st.json(created["job"])', start)
        diagnostic_block = text[start:end]
        self.assertIn("Racine _jobs retenue", diagnostic_block)
        self.assertIn("Chemin complet queued", diagnostic_block)
        self.assertIn("job_id créé", diagnostic_block)
        self.assertIn("Manifest écrit", diagnostic_block)

    def test_discovery_finds_same_job_in_each_status_folder(self):
        jobs_root = self.base / "_jobs"
        for status in ("queued", "running", "done", "failed"):
            shutil.rmtree(jobs_root, ignore_errors=True)
            folder = jobs_root / status
            folder.mkdir(parents=True)
            job = self._job(job_id=f"deepseek_ocr_2026-A60_{status}_20260806_120000")
            (folder / f"{job['job_id']}.json").write_text(json.dumps(job), encoding="utf-8")
            found = self.ns["find_deepseek_ocr_job_status_in_root"](job["job_id"], jobs_root)
            self.assertEqual(found["status"], status)
            discovered = self.ns["discover_deepseek_ocr_jobs_for_project_in_root"]("2026-A60", jobs_root)
            self.assertEqual(discovered[0]["job_id"], job["job_id"])
            self.assertEqual(discovered[0]["status"], status)

    def test_pdf_name_with_spaces_accents_and_degree_is_safe_in_job_id(self):
        job_id = self.ns["build_deepseek_ocr_job_id_value"](
            "2026-A60",
            "LETTRE N°1 du 29 mai 2026",
            stamp="20260806_120000",
        )
        self.assertEqual(job_id, "deepseek_ocr_2026-A60_LETTRE_N_1_du_29_mai_2026_20260806_120000")

    def test_tesseract_still_calls_direct_ocr_route(self):
        text = APP_PATH.read_text(encoding="utf-8-sig")
        start = text.index("        def _maybe_ocr")
        end = text.index("        def _read_ocr_csv_text_lines", start)
        maybe_ocr = text[start:end]
        self.assertIn('ocr_url = f"{SERVER_URL}/ocr"', maybe_ocr)
        self.assertIn("requests.post(", maybe_ocr)
        self.assertNotIn("write_deepseek_ocr_job", maybe_ocr)
        self.assertNotIn("get_pcfixe_jobs_queued_dir", maybe_ocr)


if __name__ == "__main__":
    unittest.main()
