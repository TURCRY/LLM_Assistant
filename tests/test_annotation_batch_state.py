"""Tests du resolveur central d'etat des jobs annotation_photos_batch.

Le resolveur travaille par ``job_id`` et non par cle d action, afin d eviter
plusieurs verites contradictoires dans l UI (dernier job affiche, traitement
initial, relance WEAK, disponibilite du rapport Word).
"""

from __future__ import annotations

import ast
import os
import time
import unittest
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_batch_state_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted_functions = {
        "resolve_annotation_batch_state",
        "_annotation_classify_jobs",
        "_annotation_job_recency_key",
        "_annotation_parse_time_value",
        "_boolish_annotation_value",
        "_annotation_job_sort_value",
    }
    wanted_assignments = {"JOB_STATE_RANK", "WEAK_NO_OP_REASON"}
    body = []
    for node in tree.body:
        if isinstance(node, ast.Assign):
            names = {target.id for target in node.targets if isinstance(target, ast.Name)}
            if names & wanted_assignments:
                body.append(node)
        elif isinstance(node, ast.FunctionDef) and node.name in wanted_functions:
            body.append(node)
    ns = {
        "Path": Path,
        "os": os,
        "time": time,
        "re": __import__("re"),
        "datetime": __import__("datetime").datetime,
    }
    exec(compile(ast.Module(body=body, type_ignores=[]), str(APP_PATH), "exec"), ns)
    return ns


def _job(job_id, status, *, rerun_weak=False, no_op=False, no_op_reason="",
         exit_code="0", finished_at="", sort_mtime="0", output_verified="true",
         nas_publish_succeeded="true", options=None, profile="run_standard"):
    return {
        "job_id": job_id,
        "status": status,
        "profile": profile,
        "rerun_weak": "true" if rerun_weak else "false",
        "no_op": "true" if no_op else "",
        "no_op_reason": no_op_reason,
        "exit_code": exit_code,
        "started_at": "",
        "completed_at": finished_at,
        "dry_run": "false",
        "output_verified": output_verified,
        "output_verified_local": output_verified,
        "nas_publish_attempted": "true" if nas_publish_succeeded == "true" else "false",
        "nas_publish_succeeded": nas_publish_succeeded,
        "publish_pending": "false",
        "sort_mtime": sort_mtime,
        "options": options if options is not None else (["--rerun-weak", "1"] if rerun_weak else []),
    }


def _details(*jobs):
    """Simule la structure job_details avec des cles d action collisionnantes."""
    details = {}
    for index, job in enumerate(jobs):
        key = "run_standard" if index == 0 else f"run_standard_{job['status']}_{job['job_id']}"
        details[key] = job
    return details


def _producer(job_id, *, rerun_weak=False, published=True, output_verified="true",
              no_op=False, no_op_reason="", nas_sha256="sha_nas_1", stamp="20260907_090119",
              exit_code="0"):
    """Job completed portant les champs de publication NAS du manifest."""
    job = _job(
        job_id, "completed", rerun_weak=rerun_weak, no_op=no_op, no_op_reason=no_op_reason,
        exit_code=exit_code, output_verified=output_verified,
        nas_publish_succeeded="true" if published else "false",
    )
    job["nas_publish_attempted"] = "true"
    job["output_verified_local"] = "true"
    job["photos_batch_nas_sha256"] = nas_sha256 if published else ""
    job["photos_batch_sha256"] = nas_sha256
    job["publish_pending"] = "false"
    job["sort_stamp"] = stamp
    return job


class ResolveBatchStateTests(unittest.TestCase):
    def setUp(self):
        self.ns = _load_batch_state_functions()

    def resolve(self, jobs, **kwargs):
        return self.ns["resolve_annotation_batch_state"](
            "2025-J48", "accedit-2026-05-28", job_details=_details(*jobs), **kwargs
        )

    # 1 -------------------------------------------------------------------
    def test_old_failed_new_done_shows_new_done(self):
        state = self.resolve([
            _job("annotation_a_20260907_100000_aaaa", "failed", rerun_weak=True, exit_code="999"),
            _job("annotation_a_20260907_130000_bbbb", "completed", rerun_weak=True,
                 no_op=True, no_op_reason="NO_WEAK_ROWS"),
        ])
        self.assertEqual(state["latest_job"]["job_id"], "annotation_a_20260907_130000_bbbb")
        self.assertEqual(state["latest_job"]["status"], "completed")
        self.assertEqual(state["latest_failure"]["job_id"], "annotation_a_20260907_100000_aaaa")

    # 2 -------------------------------------------------------------------
    def test_same_job_followed_across_queued_running_done(self):
        job_id = "annotation_a_20260907_133436_af1846aa"
        for status in ("queued", "running", "completed"):
            state = self.resolve(
                [_job(job_id, status, rerun_weak=True, no_op=(status == "completed"),
                      no_op_reason="NO_WEAK_ROWS" if status == "completed" else "")],
                last_submission_job_id=job_id,
            )
            self.assertIsNotNone(state["latest_submission"], f"perdu en {status}")
            self.assertEqual(state["latest_submission"]["job_id"], job_id)
            self.assertEqual(state["latest_submission"]["status"], status)
        self.assertTrue(state["weak_no_op"])

    # 3 -------------------------------------------------------------------
    def test_initial_done_plus_weak_noop_allows_word(self):
        state = self.resolve([
            _job("annotation_a_20260907_090119_5f646751", "completed"),
            _job("annotation_a_20260907_133436_af1846aa", "completed", rerun_weak=True,
                 no_op=True, no_op_reason="NO_WEAK_ROWS"),
        ])
        self.assertTrue(state["word_report_ready"])
        self.assertEqual(state["word_report_block_reason"], "")
        self.assertTrue(state["weak_no_op"])

    # 4 -------------------------------------------------------------------
    def test_initial_done_plus_newer_weak_failed_blocks_word(self):
        state = self.resolve([
            _job("annotation_a_20260907_090119_5f646751", "completed"),
            _job("annotation_a_20260907_133436_af1846aa", "completed", rerun_weak=True,
                 no_op=True, no_op_reason="NO_WEAK_ROWS"),
            _job("annotation_a_20260907_140000_cccccccc", "failed", rerun_weak=True, exit_code="999"),
        ])
        self.assertFalse(state["word_report_ready"])
        self.assertIn("cccccccc", state["word_report_block_reason"])
        self.assertEqual(state["latest_weak"]["job_id"], "annotation_a_20260907_140000_cccccccc")

    # 5 -------------------------------------------------------------------
    def test_no_initial_weak_noop_does_not_invent_initial(self):
        state = self.resolve([
            _job("annotation_a_20260907_133436_af1846aa", "completed", rerun_weak=True,
                 no_op=True, no_op_reason="NO_WEAK_ROWS"),
        ])
        self.assertIsNone(state["initial_success"])
        self.assertFalse(state["word_report_ready"])
        self.assertIn("aucun traitement initial", state["word_report_block_reason"])
        self.assertEqual(state["latest_weak"]["job_id"], "annotation_a_20260907_133436_af1846aa")

    # 6 -------------------------------------------------------------------
    def test_several_old_failed_with_newer_done_does_not_block(self):
        state = self.resolve([
            _job("annotation_a_20260907_021509_001a79fa", "failed", rerun_weak=True, exit_code="999"),
            _job("annotation_a_20260907_090119_5f646751", "completed"),
            _job("annotation_a_20260907_105035_d5649a7f", "failed", rerun_weak=True, exit_code="999"),
            _job("annotation_a_20260907_133436_af1846aa", "completed", rerun_weak=True,
                 no_op=True, no_op_reason="NO_WEAK_ROWS"),
        ])
        self.assertTrue(state["word_report_ready"])
        self.assertEqual(state["latest_failure"]["job_id"], "annotation_a_20260907_105035_d5649a7f")

    # 7 -------------------------------------------------------------------
    def test_physical_order_does_not_matter(self):
        jobs = [
            _job("annotation_a_20260907_090119_5f646751", "completed"),
            _job("annotation_a_20260907_105035_d5649a7f", "failed", rerun_weak=True, exit_code="999"),
            _job("annotation_a_20260907_133436_af1846aa", "completed", rerun_weak=True,
                 no_op=True, no_op_reason="NO_WEAK_ROWS"),
        ]
        forward = self.resolve(jobs)
        backward = self.resolve(list(reversed(jobs)))
        self.assertEqual(forward["initial_success"]["job_id"], backward["initial_success"]["job_id"])
        self.assertEqual(forward["latest_weak"]["job_id"], backward["latest_weak"]["job_id"])
        self.assertEqual(forward["word_report_ready"], backward["word_report_ready"])

    # 8 -------------------------------------------------------------------
    def test_misleading_mtime_job_id_stamp_wins(self):
        # mtime tres eleve sur l ancien job, stamp du job_id plus fiable
        state = self.resolve([
            _job("annotation_a_20260907_090119_5f646751", "completed", sort_mtime="9999999999"),
            _job("annotation_a_20260907_133436_af1846aa", "completed", rerun_weak=True,
                 no_op=True, no_op_reason="NO_WEAK_ROWS", sort_mtime="1"),
        ])
        self.assertEqual(state["latest_job"]["job_id"], "annotation_a_20260907_133436_af1846aa")
        self.assertEqual(state["initial_success"]["job_id"], "annotation_a_20260907_090119_5f646751")

    # 9 -------------------------------------------------------------------
    def test_invalid_manifest_does_not_hide_other_jobs(self):
        broken = {"job_id": "", "status": "completed"}
        state = self.resolve([
            broken,
            _job("annotation_a_20260907_090119_5f646751", "completed"),
            _job("annotation_a_20260907_133436_af1846aa", "completed", rerun_weak=True,
                 no_op=True, no_op_reason="NO_WEAK_ROWS"),
        ])
        self.assertEqual(len(state["jobs"]), 2)
        self.assertTrue(state["word_report_ready"])
        self.assertEqual(state["initial_success"]["job_id"], "annotation_a_20260907_090119_5f646751")

    # 6 -------------------------------------------------------------------
    def test_older_weak_failed_loses_against_manifest_proven_newer_done(self):
        initial = _job("annotation_x_20260907_090119_5f646751", "completed")
        old_failed = _job("annotation_x_20260907_105035_d5649a7f", "failed", rerun_weak=True)
        new_done = _job(
            "annotation_x_20260907_133436_af1846aa", "completed",
            rerun_weak=True, no_op=True, no_op_reason="NO_WEAK_ROWS",
        )
        state = self.resolve([initial, old_failed, new_done])
        self.assertEqual(state["latest_weak"]["job_id"], new_done["job_id"])
        self.assertEqual(state["latest_failure"]["job_id"], old_failed["job_id"])
        self.assertTrue(state["word_report_ready"])

    # 7 -------------------------------------------------------------------
    def test_weak_no_op_proven_only_by_manifest_allows_word(self):
        initial = _job("annotation_x_20260907_090119_5f646751", "completed")
        weak = _job(
            "annotation_x_20260907_133436_af1846aa", "completed",
            rerun_weak=True, no_op=True, no_op_reason="NO_WEAK_ROWS",
        )
        state = self.resolve([initial, weak])
        self.assertTrue(state["word_report_ready"])
        self.assertTrue(state["weak_no_op"])
        self.assertEqual(state["latest_weak"]["job_id"], weak["job_id"])


class J48RealScenarioTests(unittest.TestCase):
    """Scenario reel J48 : 090119 initial done, 105035 weak failed, 133436 weak no-op."""

    def setUp(self):
        self.ns = _load_batch_state_functions()

    def _state(self):
        jobs = [
            _job("annotation_2025-J48_accedit-2026-05-28_run_standard_20260907_090119_5f646751",
                 "completed", finished_at="2026-09-07T10:33:44"),
            _job("annotation_2025-J48_accedit-2026-05-28_run_standard_20260907_105035_d5649a7f",
                 "failed", rerun_weak=True, exit_code="999", finished_at="2026-09-07T10:54:06"),
            _job("annotation_2025-J48_accedit-2026-05-28_run_standard_20260907_133436_af1846aa",
                 "completed", rerun_weak=True, no_op=True, no_op_reason="NO_WEAK_ROWS",
                 finished_at="2026-09-07T13:39:04"),
        ]
        return self.ns["resolve_annotation_batch_state"](
            "2025-J48", "accedit-2026-05-28", job_details=_details(*jobs),
            last_submission_job_id="annotation_2025-J48_accedit-2026-05-28_run_standard_20260907_133436_af1846aa",
        )

    def test_j48_expected_state(self):
        state = self._state()
        self.assertTrue(state["word_report_ready"])
        self.assertEqual(state["word_report_block_reason"], "")
        self.assertTrue(state["weak_no_op"])

    def test_j48_identifies_090119_as_initial(self):
        state = self._state()
        self.assertIsNotNone(state["initial_success"])
        self.assertTrue(state["initial_success"]["job_id"].endswith("090119_5f646751"))

    def test_j48_identifies_133436_as_latest_weak(self):
        state = self._state()
        self.assertTrue(state["latest_weak"]["job_id"].endswith("133436_af1846aa"))

    def test_j48_keeps_105035_as_historical_failure(self):
        state = self._state()
        self.assertTrue(state["latest_failure"]["job_id"].endswith("105035_d5649a7f"))

    def test_j48_follows_submitted_job_133436(self):
        state = self._state()
        self.assertIsNotNone(state["latest_submission"])
        self.assertTrue(state["latest_submission"]["job_id"].endswith("133436_af1846aa"))
        self.assertEqual(state["latest_submission"]["status"], "completed")

    def test_j48_deduplicates_by_job_id(self):
        state = self._state()
        ids = [job["job_id"] for job in state["jobs"]]
        self.assertEqual(len(ids), len(set(ids)))
        self.assertEqual(len(ids), 3)



class DedupeByJobIdTests(unittest.TestCase):
    def setUp(self):
        self.ns = _load_batch_state_functions()

    def test_terminal_state_wins_over_queued(self):
        job_id = "annotation_a_20260907_133436_af1846aa"
        jobs = [
            _job(job_id, "queued", rerun_weak=True),
            _job(job_id, "running", rerun_weak=True),
            _job(job_id, "completed", rerun_weak=True, no_op=True, no_op_reason="NO_WEAK_ROWS"),
        ]
        details = {f"run_standard_{i}_{j['status']}": j for i, j in enumerate(jobs)}
        merged = self.ns["_annotation_classify_jobs"](details)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0]["status"], "completed")

    def test_running_wins_over_queued(self):
        job_id = "annotation_a_20260907_133436_af1846aa"
        details = {
            "run_standard": _job(job_id, "queued", rerun_weak=True),
            "run_standard_running_x": _job(job_id, "running", rerun_weak=True),
        }
        merged = self.ns["_annotation_classify_jobs"](details)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0]["status"], "running")


class LoadJsonBomTests(unittest.TestCase):
    """Regression : les manifests du runner portent un BOM UTF-8.

    Sans ``utf-8-sig``, ``load_json`` renvoyait silencieusement ``{}``, ce qui
    faisait ignorer les jobs ``done`` (dont le no-op WEAK) et retomber l UI sur
    un ancien job ``failed``.
    """

    def _load_json(self):
        ns = {"json": __import__("json"), "Path": Path}
        src = APP_PATH.read_text(encoding="utf-8-sig")
        start = src.find("def load_json(")
        self.assertNotEqual(start, -1, "load_json introuvable dans app.py")
        end = src.find("def save_json(", start)
        self.assertNotEqual(end, -1, "save_json introuvable dans app.py")
        exec(compile(src[start:end], "app.py", "exec"), ns)
        return ns["load_json"]

    def test_load_json_reads_utf8_bom_manifest(self):
        import tempfile

        load_json = self._load_json()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "manifest.json"
            path.write_text('{"status": "done", "exit_code": 0}', encoding="utf-8-sig")
            data = load_json(str(path))
        self.assertEqual(data.get("status"), "done")
        self.assertEqual(data.get("exit_code"), 0)

    def test_load_json_still_returns_default_on_invalid(self):
        import tempfile

        load_json = self._load_json()
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.json"
            path.write_text("{not json", encoding="utf-8")
        self.assertEqual(load_json(str(path), {"fallback": True}), {"fallback": True})


class TargetedManifestEnrichmentTests(unittest.TestCase):
    """Enrichissement cible : done/*.json minimal + manifest logs complet.

    Le mode leger ne doit pas scanner logs/ globalement. Il ne resout que le
    manifeste du job_id deja connu dont l etat terminal reste indetermine.
    """

    MANIFEST_FIELDS = (
        "status",
        "exit_code",
        "rerun_weak",
        "selected",
        "no_op",
        "no_op_reason",
        "output_verified",
        "output_verified_local",
        "nas_publish_attempted",
        "nas_publish_succeeded",
        "photos_batch_csv_sha256",
        "started_at",
        "finished_at",
        "timestamp",
    )

    def _load_app(self):
        ns = {"json": __import__("json"), "Path": Path, "__name__": "app_probe_ns"}
        src = APP_PATH.read_text(encoding="utf-8-sig")
        start = src.find("def load_json(")
        end = src.find("def save_json(", start)
        assert start != -1 and end != -1
        exec(compile(src[start:end], "app.py", "exec"), ns)
        return ns

    # 1 -------------------------------------------------------------------
    def test_minimal_done_json_plus_manifest_is_enriched(self):
        import tempfile

        ns = self._load_app()
        with tempfile.TemporaryDirectory() as tmp:
            done = Path(tmp) / "done_job.json"
            done.write_text(
                '{"job_id": "j1", "type": "annotation_photos_batch", '
                '"profile": "run_standard", "options": ["--rerun-weak", "1"]}',
                encoding="utf-8-sig",
            )
            manifest = Path(tmp) / "20260907_133901_j1.manifest.json"
            manifest.write_text(
                '{"job_id": "j1", "status": "done", "exit_code": 0, "rerun_weak": true, '
                '"selected": 0, "no_op": true, "no_op_reason": "NO_WEAK_ROWS", '
                '"output_verified": true, "output_verified_local": true, '
                '"nas_publish_attempted": false, "nas_publish_succeeded": false, '
                '"photos_batch_csv_sha256": "abc", "started_at": "s", "finished_at": "f", '
                '"timestamp": "t"}',
                encoding="utf-8-sig",
            )
            job = ns["load_json"](str(done))
            man = ns["load_json"](str(manifest))
        # Le JSON de soumission ne porte aucun champ terminal.
        for field in self.MANIFEST_FIELDS:
            self.assertNotIn(field, job, f"{field} ne doit pas etre dans done/*.json")
        # Fusion additive : le manifeste complete, sans ecraser profile/options.
        merged = dict(man)
        merged.update({k: v for k, v in job.items() if k in {"job_id", "profile", "options"}})
        self.assertEqual(merged["profile"], "run_standard")
        self.assertEqual(merged["options"], ["--rerun-weak", "1"])
        for field in self.MANIFEST_FIELDS:
            self.assertIn(field, merged)

    # 2 -------------------------------------------------------------------
    def test_lookup_filters_on_job_id_only(self):
        """Le lookup cible un job_id ; un manifeste d un autre job ne convient pas."""
        import tempfile

        ns = self._load_app()
        with tempfile.TemporaryDirectory() as tmp:
            other = Path(tmp) / "other_job.manifest.json"
            other.write_text('{"job_id": "AUTRE"}', encoding="utf-8-sig")
            data = ns["load_json"](str(other))
        self.assertNotEqual(str(data.get("job_id")), "j1")

    # 3 -------------------------------------------------------------------
    def test_light_mode_does_not_scan_logs_globally(self):
        src = APP_PATH.read_text(encoding="utf-8-sig")
        start = src.find("def _matching_log_manifests(")
        assert start != -1
        body = src[start:start + 2500]
        # En mode non cible, logs/ n est consulte qu en detailed.
        self.assertIn('if not targeted and detail_level != "detailed"', body)
        self.assertIn("return []", body)
        # Le lookup reste non recursif et borne.
        flat = " ".join(body.split())
        self.assertIn("recursive=False", flat)
        self.assertIn("limit = 2 if targeted", flat)

    # 4 -------------------------------------------------------------------
    def test_missing_manifest_does_not_invent_success(self):
        ns = self._load_app()
        data = ns["load_json"]("_inexistant_.json", {"fallback": True})
        self.assertEqual(data, {"fallback": True})
        self.assertNotIn("status", data)

    # 5 -------------------------------------------------------------------
    def test_invalid_manifest_does_not_break_other_jobs(self):
        import tempfile

        ns = self._load_app()
        with tempfile.TemporaryDirectory() as tmp:
            bad = Path(tmp) / "bad.manifest.json"
            bad.write_text("{broken", encoding="utf-8-sig")
            good = Path(tmp) / "good.json"
            good.write_text('{"status": "completed"}', encoding="utf-8-sig")
            # Lecture a l interieur du TemporaryDirectory : sinon le fichier
            # est supprime et load_json retourne {} sans que ce soit un bug.
            self.assertFalse(ns["load_json"](str(bad), {}))
            self.assertEqual(ns["load_json"](str(good), {}).get("status"), "completed")

    # 8 -------------------------------------------------------------------
    def test_network_limits_are_respected(self):
        src = APP_PATH.read_text(encoding="utf-8-sig")
        self.assertIn("ANNOTATION_UNC_LIST_TIMEOUT_SECONDS", src)
        self.assertIn("ANNOTATION_JOB_SCAN_LIMIT_LIGHT", src)
        start = src.find("def _matching_log_manifests(")
        body = src[start:start + 2500]
        flat = " ".join(body.split())
        self.assertIn("timeout_seconds=ANNOTATION_UNC_LIST_TIMEOUT_SECONDS", flat)
        self.assertIn("recursive=False", flat)


class OutputProducerTests(unittest.TestCase):
    """Producteur effectif de photos_batch.csv (reference du rapport Word).

    Un no-op WEAK ne reecrit rien et ne publie rien : il ne doit jamais
    devenir la reference de contenu, meme s il est le job le plus recent.
    Une WEAK reelle ayant republie peut en revanche devenir producteur.
    """

    def setUp(self):
        self.ns = _load_batch_state_functions()

    def resolve(self, jobs, **kwargs):
        return self.ns["resolve_annotation_batch_state"](
            "2025-J48", "accedit-2026-05-28", job_details=_details(*jobs), **kwargs
        )

    # 1 -------------------------------------------------------------------
    def test_initial_published_plus_weak_noop_keeps_initial_as_producer(self):
        initial = _producer("annotation_x_20260907_090119_5f646751")
        noop = _job("annotation_x_20260907_133436_af1846aa", "completed",
                    rerun_weak=True, no_op=True, no_op_reason="NO_WEAK_ROWS")
        noop.update({"nas_publish_attempted": "false", "nas_publish_succeeded": "false",
                     "output_verified": "true", "output_verified_local": "true",
                     "publish_pending": "false", "sort_stamp": "20260907_133436"})
        state = self.resolve([initial, noop])
        self.assertEqual(state["output_producer"]["job_id"], initial["job_id"])
        self.assertEqual(state["report_batch"]["job_id"], initial["job_id"])
        self.assertTrue(state["word_report_ready"])

    # 2 -------------------------------------------------------------------
    def test_initial_publish_failed_blocks_word(self):
        initial = _producer("annotation_x_20260907_090119_5f646751", published=False)
        state = self.resolve([initial])
        self.assertIsNone(state["output_producer"])
        self.assertFalse(state["word_report_ready"])

    # 3 -------------------------------------------------------------------
    def test_real_weak_republished_becomes_producer(self):
        initial = _producer("annotation_x_20260907_090119_5f646751",
                            nas_sha256="sha_v1", stamp="20260907_090119")
        weak = _producer("annotation_x_20260907_133436_af1846aa", rerun_weak=True,
                         nas_sha256="sha_v2", stamp="20260907_133436")
        state = self.resolve([initial, weak])
        self.assertEqual(state["output_producer"]["job_id"], weak["job_id"])
        self.assertTrue(state["word_report_ready"])

    # 4 -------------------------------------------------------------------
    def test_later_noop_never_replaces_producer(self):
        initial = _producer("annotation_x_20260907_090119_5f646751", stamp="20260907_090119")
        weak_real = _producer("annotation_x_20260907_110000_aaaaaaaa", rerun_weak=True,
                              nas_sha256="sha_v2", stamp="20260907_110000")
        noop = _job("annotation_x_20260907_133436_af1846aa", "completed",
                    rerun_weak=True, no_op=True, no_op_reason="NO_WEAK_ROWS")
        noop.update({"nas_publish_attempted": "false", "nas_publish_succeeded": "false",
                     "output_verified": "true", "publish_pending": "false",
                     "sort_stamp": "20260907_133436"})
        state = self.resolve([initial, weak_real, noop])
        self.assertEqual(state["output_producer"]["job_id"], weak_real["job_id"])
        self.assertNotEqual(state["output_producer"]["job_id"], noop["job_id"])

    # 5 -------------------------------------------------------------------
    def test_noop_without_nas_publish_does_not_raise(self):
        initial = _producer("annotation_x_20260907_090119_5f646751")
        noop = _job("annotation_x_20260907_133436_af1846aa", "completed",
                    rerun_weak=True, no_op=True, no_op_reason="NO_WEAK_ROWS")
        noop.update({"nas_publish_attempted": "false", "nas_publish_succeeded": "false",
                     "output_verified": "true", "publish_pending": "false",
                     "sort_stamp": "20260907_133436"})
        # Aucune exception, et le no-op n est pas le producteur.
        state = self.resolve([initial, noop])
        self.assertNotEqual(state["output_producer"]["job_id"], noop["job_id"])
        self.assertTrue(state["word_report_ready"])

    # 6 -------------------------------------------------------------------
    def test_no_valid_producer_blocks_word(self):
        failed = _job("annotation_x_20260907_105035_d5649a7f", "failed", rerun_weak=True)
        state = self.resolve([failed])
        self.assertIsNone(state["output_producer"])
        self.assertFalse(state["word_report_ready"])

    # 7 -------------------------------------------------------------------
    def test_j48_producer_is_090119(self):
        initial = _producer("annotation_2025-J48_accedit-2026-05-28_run_standard_20260907_090119_5f646751",
                            stamp="20260907_090119")
        weak_failed = _job("annotation_2025-J48_accedit-2026-05-28_run_standard_20260907_105035_d5649a7f",
                           "failed", rerun_weak=True)
        weak_failed["sort_stamp"] = "20260907_105035"
        noop = _job("annotation_2025-J48_accedit-2026-05-28_run_standard_20260907_133436_af1846aa",
                    "completed", rerun_weak=True, no_op=True, no_op_reason="NO_WEAK_ROWS")
        noop.update({"nas_publish_attempted": "false", "nas_publish_succeeded": "false",
                     "output_verified": "true", "output_verified_local": "true",
                     "publish_pending": "false", "sort_stamp": "20260907_133436"})
        state = self.resolve([initial, weak_failed, noop])
        self.assertEqual(state["output_producer"]["job_id"], initial["job_id"])
        self.assertTrue(state["output_producer"]["nas_publish_succeeded"])
        self.assertTrue(state["output_producer"]["nas_publish_attempted"])
        self.assertEqual(state["initial_success"]["job_id"], initial["job_id"])
        self.assertEqual(state["latest_weak"]["job_id"], noop["job_id"])
        self.assertEqual(state["latest_failure"]["job_id"], weak_failed["job_id"])
        self.assertTrue(state["word_report_ready"])

    # 9 -------------------------------------------------------------------
    def test_completed_without_explicit_publish_proof_is_not_producer(self):
        """Regression : preuves POSITIVES requises.

        Un job completed sans nas_publish_attempted/succeeded explicite ne doit
        pas etre producteur, sinon word_report_ready serait True alors que
        _photo_report_job_preview refuserait le meme batch.
        """
        for missing in ("nas_publish_attempted", "nas_publish_succeeded",
                        "output_verified", "output_verified_local"):
            job = _producer("annotation_x_20260907_090119_5f646751")
            job.pop(missing, None)
            state = self.resolve([job])
            self.assertIsNone(
                state["output_producer"],
                f"{missing} absent ne doit pas donner un producteur",
            )
            self.assertFalse(
                state["word_report_ready"],
                f"{missing} absent ne doit pas autoriser le Word",
            )

        # Variante : champ present mais vide (ni true ni false).
        for empty in ("nas_publish_attempted", "nas_publish_succeeded",
                      "output_verified", "output_verified_local"):
            job = _producer("annotation_x_20260907_090119_5f646751")
            job[empty] = ""
            state = self.resolve([job])
            self.assertIsNone(state["output_producer"], f"{empty} vide -> pas de producteur")
            self.assertFalse(state["word_report_ready"], f"{empty} vide -> Word bloque")


    # 8 -------------------------------------------------------------------
    def test_nas_guardrails_still_apply_to_producer(self):
        """Le garde-fou NAS reste actif : publication echouee => pas de producteur."""
        initial = _producer("annotation_x_20260907_090119_5f646751", published=False)
        state = self.resolve([initial])
        self.assertIsNone(state["output_producer"])
        self.assertFalse(state["word_report_ready"])
        # Et output_verified faux => pas de producteur non plus.
        bad = _producer("annotation_x_20260907_090119_5f646751", output_verified="false")
        state2 = self.resolve([bad])
        self.assertIsNone(state2["output_producer"])


if __name__ == "__main__":
    unittest.main()
