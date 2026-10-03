from __future__ import annotations

import ast
import csv
import tempfile
import unittest
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted = {
        "_read_semicolon_csv",
        "_annotation_photos_batch_business_counts",
        "_annotation_photos_batch_business_counts_for_job",
        "resolve_annotation_batch_state",
        "_annotation_classify_jobs",
        "_annotation_job_recency_key",
        "_annotation_parse_time_value",
        "_boolish_annotation_value",
        "_annotation_job_sort_value",
    }
    assignments = {"JOB_STATE_RANK", "WEAK_NO_OP_REASON"}
    body = []
    for node in tree.body:
        if isinstance(node, ast.Assign):
            names = {target.id for target in node.targets if isinstance(target, ast.Name)}
            if names & assignments:
                body.append(node)
        elif isinstance(node, ast.FunctionDef) and node.name in wanted:
            body.append(node)
    ns = {
        "Path": Path,
        "csv": csv,
        "re": __import__("re"),
        "datetime": __import__("datetime").datetime,
    }
    exec(compile(ast.Module(body=body, type_ignores=[]), str(APP_PATH), "exec"), ns)
    return ns


def _job(job_id: str, business_status: str, business_message: str) -> dict:
    return {
        "job_id": job_id,
        "status": "completed",
        "profile": "run_standard",
        "rerun_weak": "false",
        "no_op": "",
        "exit_code": "0",
        "dry_run": "false",
        "output_verified": "true",
        "output_verified_local": "true",
        "nas_publish_attempted": "true",
        "nas_publish_succeeded": "true",
        "publish_pending": "false",
        "business_status": business_status,
        "business_message": business_message,
        "options": [],
    }


class AnnotationBusinessCountsTests(unittest.TestCase):
    def setUp(self):
        self.ns = _load_functions()
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "photos_batch.csv"

    def tearDown(self):
        self.tmp.cleanup()

    def _write(self, rows):
        fields = [
            "batch_id", "vlm_status", "batch_status",
            "description_vlm_batch", "libelle_propose_batch",
            "commentaire_propose_batch",
        ]
        with self.path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=fields, delimiter=";")
            writer.writeheader()
            writer.writerows(rows)

    def _row(self, vlm_status, batch_status="OK", filled=True):
        return {
            "batch_id": "batch-1",
            "vlm_status": vlm_status,
            "batch_status": batch_status,
            "description_vlm_batch": "description" if filled else "",
            "libelle_propose_batch": "libelle" if filled else "",
            "commentaire_propose_batch": "commentaire" if filled else "",
        }

    def test_all_vlm_errors_are_invalid(self):
        self._write([self._row("ERR", "ERR_VLM_RESULT", False) for _ in range(23)])
        result = self.ns["_annotation_photos_batch_business_counts"](
            self.path, batch_id="batch-1"
        )
        self.assertEqual(result["business_status"], "invalid_vlm_total")
        self.assertEqual(result["business_message"], "23/23 VLM en erreur")
        self.assertEqual(result["selected"], 23)
        self.assertEqual(result["vlm_ok"], 0)
        self.assertEqual(result["vlm_error"], 23)
        self.assertEqual(result["description_ok"], 0)
        self.assertEqual(result["libelle_ok"], 0)
        self.assertEqual(result["commentaire_ok"], 0)

    def test_partial_vlm_success_is_partial(self):
        self._write([self._row("OK"), self._row("ERR", "ERR_VLM_RESULT", False)])
        result = self.ns["_annotation_photos_batch_business_counts"](
            self.path, batch_id="batch-1"
        )
        self.assertEqual(result["business_status"], "partial")
        self.assertEqual(result["vlm_ok"], 1)
        self.assertEqual(result["vlm_error"], 1)

    def test_all_vlm_success_is_success(self):
        self._write([self._row("OK") for _ in range(3)])
        result = self.ns["_annotation_photos_batch_business_counts"](
            self.path, batch_id="batch-1"
        )
        self.assertEqual(result["business_status"], "success")
        self.assertEqual(result["vlm_ok"], 3)
        self.assertEqual(result["vlm_error"], 0)

    def test_hash_mismatch_does_not_attribute_current_csv(self):
        self._write([self._row("OK")])
        result = self.ns["_annotation_photos_batch_business_counts_for_job"](
            self.path,
            current_hash="new-hash",
            job_hash="old-hash",
            batch_id="batch-1",
        )
        self.assertEqual(result["business_status"], "unknown")
        self.assertEqual(result["selected"], 0)
        self.assertEqual(result["vlm_ok"], 0)


class AnnotationInvalidBusinessStateTests(unittest.TestCase):
    def setUp(self):
        self.ns = _load_functions()

    def test_invalid_vlm_total_is_never_initial_or_output_producer(self):
        job = _job(
            "annotation_a_20261002_121731_ea57cea4",
            "invalid_vlm_total",
            "23/23 VLM en erreur",
        )
        state = self.ns["resolve_annotation_batch_state"](
            "2026-J54", "accedit-2026-04-16", job_details={"initial": job}
        )
        self.assertIsNone(state["initial_success"])
        self.assertIsNone(state["output_producer"])

    def test_invalid_vlm_total_blocks_word_report_explicitly(self):
        job = _job(
            "annotation_a_20261002_121731_ea57cea4",
            "invalid_vlm_total",
            "23/23 VLM en erreur",
        )
        state = self.ns["resolve_annotation_batch_state"](
            "2026-J54", "accedit-2026-04-16", job_details={"initial": job}
        )
        self.assertFalse(state["word_report_ready"])
        self.assertEqual(state["word_report_block_reason"], "23/23 VLM en erreur")


if __name__ == "__main__":
    unittest.main()
