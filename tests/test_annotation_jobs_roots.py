from __future__ import annotations

import ast
import os
import time
import unittest
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_annotation_network_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted_assignments = {
        "ANNOTATION_JOBS_SMB_TCP_TIMEOUT_SECONDS",
        "ANNOTATION_JOBS_FOLDERS",
    }
    wanted_functions = {
        "_annotation_cache_prune",
        "_annotation_perf_begin",
        "_annotation_perf_end",
        "_annotation_list_json_candidates",
        "_unc_host",
        "_annotation_pcfixe_job_hosts",
        "_annotation_tcp_port_open",
        "_annotation_jobs_roots",
        "_annotation_job_details",
    }
    body = []
    for node in tree.body:
        if isinstance(node, ast.Assign):
            names = {target.id for target in node.targets if isinstance(target, ast.Name)}
            if names & wanted_assignments:
                body.append(node)
        elif isinstance(node, ast.FunctionDef) and node.name in wanted_functions:
            body.append(node)
    module = ast.Module(body=body, type_ignores=[])
    ast.fix_missing_locations(module)
    ns = {
        "Path": Path,
        "os": os,
        "time": time,
        "traceback": __import__("traceback"),
        "PHOTO_BATCH_ACTIONS": {
            "initial": {
                "label": "initial",
                "status_label": "initial",
                "profile": "vlm_strict",
                "options": ["--reset-vlm", "1", "--vlm-strict", "1"],
            }
        },
        "PHOTO_BATCH_PUBLISH_RETRY_KEY": "publish_retry",
        "NAS_AFFAIRES_ROOT": Path(r"\\192.168.1.20\Affaires"),
        "AFFAIRES_ROOT": r"C:\Affaires",
        "PCFIXE_AFFAIRES_SHARE_CANDIDATES": [],
        "ANNOTATION_JOBS_DIAGNOSTIC_TTL_SECONDS": 45.0,
        "ANNOTATION_JOB_SCAN_LIMIT_LIGHT": 20,
        "ANNOTATION_JOB_SCAN_LIMIT_DETAILED": 50,
        "ANNOTATION_UNC_LIST_TIMEOUT_SECONDS": 4.0,
        "ANNOTATION_SLOW_BLOCK_SECONDS": 15.0,
        "_ANNOTATION_JOB_DETAILS_CACHE": {},
        "_pcfixe_vpn_active": lambda: False,
        "get_pcfixe_smb_host": lambda: "192.168.0.155",
        "_path_accessible_quick": lambda path: (True, ""),
        "_test_path_with_timeout": lambda path: (True, ""),
        "load_json": lambda path, default: default,
        "_annotation_file_profile": lambda path, csv_expected=False: {"sha256": "", "présent": "non"},
        "_annotation_latest_completed_batch": lambda details: ("", {}),
        "_annotation_has_newer_blocking_batch": lambda *args, **kwargs: (False, None),
    }
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns


class AnnotationJobsRootsTests(unittest.TestCase):
    def setUp(self):
        self.ns = _load_annotation_network_functions()

    def test_jobs_roots_stops_after_first_accessible_authoritative_host(self):
        calls = []

        def fake_probe(host, port):
            calls.append((host, port))
            return True, ""

        self.ns["_annotation_pcfixe_job_hosts"] = lambda: ["192.168.0.120", "192.168.0.155", "10.0.1.10"]
        self.ns["_annotation_tcp_port_open"] = fake_probe

        roots, probes = self.ns["_annotation_jobs_roots"]()

        self.assertEqual(calls, [("192.168.0.120", 445)])
        self.assertEqual(len(roots), 1)
        self.assertEqual(len(probes), 1)
        self.assertTrue(all(row["tcp_445"] == "oui" for row in probes if row["source"] != "local_cache"))

    def test_jobs_roots_reports_inaccessible_host_without_exception(self):
        calls = []

        def fake_probe(host, port):
            calls.append((host, port))
            return False, "TimeoutError: timed out"

        self.ns["_annotation_pcfixe_job_hosts"] = lambda: ["10.0.1.10"]
        self.ns["_annotation_tcp_port_open"] = fake_probe

        roots, probes = self.ns["_annotation_jobs_roots"]()

        self.assertEqual(calls, [("10.0.1.10", 445), ("192.168.1.20", 445)])
        self.assertEqual(roots, [{"source": "local_cache", "label": "cache laptop non autoritaire", "root": "C:\\Affaires\\_jobs", "priority": 3, "vpn_active": "non", "tcp_445": "local", "probe_ms": "0", "accessible": True, "folders": "queued, running, done", "detail": "accessible"}])
        self.assertEqual(probes[0]["detail"], "SMB 445 inaccessible (TimeoutError: timed out)")
        self.assertEqual(probes[1]["detail"], "SMB 445 inaccessible (TimeoutError: timed out)")

    def test_job_details_uses_roots_without_typeerror(self):
        calls = []

        def fake_roots(*args, **kwargs):
            calls.append("called")
            return ([], [{"source": "pcfixe", "detail": "inaccessible"}])

        self.ns["_annotation_jobs_roots"] = fake_roots

        details, diagnostics = self.ns["_annotation_job_details"](
            "2025-J47",
            "accedit-2025-11-13",
            paths=None,
            include_diagnostics=True,
        )

        self.assertEqual(calls, ["called"])
        self.assertIn("initial", details)
        self.assertEqual(diagnostics["roots"], [{"source": "pcfixe", "detail": "inaccessible"}])
        self.assertEqual(diagnostics["no_batch_reason"], "aucun registre _jobs fiable accessible")

    def test_annotation_smb_scan_exposes_timings_without_http_resolution(self):
        self.ns["resolve_flask_base_url"] = lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("HTTP resolver should not be called"))
        self.ns["request_with_endpoint_fallback"] = lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("HTTP fallback should not be called"))

        _, diagnostics = self.ns["_annotation_job_details"](
            "2025-J47",
            "accedit-2025-11-13",
            paths=None,
            include_diagnostics=True,
        )

        self.assertIn("timings", diagnostics)
        self.assertIn("jobs_root_probe_ms", diagnostics["timings"])
        self.assertIn("job_scan_ms", diagnostics["timings"])
        self.assertIn("manifest_read_ms", diagnostics["timings"])


if __name__ == "__main__":
    unittest.main()
