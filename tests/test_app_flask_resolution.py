from __future__ import annotations

import ast
import os
import unittest
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


def _load_app_flask_functions():
    tree = ast.parse(APP_PATH.read_text(encoding="utf-8-sig"))
    wanted_assignments = {
        "PCFIXE_FLASK_HOST_VPN",
        "PCFIXE_FLASK_HOST_LAN_PRIMARY",
        "PCFIXE_FLASK_HOST_LAN_FALLBACK",
    }
    wanted_functions = {
        "_sanitize_pcfixe_flask_host",
        "_pcfixe_flask_host_candidates",
        "_ordered_flask_candidate_urls",
        "detect_vpn_server_ip",
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
        "os": os,
        "_pcfixe_vpn_active": lambda: True,
        "_pcfixe_vpn_adapter": lambda *args, **kwargs: (True, "OpenVPN", "10.0.1.2"),
        "print": lambda *args, **kwargs: None,
    }
    exec(compile(module, str(APP_PATH), "exec"), ns)
    return ns


class AppFlaskResolutionTests(unittest.TestCase):
    def setUp(self):
        self.ns = _load_app_flask_functions()

    def test_vpn_candidates_prioritize_real_pcfixe_vpn_host(self):
        urls = self.ns["_ordered_flask_candidate_urls"]("5050", preferred_host="", vpn_active=True)
        self.assertEqual(
            urls,
            [
                "http://10.0.1.10:5050",
                "http://192.168.0.120:5050",
                "http://192.168.0.155:5050",
            ],
        )

    def test_obsolete_10015_is_removed_from_active_candidates(self):
        self.ns["PCFIXE_FLASK_HOST_VPN"] = "10.0.1.5"
        urls = self.ns["_ordered_flask_candidate_urls"]("5050", preferred_host="", vpn_active=True)
        self.assertNotIn("http://10.0.1.5:5050", urls)
        self.assertEqual(urls[0], "http://10.0.1.10:5050")

    def test_detect_vpn_server_ip_uses_configured_pcfixe_host_not_local_adapter_ip(self):
        result = self.ns["detect_vpn_server_ip"]("192.168.0.155")
        self.assertEqual(result, "10.0.1.10")


if __name__ == "__main__":
    unittest.main()
