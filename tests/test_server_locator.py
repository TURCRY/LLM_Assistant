import os
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import server_locator


def _fake_response(ok: bool, status_code: int = 200):
    def raise_for_status():
        if not ok:
            raise server_locator.requests.HTTPError(f"HTTP {status_code}")

    return SimpleNamespace(ok=ok, status_code=status_code, raise_for_status=raise_for_status)


class ServerLocatorTests(unittest.TestCase):
    def setUp(self):
        server_locator.invalidate_flask_base_url()
        server_locator.reset_flask_resolution_diagnostics()

    def tearDown(self):
        server_locator.invalidate_flask_base_url()
        server_locator.reset_flask_resolution_diagnostics()

    @patch.dict(os.environ, {"COMPUTERNAME": "LAPTOP"}, clear=True)
    def test_resolve_prefers_first_endpoint_that_answers_ping(self):
        calls = []

        def fake_get(url, **kwargs):
            calls.append(url)
            return _fake_response(url.startswith("http://192.168.0.155:5050/"), 200 if url.startswith("http://192.168.0.155:5050/") else 503)

        with patch("server_locator.requests.get", side_effect=fake_get):
            resolved = server_locator.resolve_flask_base_url(force_refresh=True)

        self.assertEqual(resolved, "http://192.168.0.155:5050")
        self.assertEqual(
            calls,
            [
                "http://192.168.0.120:5050/ping",
                "http://192.168.0.155:5050/ping",
            ],
        )

    @patch.dict(os.environ, {"COMPUTERNAME": "LAPTOP", "SERVER_URL": "http://127.0.0.1:5050"}, clear=True)
    def test_loopback_is_ignored_outside_pcfixe_mode(self):
        calls = []

        def fake_get(url, **kwargs):
            calls.append(url)
            return _fake_response(True)

        with patch("server_locator.requests.get", side_effect=fake_get):
            resolved = server_locator.resolve_flask_base_url(force_refresh=True)

        self.assertEqual(resolved, "http://192.168.0.120:5050")
        self.assertNotIn("http://127.0.0.1:5050/ping", calls)

    @patch.dict(os.environ, {"COMPUTERNAME": "LAPTOP"}, clear=True)
    def test_cache_reuses_working_endpoint(self):
        with patch("server_locator.requests.get", return_value=_fake_response(True)) as mock_get:
            first = server_locator.resolve_flask_base_url(force_refresh=True)
            second = server_locator.resolve_flask_base_url()

        self.assertEqual(first, "http://192.168.0.120:5050")
        self.assertEqual(second, first)
        self.assertEqual(mock_get.call_count, 1)
        diagnostics = server_locator.get_flask_resolution_diagnostics()
        self.assertEqual(diagnostics["call_count"], 2)
        self.assertEqual(diagnostics["cache_hits"], 1)
        self.assertEqual(diagnostics["last_source"], "cache")

    @patch.dict(os.environ, {"COMPUTERNAME": "LAPTOP"}, clear=True)
    def test_resolve_prefers_vpn_candidate_order_when_passed_explicitly(self):
        calls = []

        def fake_get(url, **kwargs):
            calls.append(url)
            return _fake_response(url.startswith("http://10.0.1.10:5050/"), 200 if url.startswith("http://10.0.1.10:5050/") else 503)

        with patch("server_locator.requests.get", side_effect=fake_get):
            resolved = server_locator.resolve_flask_base_url(
                force_refresh=True,
                extra_candidates=[
                    "http://10.0.1.10:5050",
                    "http://192.168.0.120:5050",
                    "http://192.168.0.155:5050",
                ],
            )

        self.assertEqual(resolved, "http://10.0.1.10:5050")
        self.assertEqual(calls, ["http://10.0.1.10:5050/ping"])

    @patch.dict(os.environ, {"COMPUTERNAME": "LAPTOP"}, clear=True)
    def test_manual_invalidate_reprobes_endpoint(self):
        with patch("server_locator.requests.get", return_value=_fake_response(True)) as mock_get:
            first = server_locator.resolve_flask_base_url(force_refresh=True)
            server_locator.invalidate_flask_base_url()
            second = server_locator.resolve_flask_base_url()

        self.assertEqual(first, second)
        self.assertEqual(mock_get.call_count, 2)

    @patch.dict(os.environ, {"COMPUTERNAME": "LAPTOP"}, clear=True)
    def test_cache_expiration_reprobes_endpoint(self):
        with patch("server_locator.requests.get", return_value=_fake_response(True)) as mock_get:
            first = server_locator.resolve_flask_base_url(force_refresh=True)
            server_locator._cached_until = 0.0
            second = server_locator.resolve_flask_base_url()

        self.assertEqual(first, second)
        self.assertEqual(mock_get.call_count, 2)

    @patch.dict(os.environ, {"COMPUTERNAME": "LAPTOP"}, clear=True)
    def test_request_retries_on_network_failure(self):
        def fake_get(url, **kwargs):
            return _fake_response(not url.startswith("http://192.168.0.120:5050/"), 200 if not url.startswith("http://192.168.0.120:5050/") else 503)

        calls = []

        def fake_request(method, url, **kwargs):
            calls.append(url)
            if url.startswith("http://192.168.0.155:5050/"):
                raise server_locator.requests.ConnectionError("down")
            return SimpleNamespace(status_code=200, ok=True)

        with patch("server_locator.requests.get", side_effect=fake_get):
            with patch("server_locator.requests.request", side_effect=fake_request):
                response = server_locator.request_with_endpoint_fallback("GET", "/models_index")

        self.assertTrue(response.ok)
        self.assertEqual(
            calls,
            [
                "http://192.168.0.155:5050/models_index",
                "http://10.0.1.10:5050/models_index",
            ],
        )


if __name__ == "__main__":
    unittest.main()
