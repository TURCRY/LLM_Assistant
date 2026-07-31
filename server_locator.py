from __future__ import annotations

import os
import socket
import time
import traceback
from pathlib import Path
from typing import Iterable
from urllib.parse import urlparse

import requests


DEFAULT_FLASK_ENDPOINTS = (
    "http://192.168.0.120:5050",
    "http://192.168.0.155:5050",
    "http://10.0.1.10:5050",
)

_CACHE_TTL_SECONDS = float(os.getenv("FLASK_ENDPOINT_CACHE_TTL", "20"))
_PING_CONNECT_TIMEOUT = float(os.getenv("FLASK_PING_CONNECT_TIMEOUT", "1.5"))
_PING_READ_TIMEOUT = float(os.getenv("FLASK_PING_READ_TIMEOUT", "5"))
_cached_url: str | None = None
_cached_until = 0.0
_last_probe_results: list[dict] = []
_resolution_stats = {
    "call_count": 0,
    "cache_hits": 0,
    "total_probe_ms": 0,
    "last_started_at": 0.0,
    "last_duration_ms": 0,
    "last_source": "",
    "last_selected_url": "",
    "last_stack": [],
    "last_probes": [],
}


def _log(message: str) -> None:
    print(f"[server_locator] {message}")


def _short_stack(limit: int = 4) -> list[str]:
    frames = traceback.extract_stack(limit=limit + 3)[:-2]
    return [f"{Path(frame.filename).name}:{frame.lineno}:{frame.name}" for frame in frames[-limit:]]


def _normalize_url(url: str) -> str:
    url = str(url or "").strip().rstrip("/")
    if not url:
        return ""
    if "://" not in url:
        url = "http://" + url
    return url.rstrip("/")


def get_flask_resolution_diagnostics() -> dict:
    return {
        "cached_url": _cached_url or "",
        "cached_until_monotonic": _cached_until,
        "cache_ttl_seconds": _CACHE_TTL_SECONDS,
        "call_count": _resolution_stats["call_count"],
        "cache_hits": _resolution_stats["cache_hits"],
        "total_probe_ms": _resolution_stats["total_probe_ms"],
        "last_started_at": _resolution_stats["last_started_at"],
        "last_duration_ms": _resolution_stats["last_duration_ms"],
        "last_source": _resolution_stats["last_source"],
        "last_selected_url": _resolution_stats["last_selected_url"],
        "last_stack": list(_resolution_stats["last_stack"]),
        "last_probes": [dict(item) for item in _resolution_stats["last_probes"]],
    }


def reset_flask_resolution_diagnostics() -> None:
    global _last_probe_results
    _last_probe_results = []
    _resolution_stats.update(
        {
            "call_count": 0,
            "cache_hits": 0,
            "total_probe_ms": 0,
            "last_started_at": 0.0,
            "last_duration_ms": 0,
            "last_source": "",
            "last_selected_url": "",
            "last_stack": [],
            "last_probes": [],
        }
    )


def _is_loopback(url: str) -> bool:
    host = (urlparse(url).hostname or "").lower()
    return host in {"127.0.0.1", "localhost", "::1"}


def _local_pcfixe_mode() -> bool:
    if os.getenv("ALLOW_LOCAL_FLASK", "0") == "1":
        return True
    if os.getenv("ON_PCFIXE", "0") == "1":
        return True
    hostname = (os.getenv("COMPUTERNAME") or os.getenv("HOSTNAME") or socket.gethostname() or "").upper()
    return "PCFIXE" in hostname


def _candidate_urls(extra_candidates: Iterable[str] | None = None) -> list[str]:
    values: list[str] = []
    for key in ("SERVER_URL", "LOCAL_LLM_BASE_URL", "FLASK_SERVER_URL"):
        values.append(os.getenv(key, ""))
    if extra_candidates:
        values.extend(extra_candidates)
    values.extend(DEFAULT_FLASK_ENDPOINTS)

    allow_loopback = _local_pcfixe_mode()
    seen: set[str] = set()
    urls: list[str] = []
    for raw in values:
        url = _normalize_url(raw)
        if not url or url in seen:
            continue
        if _is_loopback(url) and not allow_loopback:
            _log(f"ignore loopback endpoint outside explicit PC fixe mode: {url}")
            continue
        seen.add(url)
        urls.append(url)
    return urls


def _timeout_ms(timeout: float | tuple[float, float], index: int = 1) -> int:
    if isinstance(timeout, tuple):
        timeout = timeout[index]
    return int(float(timeout) * 1000)


def _is_connection_refused(exc: BaseException) -> bool:
    text = str(exc).lower()
    return "connection refused" in text or "winerror 10061" in text or "errno 111" in text


def ping_endpoint(url: str, timeout: float | tuple[float, float] | None = None) -> bool:
    return bool(probe_flask_endpoint(url, timeout=timeout).get("ok"))


def probe_flask_endpoint(url: str, timeout: float | tuple[float, float] | None = None) -> dict:
    base_url = _normalize_url(url)
    if not base_url:
        return {"url": "", "ok": False, "status": "Non testé", "detail": "URL vide", "elapsed_ms": None}
    effective_timeout = timeout if timeout is not None else (_PING_CONNECT_TIMEOUT, _PING_READ_TIMEOUT)
    started = time.monotonic()
    host = urlparse(base_url).hostname or base_url
    try:
        response = requests.get(
            base_url + "/ping",
            headers={"Connection": "close"},
            timeout=effective_timeout,
        )
        elapsed_ms = int((time.monotonic() - started) * 1000)
        if response.ok:
            _log(f"{host} repondu en {elapsed_ms} ms")
            return {
                "url": base_url,
                "ok": True,
                "status": "Disponible",
                "detail": f"HTTP {response.status_code}",
                "elapsed_ms": elapsed_ms,
            }
        try:
            response.raise_for_status()
        except requests.HTTPError as exc:
            _log(f"{host} HTTPError {response.status_code} apres {elapsed_ms} ms: {exc}")
        else:
            _log(f"{host} reponse HTTP non OK {response.status_code} apres {elapsed_ms} ms")
        return {
            "url": base_url,
            "ok": False,
            "status": "Indisponible",
            "detail": f"HTTPError {response.status_code}",
            "elapsed_ms": elapsed_ms,
        }
    except requests.ConnectTimeout as exc:
        _log(f"{host} ConnectTimeout apres {_timeout_ms(effective_timeout, 0)} ms: {exc}")
        return {"url": base_url, "ok": False, "status": "Délai dépassé", "detail": "ConnectTimeout", "elapsed_ms": _timeout_ms(effective_timeout, 0)}
    except requests.ReadTimeout as exc:
        _log(f"{host} ReadTimeout apres {_timeout_ms(effective_timeout)} ms: {exc}")
        return {"url": base_url, "ok": False, "status": "Délai dépassé", "detail": "ReadTimeout", "elapsed_ms": _timeout_ms(effective_timeout)}
    except requests.ConnectionError as exc:
        elapsed_ms = int((time.monotonic() - started) * 1000)
        kind = "ConnectionRefused" if _is_connection_refused(exc) else "ConnectionError"
        _log(f"{host} {kind} apres {elapsed_ms} ms: {exc}")
        return {"url": base_url, "ok": False, "status": "Indisponible", "detail": kind, "elapsed_ms": elapsed_ms}
    except requests.HTTPError as exc:
        elapsed_ms = int((time.monotonic() - started) * 1000)
        _log(f"{host} HTTPError apres {elapsed_ms} ms: {exc}")
        return {"url": base_url, "ok": False, "status": "Indisponible", "detail": f"HTTPError: {exc}", "elapsed_ms": elapsed_ms}
    except requests.RequestException as exc:
        elapsed_ms = int((time.monotonic() - started) * 1000)
        _log(f"{host} RequestException apres {elapsed_ms} ms: {exc}")
        return {"url": base_url, "ok": False, "status": "Indisponible", "detail": f"RequestException: {exc}", "elapsed_ms": elapsed_ms}


def probe_pcfixe_endpoints(
    endpoints: Iterable[str] | None = None,
    timeout: float | tuple[float, float] | None = None,
) -> list[dict]:
    probes = [probe_flask_endpoint(url, timeout=timeout) for url in (endpoints or DEFAULT_FLASK_ENDPOINTS)]
    reachable = [item["url"] for item in probes if item.get("ok")]
    if reachable:
        _log("[PCFIXE] endpoint joignable : " + ", ".join(reachable))
    else:
        _log("[PCFIXE] aucun endpoint joignable")
    return probes


def first_reachable_endpoint(probes: Iterable[dict], *, avoid_host: str | None = None) -> str:
    reachable = [item for item in probes if item.get("ok") and item.get("url")]
    if not reachable:
        return ""
    if avoid_host:
        for item in reachable:
            if (urlparse(item["url"]).hostname or "") != avoid_host:
                return item["url"]
    return reachable[0]["url"]


def _request_admin_json(method: str, base_url: str, path: str, api_key: str, timeout: float | tuple[float, float] = (1.5, 8)) -> dict:
    endpoint = _normalize_url(base_url)
    if not endpoint:
        return {"ok": False, "status": "Non testé", "error": "endpoint vide", "payload": None}
    try:
        response = requests.request(
            method,
            endpoint + path,
            headers={"x-api-key": api_key or "", "Connection": "close"},
            timeout=timeout,
        )
        status_code = response.status_code
        try:
            payload = response.json()
        except ValueError:
            payload = None
            if response.ok:
                return {"ok": False, "status": "JSON invalide", "status_code": status_code, "error": response.text[:500], "payload": None}
        if status_code in {401, 403}:
            return {"ok": False, "status": "Erreur d'authentification", "status_code": status_code, "payload": payload}
        if not response.ok:
            return {"ok": False, "status": "Erreur HTTP", "status_code": status_code, "payload": payload}
        return {"ok": True, "status": "Disponible", "status_code": status_code, "payload": payload}
    except requests.ConnectTimeout as exc:
        return {"ok": False, "status": "ConnectTimeout", "error": str(exc), "payload": None}
    except requests.ReadTimeout as exc:
        return {"ok": False, "status": "ReadTimeout", "error": str(exc), "payload": None}
    except requests.ConnectionError as exc:
        kind = "ConnectionRefused" if _is_connection_refused(exc) else "ConnectionError"
        return {"ok": False, "status": kind, "error": str(exc), "payload": None}
    except requests.RequestException as exc:
        return {"ok": False, "status": "RequestException", "error": str(exc), "payload": None}


def get_wifi_reconnect_status(base_url: str, api_key: str) -> dict:
    return _request_admin_json("GET", base_url, "/system/reconnect-wifi/status", api_key)


def trigger_wifi_reconnect(base_url: str, api_key: str) -> dict:
    result = _request_admin_json("POST", base_url, "/system/reconnect-wifi", api_key)
    if result.get("ok"):
        _log(f"[WIFI] tache de reconnexion declenchee via {_normalize_url(base_url)}")
    return result


def invalidate_flask_base_url(url: str | None = None) -> None:
    global _cached_url, _cached_until
    if url is None or url == _cached_url:
        _cached_url = None
        _cached_until = 0.0


def resolve_flask_base_url(
    *,
    force_refresh: bool = False,
    extra_candidates: Iterable[str] | None = None,
    timeout: float | tuple[float, float] | None = None,
) -> str:
    global _cached_url, _cached_until, _last_probe_results
    now = time.monotonic()
    started = time.monotonic()
    _resolution_stats["call_count"] += 1
    _resolution_stats["last_started_at"] = time.time()
    _resolution_stats["last_stack"] = _short_stack()
    if not force_refresh and _cached_url and now < _cached_until:
        _resolution_stats["cache_hits"] += 1
        _resolution_stats["last_duration_ms"] = int((time.monotonic() - started) * 1000)
        _resolution_stats["last_source"] = "cache"
        _resolution_stats["last_selected_url"] = _cached_url
        _resolution_stats["last_probes"] = [dict(item) for item in _last_probe_results]
        _log(f"cache hit #{_resolution_stats['call_count']} -> {_cached_url}")
        return _cached_url

    candidates = _candidate_urls(extra_candidates)
    probes: list[dict] = []
    for url in candidates:
        probe = probe_flask_endpoint(url, timeout=timeout)
        probes.append(probe)
        _resolution_stats["total_probe_ms"] += int(probe.get("elapsed_ms") or 0)
        if probe.get("ok"):
            _cached_url = url
            _cached_until = now + _CACHE_TTL_SECONDS
            _last_probe_results = [dict(item) for item in probes]
            _resolution_stats["last_duration_ms"] = int((time.monotonic() - started) * 1000)
            _resolution_stats["last_source"] = "probe"
            _resolution_stats["last_selected_url"] = url
            _resolution_stats["last_probes"] = [dict(item) for item in probes]
            _log(f"selected Flask endpoint: {url}")
            return url

    fallback = candidates[0] if candidates else _normalize_url(os.getenv("SERVER_URL", ""))
    if not fallback:
        fallback = _normalize_url(DEFAULT_FLASK_ENDPOINTS[0])
    _cached_url = fallback
    _cached_until = now + min(_CACHE_TTL_SECONDS, 5)
    _last_probe_results = [dict(item) for item in probes]
    _resolution_stats["last_duration_ms"] = int((time.monotonic() - started) * 1000)
    _resolution_stats["last_source"] = "fallback"
    _resolution_stats["last_selected_url"] = fallback
    _resolution_stats["last_probes"] = [dict(item) for item in probes]
    _log(f"aucun endpoint n'a repondu correctement a /ping; conservation du serveur configure/detecte: {fallback}")
    return fallback


def _resolve_next_flask_base_url(
    failed_url: str,
    *,
    extra_candidates: Iterable[str] | None = None,
    timeout: float | tuple[float, float] | None = None,
) -> str | None:
    global _cached_url, _cached_until
    for url in _candidate_urls(extra_candidates):
        if url == failed_url:
            continue
        if ping_endpoint(url, timeout=timeout):
            _cached_url = url
            _cached_until = time.monotonic() + _CACHE_TTL_SECONDS
            _log(f"selected fallback Flask endpoint: {url}")
            return url
    return None


def request_with_endpoint_fallback(
    method: str,
    path: str,
    *,
    retry_on_failure: bool = True,
    extra_candidates: Iterable[str] | None = None,
    **kwargs,
) -> requests.Response:
    base_url = resolve_flask_base_url(extra_candidates=extra_candidates)
    try:
        response = requests.request(method, base_url + path, **kwargs)
    except requests.RequestException:
        if not retry_on_failure:
            raise
        invalidate_flask_base_url(base_url)
        next_base_url = _resolve_next_flask_base_url(base_url, extra_candidates=extra_candidates)
        if not next_base_url:
            raise
        _log(f"retrying {method.upper()} {path} after network failure on: {next_base_url}")
        return requests.request(method, next_base_url + path, **kwargs)

    if not retry_on_failure or response.status_code < 500:
        return response

    invalidate_flask_base_url(base_url)
    next_base_url = _resolve_next_flask_base_url(base_url, extra_candidates=extra_candidates)
    if not next_base_url:
        return response
    _log(f"retrying {method.upper()} {path} on fallback endpoint: {next_base_url}")
    return requests.request(method, next_base_url + path, **kwargs)
