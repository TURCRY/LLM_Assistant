from __future__ import annotations

import os
import socket
import time
from typing import Iterable
from urllib.parse import urlparse

import requests


DEFAULT_FLASK_ENDPOINTS = (
    "http://192.168.0.120:5050",
    "http://192.168.0.155:5050",
    "http://10.0.1.10:5050",
)

_CACHE_TTL_SECONDS = float(os.getenv("FLASK_ENDPOINT_CACHE_TTL", "60"))
_cached_url: str | None = None
_cached_until = 0.0


def _log(message: str) -> None:
    print(f"[server_locator] {message}")


def _normalize_url(url: str) -> str:
    url = str(url or "").strip().rstrip("/")
    if not url:
        return ""
    if "://" not in url:
        url = "http://" + url
    return url.rstrip("/")


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


def ping_endpoint(url: str, timeout: float = 1.5) -> bool:
    base_url = _normalize_url(url)
    if not base_url:
        return False
    try:
        response = requests.get(
            base_url + "/ping",
            headers={"Connection": "close"},
            timeout=timeout,
        )
        return bool(response.ok)
    except requests.RequestException as exc:
        _log(f"ping failed for {base_url}: {exc}")
        return False


def invalidate_flask_base_url(url: str | None = None) -> None:
    global _cached_url, _cached_until
    if url is None or url == _cached_url:
        _cached_url = None
        _cached_until = 0.0


def resolve_flask_base_url(
    *,
    force_refresh: bool = False,
    extra_candidates: Iterable[str] | None = None,
    timeout: float = 1.5,
) -> str:
    global _cached_url, _cached_until
    now = time.monotonic()
    if not force_refresh and _cached_url and now < _cached_until:
        return _cached_url

    candidates = _candidate_urls(extra_candidates)
    for url in candidates:
        if ping_endpoint(url, timeout=timeout):
            _cached_url = url
            _cached_until = now + _CACHE_TTL_SECONDS
            _log(f"selected Flask endpoint: {url}")
            return url

    fallback = candidates[0] if candidates else DEFAULT_FLASK_ENDPOINTS[0]
    _cached_url = fallback
    _cached_until = now + min(_CACHE_TTL_SECONDS, 5)
    _log(f"no endpoint answered /ping; fallback configured endpoint: {fallback}")
    return fallback


def _resolve_next_flask_base_url(
    failed_url: str,
    *,
    extra_candidates: Iterable[str] | None = None,
    timeout: float = 1.5,
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
