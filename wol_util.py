# app/wol_util.py
from __future__ import annotations
import socket, struct, time, requests

def wake_on_lan(mac: str, broadcast_ip: str = "255.255.255.255", port: int = 9) -> None:
    mac_bytes = bytes.fromhex(mac.replace(":", "").replace("-", ""))
    if len(mac_bytes) != 6:
        raise ValueError(f"MAC invalide: {mac}")
    magic_packet = b"\xff" * 6 + mac_bytes * 16
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        sock.sendto(magic_packet, (broadcast_ip, port))

def is_server_up(base_url: str, ping_path: str = "/ping", timeout: float = 3.0) -> bool:
    try:
        r = requests.get(base_url.rstrip("/") + ping_path, timeout=timeout)
        return r.ok
    except Exception:
        return False

def wait_for_server(base_url: str,
                    ping_path: str = "/ping",
                    max_wait_sec: int = 90,
                    poll_interval_sec: int = 3,
                    timeout_single: float = 3.0) -> bool:
    deadline = time.time() + max_wait_sec
    while time.time() < deadline:
        if is_server_up(base_url, ping_path=ping_path, timeout=timeout_single):
            return True
        time.sleep(poll_interval_sec)
    return False
