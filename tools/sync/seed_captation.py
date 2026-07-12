# -*- coding: utf-8 -*-
"""
Local abstraction for captation seeding.

The Streamlit UI calls this wrapper, not a versioned .bat directly.  The
default backend is the current operational chain:
run_all_from_jpg_v5.bat.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any


DEFAULT_ROOT_DST_NAS = r"\\192.168.1.20\Affaires"
DEFAULT_ROOT_PCFIXE = rf"\\{os.getenv('PCFIXE_SMB_HOST_VPN', '10.0.1.10')}\Affaires"
DEFAULT_MODE = "FULL"
BACKEND_LABEL = "run_all_from_jpg_v5"


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def is_jpg_folder(cwd: Path) -> bool:
    return cwd.is_dir() and cwd.name.lower() == "jpg"


def safe_mkdir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def write_jsonl_line(path: Path, obj: dict[str, Any]) -> None:
    safe_mkdir(path.parent)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")


def default_bat_candidates() -> list[Path]:
    here = Path(__file__).resolve().parent
    workspace = Path(r"C:\CodexWorkspace")
    return [
        Path(os.getenv("CAPTATION_SEED_BAT", "")) if os.getenv("CAPTATION_SEED_BAT") else None,
        here / "run_all_from_jpg_v5.bat",
        here / "run_all_from_JPG_v5.bat",
        here.parent / "compression_photos" / "run_all_from_jpg_v5.bat",
        Path(r"C:\LLM_Assistant\tools\sync\run_all_from_jpg_v5.bat"),
        Path(r"C:\LLM_Assistant\tools\sync\run_all_from_JPG_v5.bat"),
        Path(r"C:\LLM_Assistant\tools\compression_photos\run_all_from_jpg_v5.bat"),
        workspace
        / "_codex_context"
        / "gpt4all_local_context"
        / "scripts"
        / "copie_photos_laptop_vers_PCfixe"
        / "run_all_from_jpg_v5.bat",
        workspace
        / "_codex_context"
        / "gpt4all_local_context"
        / "docs"
        / "app_reference"
        / "synchronisation"
        / "laptop-naspcfixe"
        / "copie_gros_fichiers"
        / "run_all_from_jpg_v5.bat",
        workspace / "copie_gros_fichiers" / "codex_proposal" / "run_all_from_jpg_v5.bat",
    ]


def resolve_backend_bat(explicit_bat: str = "") -> Path:
    if explicit_bat.strip():
        path = Path(explicit_bat).expanduser()
        if path.exists():
            return path.resolve()
        raise FileNotFoundError(f".bat introuvable: {path}")

    candidates = [p for p in default_bat_candidates() if p is not None]
    for candidate in candidates:
        try:
            if candidate.exists():
                return candidate.resolve()
        except OSError:
            continue

    checked = "\n".join(f"- {p}" for p in candidates)
    raise FileNotFoundError(
        "Aucun backend de seed captation trouve. "
        "Definir CAPTATION_SEED_BAT ou installer run_all_from_jpg_v5.bat.\n"
        f"Chemins testes:\n{checked}"
    )


def default_local_log_path(affaire_id: str) -> Path:
    return Path(__file__).resolve().parent / "logs" / f"seed_captation_{affaire_id}.jsonl"


def build_command(bat_path: Path, affaire_id: str) -> list[str]:
    return ["cmd.exe", "/c", str(bat_path), affaire_id]


def validate_inputs(args: argparse.Namespace, bat_path: Path) -> dict[str, Any]:
    affaire_id = args.affaire.strip()
    cwd = Path(args.cwd).resolve()
    photos_csv = Path(args.photos_csv_path).resolve() if args.photos_csv_path else None

    errors: list[str] = []
    warnings: list[str] = []

    if not affaire_id:
        errors.append("--affaire est requis")
    if not is_jpg_folder(cwd):
        errors.append(f"--cwd doit pointer vers un dossier nomme 'JPG': {cwd}")
    if not bat_path.exists():
        errors.append(f"backend introuvable: {bat_path}")
    if photos_csv and not photos_csv.exists():
        warnings.append(f"--photos-csv-path introuvable: {photos_csv}")

    root_dst = args.root_dst.strip()
    root_pcfixe = args.root_pcfixe.strip()
    mode = args.mode.strip().upper()
    if root_dst and root_dst != DEFAULT_ROOT_DST_NAS:
        warnings.append(
            "--root-dst est conserve pour compatibilite UI, mais le backend v5 "
            "utilise ses racines operationnelles internes."
        )
    if not root_pcfixe:
        warnings.append("--root-pcfixe absent; le backend utilisera ses valeurs internes si elles existent.")
    if mode not in {"FULL", "NAS", "PCFIXE"}:
        warnings.append("--mode non standard; valeurs attendues: FULL, NAS, PCFIXE")
    elif mode in {"NAS", "PCFIXE"}:
        warnings.append(
            "--mode est informatif avec le backend v5: la chaine actuelle copie NAS puis PC fixe."
        )

    return {
        "ok": not errors,
        "errors": errors,
        "warnings": warnings,
        "affaire": affaire_id,
        "cwd": str(cwd),
        "photos_csv_path": str(photos_csv) if photos_csv else "",
        "backend_label": BACKEND_LABEL,
        "backend_bat": str(bat_path),
        "root_dst_ui": root_dst,
        "root_pcfixe_ui": root_pcfixe,
        "mode_ui": mode,
        "operational_roots": {
            "nas": root_dst or DEFAULT_ROOT_DST_NAS,
            "pcfixe": root_pcfixe or DEFAULT_ROOT_PCFIXE,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--affaire", required=True, help="ID affaire, ex: 2025-J46")
    parser.add_argument(
        "--cwd",
        default=os.getcwd(),
        help=r"Dossier de travail, normalement ...\photos\JPG",
    )
    parser.add_argument(
        "--bat",
        default="",
        help="Backend .bat explicite. Par defaut: resolution de run_all_from_jpg_v5.bat.",
    )
    parser.add_argument(
        "--root-dst",
        default=DEFAULT_ROOT_DST_NAS,
        help=r"Valeur UI historique. Backend v5: racines operationnelles internes.",
    )
    parser.add_argument(
        "--root-pcfixe",
        default=os.getenv("PCFIXE_AFFAIRES_UNC_ROOT", DEFAULT_ROOT_PCFIXE),
        help=r"Racine SMB du miroir PC fixe resolue par l'UI. Ex: \\10.0.1.10\Affaires",
    )
    parser.add_argument(
        "--mode",
        default=DEFAULT_MODE,
        help="Valeur UI historique/informative: FULL, NAS ou PCFIXE.",
    )
    parser.add_argument("--timeout", type=int, default=0, help="Timeout secondes (0 = aucun)")
    parser.add_argument("--log-path", default="", help="Chemin log JSONL optionnel")
    parser.add_argument("--photos-csv-path", default="", help="Chemin du CSV photos, optionnel")
    parser.add_argument("--dry-run", action="store_true", help="Verifie et affiche sans executer")

    args = parser.parse_args()
    bat_path = resolve_backend_bat(args.bat)
    preflight = validate_inputs(args, bat_path)
    command = build_command(bat_path, preflight["affaire"])
    preflight["command"] = command
    preflight["dry_run"] = bool(args.dry_run)

    log_path = Path(args.log_path) if args.log_path else default_local_log_path(preflight["affaire"])
    write_jsonl_line(log_path, {"ts": now_iso(), "event": "seed_captation_preflight", **preflight})

    print("=== Captation seed preflight ===")
    print(json.dumps(preflight, ensure_ascii=False, indent=2))

    if not preflight["ok"]:
        return 2

    if args.dry_run:
        print("DRY-RUN: aucune copie lancee.")
        return 0

    timeout_s = None if args.timeout <= 0 else int(args.timeout)
    start_evt = {
        "ts": now_iso(),
        "event": "seed_captation_start",
        "backend_label": BACKEND_LABEL,
        "backend_bat": str(bat_path),
        "command": command,
        "cwd": preflight["cwd"],
    }
    write_jsonl_line(log_path, start_evt)

    try:
        child_env = os.environ.copy()
        child_env["ROOT_DST"] = preflight["operational_roots"]["nas"]
        child_env["ROOT_DST_NAS"] = preflight["operational_roots"]["nas"]
        child_env["ROOT_PCFIXE"] = preflight["operational_roots"]["pcfixe"]
        child_env["PCFIXE_AFFAIRES_UNC_ROOT"] = preflight["operational_roots"]["pcfixe"]
        proc = subprocess.run(
            command,
            cwd=preflight["cwd"],
            capture_output=True,
            text=True,
            timeout=timeout_s,
            shell=False,
            env=child_env,
        )
        rc = proc.returncode
        out = proc.stdout or ""
        err = proc.stderr or ""
    except subprocess.TimeoutExpired as exc:
        rc = 124
        out = exc.stdout or ""
        err = (exc.stderr or "") + f"\n[TIMEOUT] {exc}"
    except Exception as exc:
        rc = 125
        out = ""
        err = f"[EXCEPTION] {type(exc).__name__}: {exc}"

    write_jsonl_line(
        log_path,
        {
            "ts": now_iso(),
            "event": "seed_captation_end",
            "backend_label": BACKEND_LABEL,
            "backend_bat": str(bat_path),
            "returncode": rc,
            "stdout_tail": out[-8000:] if out else "",
            "stderr_tail": err[-8000:] if err else "",
        },
    )

    print(f"[{now_iso()}] seed_captation: backend={BACKEND_LABEL} affaire={preflight['affaire']} rc={rc}")
    if err:
        print("---- stderr (tail) ----")
        print(err[-4000:])
    if out:
        print("---- stdout (tail) ----")
        print(out[-4000:])
    return int(rc)


if __name__ == "__main__":
    raise SystemExit(main())
