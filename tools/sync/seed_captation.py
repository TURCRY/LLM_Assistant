# -*- coding: utf-8 -*-
"""
seed_captation.py (Laptop) — wrapper d'exécution du .bat de propagation "fichiers lourds"

But:
- Exécuter run_all_from_jpg_v3.bat depuis un dossier ...\\Photos\\JPG
- Forcer par défaut la destination NAS: \\192.168.1.20\\volume1\\Affaires
- Capturer stdout/stderr, code retour, et écrire un log JSON dans:
  \\192.168.1.20\\volume1\\Affaires\\<ID_AFFAIRE>\\AA_Expert_Admin\\_Logs\\seed_captation.jsonl

Usage (exemples):
  python seed_captation.py --affaire 2025-J46 --cwd "D:\\Captations\\Accedit 06 11 2025\\Photos\\JPG"
  python seed_captation.py --affaire 2025-J46 --cwd "%CD%"
  python seed_captation.py --affaire 2025-J46 --cwd "%CD%" --root-dst "\\\\192.168.0.155\\Affaires" --mode PCFIXE

Pré-requis:
- run_all_from_jpg_v3.bat dans C:\\LLM_Assistant\\tools\\sync\\
- Le .bat doit accepter: run_all_from_jpg_v3.bat <ID_AFFAIRE> [ROOT_DST] [MODE]
  (si vous n'avez pas encore modifié le .bat, adaptez la commande ci-dessous)
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Tuple


DEFAULT_ROOT_DST_NAS = r"\\192.168.1.20\volume1\Affaires"
DEFAULT_MODE = "NAS"
DEFAULT_BAT = r"C:\LLM_Assistant\tools\sync\run_all_from_jpg_v3.bat"
LOCAL_LOG_DIR = Path(r"C:\LLM_Assistant\tools\sync\logs")



def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def is_jpg_folder(cwd: Path) -> bool:
    return cwd.is_dir() and cwd.name.lower() == "jpg"


def safe_mkdir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def write_jsonl_line(path: Path, obj: Dict[str, Any]) -> None:
    safe_mkdir(path.parent)
    with path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")

def run_bat(
    bat_path: Path,
    affaire_id: str,
    root_dst: str,
    mode: str,
    cwd: Path,
    photos_csv_path: str,
    timeout_s: Optional[int] = None,
) -> Tuple[int, str, str]:

    """
    Exécute le .bat en fixant le working directory sur le dossier JPG.
    Retourne: (returncode, stdout, stderr)
    """
    if not bat_path.exists():
        raise FileNotFoundError(f".bat introuvable: {bat_path}")

    # Important: sous Windows, exécuter via cmd.exe /c
    cmd = [
        "cmd.exe",
        "/c",
        str(bat_path),
        affaire_id,
        root_dst,
        mode,
        photos_csv_path,
    ]

    proc = subprocess.run(
        cmd,
        cwd=str(cwd),
        capture_output=True,
        text=True,
        timeout=timeout_s,
        shell=False,
    )
    return proc.returncode, proc.stdout, proc.stderr


def default_local_log_path(affaire_id: str) -> Path:
    # Log sur la destination (NAS par défaut), dans l’affaire
    # \\...\Affaires\<ID>\AA_Expert_Admin\_Logs\seed_captation.jsonl
    return Path(r"C:\LLM_Assistant\tools\sync\logs") / f"seed_captation_{affaire_id}.jsonl"

def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--affaire", required=True, help="ID affaire (ex: 2025-J46)")
    p.add_argument(
        "--cwd",
        default=os.getcwd(),
        help="Dossier de travail: doit être ...\\Photos\\JPG (par défaut: dossier courant)",
    )
    p.add_argument("--bat", default=DEFAULT_BAT, help="Chemin du .bat à exécuter")
    p.add_argument(
        "--root-dst",
        default=DEFAULT_ROOT_DST_NAS,
        help=r"Destination UNC (NAS par défaut). Ex: \\192.168.1.20\volume1\Affaires",
    )
    p.add_argument(
        "--mode",
        default=DEFAULT_MODE,
        help="Mode indicatif (NAS|PCFIXE|...). Sert au log et (optionnel) au .bat.",
    )
    p.add_argument("--timeout", type=int, default=0, help="Timeout secondes (0 = pas de timeout)")
    p.add_argument(
        "--log-path",
        default="",
        help="Chemin log JSONL. Si vide: log dans AA_Expert_Admin/_Logs sur la destination.",
    )
    p.add_argument("--photos-csv-path", required=True, help="chemin vers le csv des photos")

    args = p.parse_args()

    affaire_id = args.affaire.strip()
    cwd = Path(args.cwd).resolve()
    
    photos_csv = Path(args.photos_csv_path).resolve()
    if not photos_csv.exists():
        raise SystemExit(f"ERREUR: --photos-csv-path introuvable: {photos_csv}")

    bat_path = Path(args.bat).resolve()
    root_dst = args.root_dst.strip()
    mode = args.mode.strip()
    timeout_s = None if args.timeout <= 0 else int(args.timeout)

    if not is_jpg_folder(cwd):
        raise SystemExit(
            f"ERREUR: --cwd doit pointer vers un dossier nommé 'JPG'. Reçu: {cwd}"
        )

    if mode.upper() not in {"NAS", "PCFIXE"}:
        raise SystemExit("ERREUR: --mode doit valoir NAS ou PCFIXE")
    mode = mode.upper()

    log_path = Path(args.log_path) if args.log_path else default_local_log_path(affaire_id)

    # Évènement "start"
    start_evt: Dict[str, Any] = {
        "ts": now_iso(),
        "event": "seed_captation_start",
        "affaire": affaire_id,
        "mode": mode,
        "root_dst": root_dst,
        "cwd": str(cwd),
        "bat": str(bat_path),
        "user": os.environ.get("USERNAME", ""),
        "host": os.environ.get("COMPUTERNAME", ""),
    }
    write_jsonl_line(log_path, start_evt)

    # Exécution
    rc = 999
    out = ""
    err = ""
    try:
        rc, out, err = run_bat(
            bat_path=bat_path,
            affaire_id=affaire_id,
            root_dst=root_dst,
            mode=mode,
            cwd=cwd,
            photos_csv_path=args.photos_csv_path,
            timeout_s=timeout_s,
        )

    except subprocess.TimeoutExpired as e:
        rc = 124
        out = (e.stdout or "")
        err = (e.stderr or "") + f"\n[TIMEOUT] {e}"
    except Exception as e:
        rc = 125
        err = f"[EXCEPTION] {type(e).__name__}: {e}"

    # Évènement "end"
    end_evt: Dict[str, Any] = {
        "ts": now_iso(),
        "event": "seed_captation_end",
        "affaire": affaire_id,
        "mode": mode,
        "root_dst": root_dst,
        "cwd": str(cwd),
        "bat": str(bat_path),
        "returncode": rc,
        # On logge stdout/stderr tronqués pour éviter des logs énormes.
        "stdout_tail": out[-8000:] if out else "",
        "stderr_tail": err[-8000:] if err else "",
    }
    write_jsonl_line(log_path, end_evt)

    # Sortie console (utile en direct)
    print(f"[{now_iso()}] seed_captation: affaire={affaire_id} mode={mode} rc={rc}")
    if err:
        print("---- stderr (tail) ----")
        print(err[-4000:])
    if out:
        print("---- stdout (tail) ----")
        print(out[-4000:])

    return int(rc)


if __name__ == "__main__":
    raise SystemExit(main())