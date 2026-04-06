# -*- coding: utf-8 -*-
"""
create_affaire.py
- Crée une nouvelle affaire /Affaires/<AFFAIRE_ID>/ avec arborescence standard
- Initialise _Config/project_config.json
- Initialise _DB/project.sqlite (DDL minimal)
- Écrit un log dans AA_Expert_Admin/_Logs/create_affaire.log

Usage:
  python create_affaire.py --root "/volume1/Affaires" --id "2025-J37" --title "Affaire X c/ Y"
  python create_affaire.py --root "Z:\\Affaires" --id "2025-M25" --type "mediation"
"""

from __future__ import annotations

import argparse
import json
import os
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional


DDL_SQL = """
CREATE TABLE IF NOT EXISTS Documents(
  id_document TEXT PRIMARY KEY,
  numero_expert TEXT NOT NULL,
  code_partie TEXT NOT NULL,
  numero_avocat TEXT,
  description TEXT NOT NULL,
  date_reception TEXT NOT NULL,
  emetteur TEXT NOT NULL,
  pages INTEGER NOT NULL DEFAULT 0,
  a_annexer INTEGER NOT NULL DEFAULT 0,
  a_vectoriser INTEGER NOT NULL DEFAULT 0,
  est_dire INTEGER NOT NULL DEFAULT 0,
  uuid_paperless TEXT,
  chemin_nas TEXT NOT NULL,
  sha256 TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS Parties(
  code TEXT PRIMARY KEY,
  nom_affiche TEXT NOT NULL,
  type TEXT NOT NULL,
  aliases_json TEXT
);

CREATE TABLE IF NOT EXISTS Transcriptions(
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  fichier_audio TEXT NOT NULL,
  jsonl_path TEXT,
  csv_path TEXT,
  speakers_path TEXT,
  duree_s INTEGER,
  langue TEXT,
  quality REAL,
  a_annexer INTEGER NOT NULL DEFAULT 0,
  uuid_paperless TEXT
);

CREATE TABLE IF NOT EXISTS Journal(
  ts TEXT NOT NULL,
  evenement TEXT NOT NULL,
  meta_json TEXT,
  actor TEXT DEFAULT 'system'
);
"""

DEFAULT_PARTY_FOLDERS = [
    "00_Juridiction",
    *[f"{i:02d}_Partie_{i:02d}" for i in range(1, 41)],
    "99_Autre_Source",
]

DEFAULT_EXPERT_FOLDERS = [
    "AA_Expert_Admin",
    "AB_Expert_Dires",
    "AC_Expert_Rapports",
    "AD_Expert_Traitements",
    "AE_Expert_ASR",
    "AF_Exports",
    "AZ_Cloture",
]

DEFAULT_TECH_SUBFOLDERS = {
    "AA_Expert_Admin": ["Depot_initial", "_Paperless_Inbox", "_Logs"],
    "AD_Expert_Traitements": ["_Queue_OCR", "_OCR_Texte", "_Splits", "_CSV_RAG", "_Manifests"],
    "AE_Expert_ASR": ["audio_raw", "transcriptions"],
    "_DB": [],
    "_Config": [],
    "_Manifests": [],
}


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def write_log(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"[{now_iso()}] {message}\n")


def init_sqlite(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    try:
        conn.executescript(DDL_SQL)
        conn.commit()
    finally:
        conn.close()


def default_project_config(affaire_id: str, title: str, affaire_type: str) -> Dict:
    return {
        "id": affaire_id,
        "title": title,
        "type": affaire_type,
        "paths": {
            "paperless_inbox": "AA_Expert_Admin/_Paperless_Inbox",
            "queue_ocr": "AD_Expert_Traitements/_Queue_OCR",
            "splits": "AD_Expert_Traitements/_Splits",
            "csv_rag": "AD_Expert_Traitements/_CSV_RAG",
            "sqlite": "_DB/project.sqlite",
        },
        "numbering": {
            "expert_strategy": "global",   # ou "triplet"
            "last_global": 0,
            "prefix": "2",
            "last_suffix": "00"
        },
        "rag": {
            "enabled": True,
            "backend": "chroma",          # ou "qdrant"
            "collection": affaire_id
        },
        "security": {
            "api_key_required": True
        }
    }


def create_structure(root: Path, affaire_id: str, title: str, affaire_type: str) -> Path:
    affaire_root = root / affaire_id
    if affaire_root.exists():
        raise FileExistsError(f"Affaire existe déjà: {affaire_root}")

    # Dossiers principaux
    affaire_root.mkdir(parents=True, exist_ok=False)

    # Parties / juridiction
    for d in DEFAULT_PARTY_FOLDERS:
        (affaire_root / d).mkdir(parents=True, exist_ok=True)

    # Expert
    for d in DEFAULT_EXPERT_FOLDERS:
        (affaire_root / d).mkdir(parents=True, exist_ok=True)

    # Dossiers transverses
    (affaire_root / "_DB").mkdir(parents=True, exist_ok=True)
    (affaire_root / "_Config").mkdir(parents=True, exist_ok=True)
    (affaire_root / "_Manifests").mkdir(parents=True, exist_ok=True)

    # Sous-dossiers techniques
    for parent, subs in DEFAULT_TECH_SUBFOLDERS.items():
        base = affaire_root / parent
        for s in subs:
            (base / s).mkdir(parents=True, exist_ok=True)

    # Config JSON
    cfg = default_project_config(affaire_id, title, affaire_type)
    cfg_path = affaire_root / "_Config" / "project_config.json"
    cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    # Lexique ASR vide (à compléter)
    lex_path = affaire_root / "_Config" / "asr_lexique.json"
    if not lex_path.exists():
        lex_path.write_text(json.dumps({"parties": [], "termes_techniques": []}, ensure_ascii=False, indent=2),
                            encoding="utf-8")

    # SQLite
    db_path = affaire_root / "_DB" / "project.sqlite"
    init_sqlite(db_path)

    # Log
    log_path = affaire_root / "AA_Expert_Admin" / "_Logs" / "create_affaire.log"
    write_log(log_path, f"Création affaire {affaire_id} OK. Config: {cfg_path}. DB: {db_path}.")

    return affaire_root


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--root", required=True, help="Racine /Affaires (NAS monté localement). Ex: /volume1/Affaires ou Z:\\Affaires")
    p.add_argument("--id", required=True, help="Identifiant affaire. Ex: 2025-J37")
    p.add_argument("--title", default="", help="Titre libre")
    p.add_argument("--type", default="judiciaire", help="Type: judiciaire|mediation|amiable|sapiteur|autre")
    args = p.parse_args()

    root = Path(args.root)
    if not root.exists():
        raise FileNotFoundError(f"Racine introuvable: {root}")

    affaire_root = create_structure(root, args.id, args.title, args.type)
    print(f"OK: {affaire_root}")


if __name__ == "__main__":
    main()