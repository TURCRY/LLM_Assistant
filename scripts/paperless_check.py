# -*- coding: utf-8 -*-
"""
paperless_check.py
------------------
But:
- Déposer un PDF test dans AA_Expert_Admin/_Paperless_Inbox/<AFFAIRE_ID>/
- Attendre ingestion Paperless-ngx
- Vérifier présence du document via API
- Vérifier tag attendu (ex: "Affaire:2025-J37") si demandé
- Option: journaliser dans la SQLite projet (table Journal si elle existe)

Pré-requis:
  pip install requests reportlab

Notes:
- API Paperless-ngx: /api/documents/?query=... ; /api/tags/ ; auth Token via "Authorization: Token <token>".
- La surveillance récursive du consume est contrôlée côté Paperless par PAPERLESS_CONSUMER_RECURSIVE. (à vérifier sur votre instance)
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sqlite3
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def write_log(log_path: Path, msg: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"[{now_iso()}] {msg}\n")


def make_test_pdf(pdf_path: Path, affaire_id: str, marker: str) -> None:
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    w, h = A4
    c.setTitle(f"paperless_check {affaire_id}")
    c.drawString(72, h - 72, "Paperless check (test de consommation)")
    c.drawString(72, h - 96, f"Affaire: {affaire_id}")
    c.drawString(72, h - 120, f"Marker: {marker}")
    c.drawString(72, h - 144, f"Generated: {now_iso()}")
    c.showPage()
    c.save()


@dataclass
class PaperlessDoc:
    id: int
    original_file_name: str
    tags: List[int]


def paperless_get(session: requests.Session, base: str, endpoint: str, params: Dict[str, Any] | None = None) -> Dict[str, Any]:
    url = base.rstrip("/") + endpoint
    r = session.get(url, params=params, timeout=20)
    r.raise_for_status()
    return r.json()


def find_doc_by_filename(session: requests.Session, base: str, filename: str) -> Optional[PaperlessDoc]:
    """
    Recherche via /api/documents/?query=<filename>.
    """
    data = paperless_get(session, base, "/api/documents/", params={"query": filename})
    results = data.get("results") or []
    if not results:
        return None

    # Priorité: correspondance exacte sur original_file_name
    for d in results:
        if d.get("original_file_name") == filename:
            return PaperlessDoc(
                id=int(d["id"]),
                original_file_name=str(d.get("original_file_name") or ""),
                tags=list(d.get("tags") or []),
            )

    d0 = results[0]
    return PaperlessDoc(
        id=int(d0["id"]),
        original_file_name=str(d0.get("original_file_name") or ""),
        tags=list(d0.get("tags") or []),
    )


def fetch_all_tags(session: requests.Session, base: str) -> Dict[int, str]:
    """
    Récupère tous les tags (pagination DRF).
    """
    out: Dict[int, str] = {}
    endpoint = "/api/tags/"
    params: Dict[str, Any] | None = None

    while True:
        data = paperless_get(session, base, endpoint, params=params)
        for t in data.get("results") or []:
            try:
                out[int(t["id"])] = str(t.get("name") or "")
            except Exception:
                continue

        nxt = data.get("next")
        if not nxt:
            break

        # DRF "next" est une URL complète ; on en extrait le path + query minimalement.
        # On reste robuste sans dépendance externe:
        from urllib.parse import urlparse, parse_qs

        u = urlparse(nxt)
        endpoint = u.path
        q = parse_qs(u.query)
        params = {k: (v[0] if len(v) == 1 else v) for k, v in q.items()}

    return out


def sqlite_journal(db_path: Path, event: str, meta: Dict[str, Any]) -> Tuple[bool, str]:
    """
    Écrit dans table Journal si elle existe (sinon no-op).
    """
    if not db_path.exists():
        return False, f"SQLite absente: {db_path}"

    try:
        conn = sqlite3.connect(str(db_path))
        try:
            cur = conn.cursor()
            # Vérifie l'existence de la table Journal
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Journal';")
            if not cur.fetchone():
                return False, "Table Journal absente (pas d'écriture)."

            cur.execute(
                "INSERT INTO Journal(ts, evenement, meta_json, actor) VALUES(?, ?, ?, ?);",
                (now_iso(), event, json.dumps(meta, ensure_ascii=False), "paperless_check"),
            )
            conn.commit()
            return True, "Journal écrit."
        finally:
            conn.close()
    except Exception as e:
        return False, f"Erreur SQLite: {e}"


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--affaires-root", required=True, help="Racine /Affaires (NAS monté). Ex: Z:\\Affaires ou /volume1/Affaires")
    p.add_argument("--id", required=True, help="Affaire id ex: 2025-J37")
    p.add_argument("--paperless-url", required=True, help="URL Paperless ex: http://nas:8000")
    p.add_argument("--token", required=True, help="Token API Paperless")
    p.add_argument("--timeout", type=int, default=90, help="Secondes attente ingestion")
    p.add_argument("--poll", type=int, default=3, help="Secondes entre polls API")
    p.add_argument("--expected-tag", default="", help='Tag attendu ex: "Affaire:2025-J37" (optionnel)')
    p.add_argument("--sqlite-write", action="store_true", help="Tente d'écrire un évènement dans _DB/project.sqlite (table Journal si existante)")
    args = p.parse_args()

    affaire_id = args.id.strip()
    affaire_root = Path(args.affaires_root) / affaire_id

    inbox = affaire_root / "AA_Expert_Admin" / "_Paperless_Inbox"
    logs = affaire_root / "AA_Expert_Admin" / "_Logs"
    log_path = logs / "paperless_check.log"
    db_path = affaire_root / "_DB" / "project.sqlite"

    if not inbox.exists():
        raise FileNotFoundError(f"Inbox introuvable: {inbox}")

    marker = f"{affaire_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    pdf_name = f"_paperless_test_{marker}.pdf"
    pdf_path = inbox / pdf_name

    make_test_pdf(pdf_path, affaire_id, marker)
    file_sha = sha256_file(pdf_path)

    write_log(log_path, f"Déposé PDF test: {pdf_path.name} sha256={file_sha}")
    print(f"[{now_iso()}] Déposé: {pdf_path} sha256={file_sha}")

    s = requests.Session()
    s.headers.update({"Authorization": f"Token {args.token}"})

    start = time.time()
    found: Optional[PaperlessDoc] = None

    while time.time() - start < args.timeout:
        try:
            found = find_doc_by_filename(s, args.paperless_url, pdf_name)
        except Exception as e:
            write_log(log_path, f"[WARN] API query failed: {e}")
            found = None

        if found and found.original_file_name == pdf_name:
            break

        time.sleep(max(1, args.poll))

    if not found:
        write_log(log_path, f"[FAIL] Document non trouvé après {args.timeout}s: {pdf_name}")
        raise RuntimeError(f"Document non trouvé dans Paperless après {args.timeout}s. Vérifier consume + récursivité + droits.")

    write_log(log_path, f"[OK] Ingestion: doc_id={found.id} filename={found.original_file_name}")
    print(f"[{now_iso()}] OK ingestion. doc_id={found.id} filename={found.original_file_name}")

    # Vérification du tag si demandé
    tag_ok: Optional[bool] = None
    tag_names: List[str] = []
    if args.expected_tag:
        id_to_name = fetch_all_tags(s, args.paperless_url)
        tag_names = [id_to_name.get(tid, f"<tag_id:{tid}>") for tid in (found.tags or [])]
        tag_ok = args.expected_tag in tag_names

        if tag_ok:
            write_log(log_path, f"[OK] Tag présent: {args.expected_tag}")
            print(f"[{now_iso()}] Tag OK: {args.expected_tag}")
        else:
            write_log(log_path, f"[WARN] Tag absent: expected={args.expected_tag} got={tag_names}")
            print(f"[WARN] Tag attendu absent. expected={args.expected_tag} got={tag_names}")

    # Option SQLite (sans supposer votre schéma Documents)
    if args.sqlite_write:
        meta = {
            "affaire_id": affaire_id,
            "paperless_doc_id": found.id,
            "filename": found.original_file_name,
            "sha256": file_sha,
            "expected_tag": args.expected_tag,
            "tag_ok": tag_ok,
            "tag_names": tag_names,
        }
        ok, msg = sqlite_journal(db_path, "paperless_check", meta)
        write_log(log_path, f"[{'OK' if ok else 'INFO'}] SQLite: {msg}")
        print(f"[{now_iso()}] SQLite: {msg}")

    write_log(log_path, "[DONE]")
    print("[DONE]")


if __name__ == "__main__":
    main()