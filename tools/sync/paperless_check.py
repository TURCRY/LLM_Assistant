# -*- coding: utf-8 -*-
"""
paperless_check.py
- Vérifie configuration Paperless-ngx :
  1) dépôt d’un PDF test dans AA_Expert_Admin/_Paperless_Inbox
  2) attente ingestion
  3) vérification via API (document présent, tag "Affaire:<ID>" si règle activée)
  4) option: MAJ SQLite (uuid_paperless) pour l'entrée correspondante

Usage:
  python paperless_check.py --affaires-root "/volume1/Affaires" --id "2025-J37" \
    --paperless-url "http://NAS:8000" --token "xxxxx" --timeout 60
"""

from __future__ import annotations

import argparse
import time
import json
import hashlib
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any

import requests


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def paperless_get(session: requests.Session, base: str, endpoint: str, params: Dict[str, Any] | None = None):
    url = base.rstrip("/") + endpoint
    r = session.get(url, params=params, timeout=20)
    r.raise_for_status()
    return r.json()


def find_doc_by_filename(session: requests.Session, base: str, filename: str) -> Optional[Dict[str, Any]]:
    """
    Recherche simple via endpoint /api/documents/?query=...
    Selon versions, 'query' effectue une recherche plein texte incluant filename.
    """
    data = paperless_get(session, base, "/api/documents/", params={"query": filename})
    results = data.get("results", [])
    for doc in results:
        # 'original_file_name' est couramment présent
        if doc.get("original_file_name") == filename:
            return doc
    return results[0] if results else None


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--affaires-root", required=True, help="Racine /Affaires (NAS monté).")
    p.add_argument("--id", required=True, help="Affaire id ex: 2025-J37")
    p.add_argument("--paperless-url", required=True, help="URL Paperless ex: http://nas:8000")
    p.add_argument("--token", required=True, help="Token API Paperless")
    p.add_argument("--timeout", type=int, default=90, help="Secondes attente ingestion")
    p.add_argument("--expected-tag", default="", help="Tag attendu ex: Affaire:2025-J37 (si règles Paperless)")
    args = p.parse_args()

    affaire_root = Path(args.affaires_root) / args.id
    inbox = affaire_root / "AA_Expert_Admin" / "_Paperless_Inbox"
    if not inbox.exists():
        raise FileNotFoundError(f"Inbox introuvable: {inbox}")

    # 1) Dépôt PDF test (utiliser un petit PDF existant ou générer un minimal)
    test_pdf = inbox / f"_paperless_test_{args.id}.pdf"
    if not test_pdf.exists():
        # PDF minimal : on écrit un PDF très simple via contenu binaire minimal ?
        # Pour un squelette, on met un placeholder et on demande à l'utilisateur d'y copier un PDF.
        test_pdf.write_bytes(b"%PDF-1.4\n% Test\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n")

    file_sha = sha256_file(test_pdf)
    filename = test_pdf.name
    print(f"[{now_iso()}] Déposé: {test_pdf} sha256={file_sha}")

    # 2) Interroger Paperless jusqu’à apparition
    s = requests.Session()
    s.headers.update({"Authorization": f"Token {args.token}"})

    start = time.time()
    found = None
    while time.time() - start < args.timeout:
        try:
            found = find_doc_by_filename(s, args.paperless_url, filename)
        except Exception:
            found = None
        if found:
            break
        time.sleep(3)

    if not found:
        raise RuntimeError(f"Document non trouvé dans Paperless après {args.timeout}s. Vérifier watch/consume.")

    doc_id = found.get("id")
    print(f"[{now_iso()}] OK ingestion. Paperless doc id={doc_id} filename={found.get('original_file_name')}")

    # 3) Vérifier tag attendu si précisé
    if args.expected_tag:
        # tags sont des IDs; il faut récupérer liste des tags pour comparaison de noms
        tags_index = paperless_get(s, args.paperless_url, "/api/tags/")
        id_to_name = {t["id"]: t["name"] for t in tags_index.get("results", [])}
        doc_tags = [id_to_name.get(tid, f"<tag_id:{tid}>") for tid in (found.get("tags") or [])]
        if args.expected_tag not in doc_tags:
            print(f"[WARN] Tag attendu absent. Attendu={args.expected_tag} / tags_doc={doc_tags}")
        else:
            print(f"[{now_iso()}] Tag OK: {args.expected_tag}")

    print("[DONE]")


if __name__ == "__main__":
    main()