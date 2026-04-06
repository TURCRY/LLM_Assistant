# -*- coding: utf-8 -*-
"""
sync_activate.py
----------------
Gestion des affaires actives côté Laptop pour Syncthing.

Fonctions :
- maintient C:\\Affaires\\_active_affaires.list
- régénère C:\\Affaires\\.stignore
- lit, pour chaque affaire active, _Config\\project_config.json si présent
- applique les patterns sync/exclude_patterns et sync/include_light_patterns
- journalise dans C:\\Affaires\\_logs\\sync_activate.log

Architecture visée :
- Syncthing ne diffuse vers le laptop que les affaires actives
- les contenus lourds restent exclus
- certains contenus "légers" explicitement réouverts restent synchronisés

Usage :
  python sync_activate.py --root "C:\\Affaires" --activate 2025-J37
  python sync_activate.py --root "C:\\Affaires" --deactivate 2025-J53
  python sync_activate.py --root "C:\\Affaires" --rebuild
  python sync_activate.py --root "C:\\Affaires" --show
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Set

ACTIVE_FILE = "_active_affaires.list"
STIGNORE_FILE = ".stignore"
LOG_DIR = "_logs"
LOG_FILE = "sync_activate.log"

AFFAIRE_ID_RX = re.compile(r"^\d{4}-[A-Z]+\d+$", re.IGNORECASE)

# Fallbacks cohérents avec app.py si project_config.json est absent/incomplet
DEFAULT_EXCLUDE_PATTERNS = [
    "**/AD_Expert_Traitements/_Queue_OCR/**",
    "**/AD_Expert_Traitements/_Splits/**",
    "**/AD_Expert_Traitements/_OCR_Texte/**",
    "**/AD_Expert_Traitements/_CSV_RAG/**",
    "**/AE_Expert_captations/**/audio/**",
    "**/AE_Expert_captations/**/photos/JPG/**",
    "**/AE_Expert_captations/**/photos/RAW2/**",
    "**/CZ_Cloture/**",
]

DEFAULT_INCLUDE_LIGHT_PATTERNS = [
    "**/AA_Expert_Admin/**",
    "**/AA_Expert_Admin/_RAG_PC/**",
    "**/AB_Organisation_expertise/**",
    "**/AC_Journaux/**",
    "**/_Config/**",
    "**/_DB/**",
    "**/_Manifests/**",
    "**/AF_Expert_ASR/transcriptions/**",
    "**/AE_Expert_captations/**/photos/JPG reduit/**",
    "**/AE_Expert_captations/**/photos/*.csv",
    "**/AE_Expert_captations/**/photos/*.xls",
    "**/AE_Expert_captations/**/photos/*.xlsx",
    "**/BA_Pieces_de_expert/**",
    "**/BB_Préparation_livrables/**",
    "**/BC_Traitement_automatise_livrables/**",
    "**/BD_Etudes_diveres_Expert/**",
    "**/BE_Traitement_captations/**",
    "**/BF_Prefabrication_livrables/**",
    "**/BD_Exports/**",
]


@dataclass
class SyncRules:
    exclude_patterns: List[str]
    include_light_patterns: List[str]


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log_message(root: Path, message: str) -> None:
    log_path = root / LOG_DIR / LOG_FILE
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"[{now_iso()}] {message}\n")


def normalize_affaire_id(value: str) -> str:
    v = value.strip()
    if not v:
        raise ValueError("Identifiant affaire vide.")
    v = v.upper()
    if not AFFAIRE_ID_RX.match(v):
        raise ValueError(f"Identifiant affaire invalide : {value!r}")
    return v


def unique_sorted(items: Iterable[str]) -> List[str]:
    return sorted({x.strip() for x in items if x and x.strip()})


def read_active_list(root: Path) -> List[str]:
    path = root / ACTIVE_FILE
    if not path.exists():
        return []
    items = []
    for line in path.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        try:
            items.append(normalize_affaire_id(s))
        except ValueError:
            # On ignore les lignes invalides, mais elles seront journalisées au rebuild.
            continue
    return unique_sorted(items)


def write_active_list(root: Path, items: List[str]) -> None:
    path = root / ACTIVE_FILE
    normalized = unique_sorted(normalize_affaire_id(x) for x in items)
    content = "\n".join(normalized)
    if content:
        content += "\n"
    path.write_text(content, encoding="utf-8")


def load_project_sync_rules(root: Path, affaire_id: str) -> SyncRules:
    cfg_path = root / affaire_id / "_Config" / "project_config.json"
    if not cfg_path.exists():
        return SyncRules(
            exclude_patterns=DEFAULT_EXCLUDE_PATTERNS[:],
            include_light_patterns=DEFAULT_INCLUDE_LIGHT_PATTERNS[:],
        )

    try:
        data = json.loads(cfg_path.read_text(encoding="utf-8"))
        sync = data.get("sync") or {}
        excludes = sync.get("exclude_patterns") or DEFAULT_EXCLUDE_PATTERNS
        includes = sync.get("include_light_patterns") or DEFAULT_INCLUDE_LIGHT_PATTERNS

        excludes = [str(x).strip() for x in excludes if str(x).strip()]
        includes = [str(x).strip() for x in includes if str(x).strip()]

        return SyncRules(
            exclude_patterns=excludes or DEFAULT_EXCLUDE_PATTERNS[:],
            include_light_patterns=includes or DEFAULT_INCLUDE_LIGHT_PATTERNS[:],
        )
    except Exception as e:
        log_message(root, f"[WARN] Lecture config impossible pour {affaire_id}: {e}")
        return SyncRules(
            exclude_patterns=DEFAULT_EXCLUDE_PATTERNS[:],
            include_light_patterns=DEFAULT_INCLUDE_LIGHT_PATTERNS[:],
        )


def _strip_any_leading_glob(pattern: str) -> str:
    """
    Convertit par ex.:
      **/AD_Expert_Traitements/_Queue_OCR/** -> AD_Expert_Traitements/_Queue_OCR/**
      AD_Expert_Traitements/**              -> AD_Expert_Traitements/**
    """
    p = pattern.replace("\\", "/").strip()
    while p.startswith("**/"):
        p = p[3:]
    while p.startswith("./"):
        p = p[2:]
    while p.startswith("/"):
        p = p[1:]
    return p


def affix_pattern_to_affaire(affaire_id: str, pattern: str) -> str:
    """
    Transforme un pattern 'global affaire' en pattern relatif à la racine Syncthing.
    Exemple:
      affaire_id=2025-J37
      pattern=**/CZ_Cloture/**
      -> 2025-J37/CZ_Cloture/**
    """
    p = _strip_any_leading_glob(pattern)
    if not p:
        return f"{affaire_id}/**"
    return f"{affaire_id}/{p}"


def build_stignore_lines(root: Path, active_affaires: List[str]) -> List[str]:
    """
    Stratégie:
    1) tout ignorer
    2) rouvrir uniquement les affaires actives
    3) réappliquer les exclusions 'lourdes'
    4) réouvrir explicitement certains contenus légers
    """
    lines: List[str] = []

    lines.append("# AUTOGENERATED - do not edit manually")
    lines.append(f"# Generated: {now_iso()}")
    lines.append("")

    # Tout ignorer par défaut
    lines.append("**")
    lines.append("")

    # Toujours garder les fichiers de pilotage à la racine
    lines.append(f"!/{ACTIVE_FILE}")
    lines.append(f"!/{STIGNORE_FILE}")
    lines.append(f"!/{LOG_DIR}/")
    lines.append(f"!/{LOG_DIR}/**")
    lines.append("")

    for affaire_id in active_affaires:
        # Réouvrir l'affaire
        lines.append(f"!/{affaire_id}/")
        lines.append(f"!/{affaire_id}/**")

        rules = load_project_sync_rules(root, affaire_id)

        # Exclure les lourds
        for pattern in unique_sorted(rules.exclude_patterns):
            lines.append("/" + affix_pattern_to_affaire(affaire_id, pattern))

        # Réouvrir explicitement certains légers
        for pattern in unique_sorted(rules.include_light_patterns):
            lines.append("!/" + affix_pattern_to_affaire(affaire_id, pattern))

        lines.append("")

    return lines


def write_stignore(root: Path, lines: List[str]) -> Path:
    path = root / STIGNORE_FILE
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    return path


def activate_affaires(current: List[str], to_add: Iterable[str]) -> List[str]:
    out: Set[str] = set(current)
    for item in to_add:
        out.add(normalize_affaire_id(item))
    return unique_sorted(out)


def deactivate_affaires(current: List[str], to_remove: Iterable[str]) -> List[str]:
    remove = {normalize_affaire_id(x) for x in to_remove}
    return unique_sorted(x for x in current if x not in remove)


def show_state(root: Path, active: List[str]) -> None:
    print("Affaires actives :")
    if not active:
        print("  (aucune)")
        return
    for a in active:
        cfg = root / a / "_Config" / "project_config.json"
        marker = "config OK" if cfg.exists() else "config absente"
        print(f"  - {a} [{marker}]")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True, help=r"Racine Syncthing Laptop. Ex: C:\Affaires")
    parser.add_argument("--activate", nargs="*", default=[], help="Affaires à activer")
    parser.add_argument("--deactivate", nargs="*", default=[], help="Affaires à désactiver")
    parser.add_argument("--rebuild", action="store_true", help="Reconstruit .stignore à partir de la liste active")
    parser.add_argument("--show", action="store_true", help="Affiche l'état courant")
    args = parser.parse_args()

    root = Path(args.root)
    root.mkdir(parents=True, exist_ok=True)

    current = read_active_list(root)
    updated = current[:]

    if args.activate:
        updated = activate_affaires(updated, args.activate)

    if args.deactivate:
        updated = deactivate_affaires(updated, args.deactivate)

    if args.activate or args.deactivate or args.rebuild:
        write_active_list(root, updated)
        lines = build_stignore_lines(root, updated)
        stignore_path = write_stignore(root, lines)

        log_message(
            root,
            f"[OK] rebuild stignore | active={updated} | stignore={stignore_path}"
        )

    if args.show or (not args.activate and not args.deactivate and not args.rebuild):
        show_state(root, updated)

    if args.activate or args.deactivate or args.rebuild:
        print("OK")
        print(f"Active: {updated}")
        print(f"Written: {root / ACTIVE_FILE}")
        print(f"Written: {root / STIGNORE_FILE}")


if __name__ == "__main__":
    main()