#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Estimation des paramètres de compression JPEG par appariement original/réduit.

Pour chaque image du dossier "JPG" (original), le script cherche un fichier du même nom
dans "JPG reduits". Il recompresse l'original sur une grille de qualités et compare
(PSNR + taille) avec le fichier réduit pour estimer la "clé" la plus proche :
- qualité (0–100)
- progressif (True/False)
Le sous-échantillonnage est fixé par défaut à 4:2:0 (valeur Pillow = 2), car vos CSV indiquent 2 2.

Sortie : un CSV "estimation_compression.csv" dans le dossier parent, et un résumé en console.

Dépendances : pip install pillow numpy
"""

import os
from pathlib import Path
import csv
import io
import math
from typing import Tuple, Optional, Dict, Any, List

import numpy as np
from PIL import Image, JpegImagePlugin

# ---------- CONFIGURABLE ----------
SRC_DIR = r"C:\Users\Utilisateur\Documents\NTU-Consult\4 - Dossiers\4 - Judiciaire\4.5 - Préventifs avec suivi\J04 - OPH Vallée Sud Clamart TA95\C - Photos Audios\Photos du 02 juillet 2024\JPG"
RED_DIR = r"C:\Users\Utilisateur\Documents\NTU-Consult\4 - Dossiers\4 - Judiciaire\4.5 - Préventifs avec suivi\J04 - OPH Vallée Sud Clamart TA95\C - Photos Audios\Photos du 02 juillet 2024\JPG reduits"

QUALITY_GRID: List[int] = list(range(60, 96, 2))  # 60,62,...,94
TEST_PROGRESSIVE = [True, False]                   # on teste les deux
ASSUMED_SUBSAMPLING = 2                            # 0=444,1=422,2=420
MAX_CANDIDATES = 2000                              # sécurité
# Pondération score: d'abord on minimise l'écart de taille, puis on maximise PSNR
SIZE_TOLERANCE_RATIO = 0.10                        # 10% autour de la taille cible

# ----------------------------------

def jpeg_is_progressive(img: Image.Image) -> Optional[bool]:
    # Essayer via info; sinon inspecter les marqueurs bruts
    try:
        if "progression" in img.info:
            return bool(img.info.get("progression"))
        if "progressive" in img.info:
            return bool(img.info.get("progressive"))
    except Exception:
        pass
    # Fallback binaire : SOF2 (0xFFC2) indique progressif, SOF0 (0xFFC0) baseline
    try:
        with open(img.fp.name, "rb") as f:
            data = f.read(4096)
        # Chercher marqueurs
        for i in range(len(data)-1):
            if data[i] == 0xFF and data[i+1] in (0xC0, 0xC1, 0xC2):
                return data[i+1] == 0xC2
    except Exception:
        pass
    return None

def to_array(im: Image.Image) -> np.ndarray:
    if im.mode != "RGB":
        im = im.convert("RGB")
    return np.asarray(im, dtype=np.float32)

def psnr(img1: np.ndarray, img2: np.ndarray) -> float:
    if img1.shape != img2.shape:
        return 0.0
    mse = np.mean((img1 - img2) ** 2, dtype=np.float64)
    if mse <= 1e-12:
        return 99.0
    PIXEL_MAX = 255.0
    return 20.0 * math.log10(PIXEL_MAX) - 10.0 * math.log10(mse)

def resize_like(im: Image.Image, target: Image.Image) -> Image.Image:
    if im.size == target.size:
        return im
    return im.resize(target.size, Image.LANCZOS)

def recompress_bytes(im: Image.Image, quality: int, progressive: bool, subsampling: Optional[int]) -> bytes:
    buf = io.BytesIO()
    params = {
        "format": "JPEG",
        "quality": int(quality),
        "optimize": True,
        "progressive": bool(progressive),
    }
    if subsampling is not None:
        params["subsampling"] = int(subsampling)
    # On ne passe pas d'EXIF/ICC pour comparaison "en contenu" uniquement
    im.save(buf, **params)
    return buf.getvalue()

def estimate_one_pair(src_path: Path, red_path: Path) -> Dict[str, Any]:
    # Charger
    with Image.open(src_path) as src_im, Image.open(red_path) as red_im:
        # Comparer à résolution de la réduite (le plus prudent)
        src_cmp = resize_like(src_im, red_im)
        src_arr_base = to_array(src_cmp)
        red_arr = to_array(red_im)
        target_size = red_path.stat().st_size
        red_prog = jpeg_is_progressive(red_im)

        best = {
            "quality": None,
            "progressive": None,
            "psnr": -1.0,
            "size": None,
            "size_diff_ratio": None,
        }

        # Première passe: on filtre les candidats par taille dans une tolérance
        candidates = []
        for q in QUALITY_GRID:
            for prog in TEST_PROGRESSIVE:
                jpg_bytes = recompress_bytes(src_cmp, q, prog, ASSUMED_SUBSAMPLING)
                s = len(jpg_bytes)
                size_diff_ratio = abs(s - target_size) / max(1, target_size)
                candidates.append((size_diff_ratio, q, prog, jpg_bytes, s))
        candidates.sort(key=lambda t: t[0])  # trier par écart de taille ascendant

        # Limiter aux meilleurs par taille (ex. 25 premiers) pour calcul PSNR
        top_candidates = candidates[:min(25, len(candidates))]

        for size_diff_ratio, q, prog, jpg_bytes, s in top_candidates:
            try:
                with Image.open(io.BytesIO(jpg_bytes)) as cand_im:
                    cand_arr = to_array(cand_im)
                cand_psnr = psnr(cand_arr, red_arr)
                # Critère: d'abord taille, puis PSNR
                if (best["quality"] is None or
                    size_diff_ratio < (best["size_diff_ratio"] if best["size_diff_ratio"] is not None else 1e9) or
                    (abs(size_diff_ratio - (best["size_diff_ratio"] or 0)) < 1e-6 and cand_psnr > best["psnr"])):
                    best.update({
                        "quality": q,
                        "progressive": prog,
                        "psnr": cand_psnr,
                        "size": s,
                        "size_diff_ratio": size_diff_ratio,
                    })
            except Exception:
                continue

        return {
            "file": src_path.name,
            "src_pixels": f"{src_im.size[0]}x{src_im.size[1]}",
            "red_pixels": f"{red_im.size[0]}x{red_im.size[1]}",
            "red_progressive": red_prog,
            "red_size_bytes": target_size,
            "est_quality": best["quality"],
            "est_progressive": best["progressive"],
            "est_psnr": round(best["psnr"], 2) if best["psnr"] is not None else None,
            "est_size_bytes": best["size"],
            "est_size_diff_pct": round(100.0 * (best["size_diff_ratio"] or 0), 2) if best["size_diff_ratio"] is not None else None,
            "assumed_subsampling": "4:2:0",
        }

def main() -> int:
    src_dir = Path(SRC_DIR)
    red_dir = Path(RED_DIR)
    if not src_dir.is_dir():
        print(f"ERREUR: dossier source introuvable: {src_dir}")
        return 1
    if not red_dir.is_dir():
        print(f"ERREUR: dossier réduits introuvable: {red_dir}")
        return 1

    # Appariement par nom de fichier (insensible à la casse)
    red_map = {p.name.lower(): p for p in red_dir.iterdir() if p.suffix.lower() in {'.jpg', '.jpeg'}}
    pairs = [(p, red_map.get(p.name.lower())) for p in src_dir.iterdir() if p.suffix.lower() in {'.jpg', '.jpeg'} and p.name.lower() in red_map]

    if not pairs:
        print("Aucun couple trouvé (vérifier les noms identiques).")
        return 0

    out_csv = src_dir.parent / "estimation_compression.csv"
    header = ["file","src_pixels","red_pixels","red_progressive","red_size_bytes","est_quality","est_progressive","est_psnr","est_size_bytes","est_size_diff_pct","assumed_subsampling"]

    rows = []
    for src_path, red_path in pairs:
        try:
            rows.append(estimate_one_pair(src_path, red_path))
            print(f"OK: {src_path.name}")
        except Exception as e:
            print(f"ERREUR sur {src_path.name}: {e}")

    # Ecriture CSV
    with open(out_csv, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=header, delimiter=";")
        w.writeheader()
        for r in rows:
            w.writerow(r)

    # Résumé
    qs = [r["est_quality"] for r in rows if isinstance(r.get("est_quality"), int)]
    if qs:
        med = sorted(qs)[len(qs)//2]
        print(f"Médiane qualité estimée: {med}")
    progs = [r["est_progressive"] for r in rows if r.get("est_progressive") is not None]
    if progs:
        print(f"Progressif (True/False) majoritaire: {max(set(progs), key=progs.count)}")
    print(f"CSV écrit: {out_csv}")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
