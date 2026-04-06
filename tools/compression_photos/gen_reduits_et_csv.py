#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génération d'images réduites + CSV "format guide".

Fonctions :
- Parcourt un dossier source (JPG/JPEG).
- Produit un sous-dossier voisin (par défaut "JPG reduits") avec les images compressées.
- Préserve l'EXIF (DateTimeOriginal, Orientation, etc.) et le profil ICC si présents.
- Génère un CSV "photos_reduits.csv" (délimiteur ";") dans le dossier parent
  avec les colonnes :
  nom_fichier_image;horodatage_photo;orientation_photo;horodatage_secondes;t_audio;decalage_individuel;synchro_audio;decalage_moyen

Dépendances :
    pip install pillow piexif

Testé sous Python ≥ 3.9.
"""
from __future__ import annotations
import csv
import sys
from pathlib import Path
from typing import Optional, Tuple

from PIL import Image, ExifTags
try:
    import piexif  # type: ignore
except Exception:
    piexif = None  # Fallback sans piexif

# =====================
# CONFIGURATION UTILISATEUR
# =====================
# Renseignez le dossier source contenant les JPG/JPEG (vos originaux)
SRC_DIR = r"C:\\Users\\Utilisateur\\Documents\\NTU-Consult\\4 - Dossiers\\4 - Judiciaire\\4.5 - Préventifs avec suivi\\J04 - OPH Vallée Sud Clamart TA95\\C - Photos Audios\\Photos du 02 juillet 2024\\JPG"

# Nom du sous-dossier (créé à côté du dossier source) où seront écrites les images réduites
OUT_DIR_NAME = "JPG reduits"

# Nom du CSV à générer dans le dossier parent du SRC_DIR
CSV_NAME = "photos_reduits.csv"

# Paramètres de compression (modifiables)
QUALITY: int = 82              # Valeur classique robuste (80–85). Pour reproduire votre lot actuel, mettez 62–64.
SUBSAMPLING: Optional[int] = 2 # 0=4:4:4, 1=4:2:2, 2=4:2:0 ; None = laisser Pillow décider
PROGRESSIVE: bool = True       # JPEG progressif conseillé

# =====================
# FONCTIONS UTILITAIRES
# =====================

def ensure_rgb(im: Image.Image) -> Image.Image:
    """Convertit en RGB si nécessaire (JPEG ne gère pas tous les modes)."""
    return im if im.mode in ("RGB", "L") else im.convert("RGB")


def extract_datetime_orientation(im: Image.Image) -> Tuple[str, str]:
    """Extrait DateTimeOriginal et Orientation dans des formats adaptés au CSV guide."""
    horod = ""
    orient_str = ""

    exif = im.getexif()
    if exif:
        dto_tag = next((k for k, v in ExifTags.TAGS.items() if v == "DateTimeOriginal"), None)
        ori_tag = next((k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None)

        if dto_tag and dto_tag in exif:
            try:
                val = exif.get(dto_tag)
                if isinstance(val, bytes):
                    val = val.decode("utf-8", errors="ignore")
                if isinstance(val, str) and len(val) >= 10:
                    # Normaliser "YYYY:MM:DD HH:MM:SS" -> "YYYY-MM-DD HH:MM:SS" (remplacer seulement les 2 premiers ":")
                    horod = val.replace(":", "-", 2)
            except Exception:
                pass

        if ori_tag and ori_tag in exif:
            try:
                orient_str = str(exif.get(ori_tag))
            except Exception:
                pass

    return horod, orient_str


def save_with_metadata(im: Image.Image, dest: Path, quality: int, subs: Optional[int], progressive: bool) -> None:
    """Sauvegarde JPEG en conservant EXIF/ICC si disponibles."""
    info = im.info.copy()
    params = {
        "format": "JPEG",
        "quality": int(quality),
        "optimize": True,
        "progressive": bool(progressive),
    }
    if subs is not None:
        params["subsampling"] = int(subs)

    # Conserver ICC s'il existe
    icc = info.get("icc_profile")
    if icc:
        params["icc_profile"] = icc

    # Conserver EXIF (priorité à piexif si dispo)
    exif_bytes = info.get("exif")
    if exif_bytes and piexif is not None:
        try:
            exif_dict = piexif.load(exif_bytes)
            exif_bytes = piexif.dump(exif_dict)
        except Exception:
            pass
    if exif_bytes:
        params["exif"] = exif_bytes

    im.save(dest, **params)


# =====================
# PROGRAMME PRINCIPAL
# =====================

def main() -> int:
    src = Path(SRC_DIR)
    if not src.is_dir():
        print(f"ERREUR: dossier source introuvable: {src}")
        return 1

    parent = src.parent
    out_dir = parent / OUT_DIR_NAME
    out_dir.mkdir(parents=True, exist_ok=True)

    csv_path = parent / CSV_NAME
    header = [
        "nom_fichier_image",
        "horodatage_photo",
        "orientation_photo",
        "horodatage_secondes",
        "t_audio",
        "decalage_individuel",
        "synchro_audio",
        "decalage_moyen",
    ]

    images = sorted([p for p in src.iterdir() if p.suffix.lower() in {".jpg", ".jpeg"}])
    if not images:
        print("Aucune image .jpg/.jpeg trouvée dans le dossier source.")
        return 0

    with open(csv_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(header)

        print(f"{len(images)} images détectées. Écriture vers: {out_dir}")
        print(f"Paramètres: quality={QUALITY}, subsampling={SUBSAMPLING}, progressive={PROGRESSIVE}")

        for p in images:
            try:
                with Image.open(p) as im:
                    horod, orient = extract_datetime_orientation(im)
                    im2 = ensure_rgb(im)
                    dest = out_dir / p.name
                    save_with_metadata(im2, dest, QUALITY, SUBSAMPLING, PROGRESSIVE)

                    # Ligne CSV selon le modèle (les 4 dernières colonnes vides)
                    writer.writerow([
                        p.name,
                        horod,
                        orient,
                        "",
                        "",
                        "",
                        "",
                        "",
                    ])
            except Exception as e:
                print(f"ERREUR sur {p.name}: {e}")

    print(f"CSV généré: {csv_path}")
    print("Terminé.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
