#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script à lancer DEPUIS le dossier JPG\

Chaîne réalisée :
1) Vérifie que le script est lancé depuis un dossier nommé "JPG" (insensible à la casse)
2) Crée le sous-dossier frère "JPG reduit" s'il n'existe pas
3) Traite les fichiers .jpg/.jpeg, .heic/.heif, .mp4/.mov :
   - .jpg/.jpeg : crée la version réduite dans "JPG reduit"
   - .heic/.heif : crée un JPG de substitution dans "JPG", puis une version réduite
   - .mp4/.mov : extrait la première frame en JPG dans "JPG", puis crée une version réduite
   - conserve EXIF/ICC pour les JPG/JPEG lorsque possible
4) Génère dans le dossier parent un CSV "photos.csv" (séparateur ";", UTF-8-SIG)
5) Génère "photos_batch.csv"
6) Tente de produire "photos.xls" identique via Excel COM si disponible

Dépendances :
    pip install pillow piexif pillow-heif imageio
    # Optionnel pour XLS via Excel
    pip install pywin32

Utilisation :
    Ouvrir PowerShell dans le dossier JPG\ puis :
        python gen_tout_depuis_JPG.py
"""

import os
import csv
import sys
from pathlib import Path
from typing import Optional, Tuple

from PIL import Image
from pillow_heif import register_heif_opener
import imageio.v3 as iio

try:
    import piexif
except Exception:
    piexif = None

register_heif_opener()

# ========= PARAMÈTRES =========
QUALITY: int = 63
SUBSAMPLING: Optional[int] = 2
PROGRESSIVE: bool = True
OUT_DIR_NAME = "JPG reduit"
CSV_NAME = "photos.csv"
XLS_NAME = "photos.xls"
# ==============================


def ensure_rgb(im: Image.Image) -> Image.Image:
    return im if im.mode in ("RGB", "L") else im.convert("RGB")


def extract_datetime_orientation(im: Image.Image) -> Tuple[str, str]:
    """Retourne (horodatage_photo, orientation_photo)."""
    def _norm_dt_fr(s: str) -> str:
        if not isinstance(s, str):
            return ""
        s = s.strip()
        if len(s) < 10:
            return ""
        y = s[0:4]
        m = s[5:7]
        d = s[8:10]
        t = s[11:19] if len(s) >= 19 else "00:00:00"
        return f"{d}/{m}/{y} {t}"

    horod = ""
    orient_str = ""
    try:
        ex = im.getexif()
        if ex:
            for tag_id in (36867, 36868, 306):
                if tag_id in ex and not horod:
                    val = ex.get(tag_id)
                    if isinstance(val, bytes):
                        try:
                            val = val.decode("utf-8", errors="ignore")
                        except Exception:
                            val = ""
                    if isinstance(val, str) and len(val) >= 10:
                        horod = _norm_dt_fr(val)
                        break
            ori_tag = 274
            if ori_tag in ex:
                try:
                    code = int(ex.get(ori_tag))
                except Exception:
                    code = None
                mapping_deg = {1: 0, 2: 0, 3: 180, 4: 180, 5: 90, 6: 90, 7: 270, 8: 270}
                if code in mapping_deg:
                    orient_str = str(mapping_deg[code])
    except Exception:
        pass
    return horod, orient_str


def save_with_metadata(im: Image.Image, dest: Path, quality: int, subs: Optional[int], progressive: bool) -> None:
    info = im.info.copy()
    params = {
        "format": "JPEG",
        "quality": int(quality),
        "optimize": True,
        "progressive": bool(progressive),
    }
    if subs is not None:
        params["subsampling"] = int(subs)

    icc = info.get("icc_profile")
    if icc:
        params["icc_profile"] = icc

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


def save_plain_jpeg(im: Image.Image, dest: Path, quality: int, subs: Optional[int], progressive: bool) -> None:
    params = {
        "format": "JPEG",
        "quality": int(quality),
        "optimize": True,
        "progressive": bool(progressive),
    }
    if subs is not None:
        params["subsampling"] = int(subs)
    im.save(dest, **params)


def try_make_xls_with_excel(csv_path: Path, xls_path: Path) -> bool:
    try:
        import win32com.client
    except Exception:
        return False

    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.UseSystemSeparators = False
        excel.DecimalSeparator = ","
        excel.ThousandsSeparator = " "
        wb = excel.Workbooks.Open(str(csv_path))
        ws = wb.Worksheets(1)
        # horodatage_photo = 7e colonne = G
        ws.Columns(7).NumberFormatLocal = "jj/mm/aaaa hh:mm:ss"
        xlWorkbookNormal = -4143
        wb.SaveAs(str(xls_path), xlWorkbookNormal)
        wb.Close(SaveChanges=False)
        excel.Quit()
        return True
    except Exception:
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        return False


def jpg_output_name(p: Path) -> str:
    return p.stem + ".jpg"


def load_media_as_image(p: Path) -> Tuple[Image.Image, str, str]:
    """
    Retourne (image PIL RGB, horodatage_photo, orientation_photo)
    - JPG/JPEG/HEIC/HEIF : lecture image
    - MP4/MOV : première frame
    """
    ext = p.suffix.lower()

    if ext in {".jpg", ".jpeg", ".heic", ".heif"}:
        with Image.open(p) as im:
            horod, orient = extract_datetime_orientation(im)
            return ensure_rgb(im.copy()), horod, orient

    if ext in {".mp4", ".mov"}:
        frame = iio.imread(str(p), index=0)
        im = Image.fromarray(frame)
        return ensure_rgb(im), "", ""

    raise ValueError(f"Extension non supportée: {p.suffix}")


def main() -> int:
    cwd = Path.cwd()
    if cwd.name.lower() != "jpg":
        print("ERREUR: lancez ce script depuis le dossier 'JPG\\'. Dossier courant: " + str(cwd))
        return 1

    parent = cwd.parent
    out_dir = parent / OUT_DIR_NAME
    out_dir.mkdir(parents=True, exist_ok=True)

    csv_path = parent / CSV_NAME
    xls_path = parent / XLS_NAME

    accepted_exts = {".jpg", ".jpeg", ".heic", ".heif", ".mp4", ".mov"}
    images = sorted([p for p in cwd.iterdir() if p.is_file() and p.suffix.lower() in accepted_exts])

    if not images:
        print("Aucun fichier compatible (.jpg, .jpeg, .heic, .heif, .mp4, .mov) trouvé dans ce dossier.")
        return 0

    ui_header = [
        "photo_rel_native",
        "nom_fichier_image",
        "id_affaire",
        "id_captation",
        "chemin_photo_native",
        "chemin_photo_reduite",
        "horodatage_photo",
        "horodatage_secondes",
        "synchro_audio",
        "t_audio",
        "decalage_individuel",
        "decalage_moyen",
        "orientation_photo",
        "retenue",
        "description_vlm_ui",
        "vlm_ui_status",
        "vlm_ui_ts",
        "ui_ts",
        "libelle_propose_ui",
        "libelle_ui_ts",
        "libelle_ui_status",
        "commentaire_propose_ui",
        "commentaire_ui_ts",
        "commentaire_ui_status",
        "annotation_validee",
        "dictee_asr_text",
        "dictee_asr_status",
        "dictee_asr_ts",
        "dictee_audio_path_pcfixe",
        "dictee_asr_csv_path_pcfixe",
        "dictee_asr_photo_csv_path_pcfixe",
        "dictee_audio_sha256",
        "dictee_audio_size",
    ]

    batch_header = [
        "photo_rel_native",
        "chemin_photo_native_pcfixe",
        "chemin_photo_reduite_pcfixe",
        "photo_disponible_pcfixe",
        "date_copie_pcfixe",
        "description_vlm_batch",
        "libelle_propose_batch",
        "commentaire_propose_batch",
        "batch_status",
        "batch_id",
        "batch_ts",
        "vlm_batch_ts",
        "vlm_status",
        "vlm_batch_id",
        "vlm_err",
        "vlm_prompt_ctx_len",
        "vlm_img_bytes",
        "vlm_mode",
        "vlm_call_id",
        "llm_err",
        "llm_err_lib",
        "lll_errcom",
        "sujets_ids",
        "sujets_scores",
        "sujets_method",
        "sujets_justif",
    ]

    Header3 = [
        "nom_fichier_image",
        "horodatage_photo",
        "orientation_photo",
        "transcription_libelle",
        "libelle",
        "transcription_commentaire",
        "commentaire",
        "chemin_photo_reduite",
        "retenue",
        "t_audio_sec",
        "audio_timecode_hms",
        "audio_datetime_abs",
        "audio_start_sec",
        "audio_end_sec",
        "annotation_validee",
    ]

    print(f"{len(images)} fichiers détectés. Ecriture dans: '{out_dir.name}'")
    print(f"Paramètres: quality={QUALITY}, subsampling={SUBSAMPLING}, progressive={PROGRESSIVE}")

    chemin_native = os.path.join(str(cwd), "")
    chemin_reduite = os.path.join(str(out_dir), "")

    if cwd.name.lower() != "jpg" or cwd.parent.name.lower() != "photos":
        raise RuntimeError("Lancer depuis .../<id_captation>/photos/JPG")

    id_captation = cwd.parent.parent.name
    batch_rows = []
    written_count = 0

    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f_ui:
        w_ui = csv.DictWriter(f_ui, fieldnames=ui_header, delimiter=";")
        w_ui.writeheader()

        for p in images:
            try:
                ext = p.suffix.lower()
                out_name = p.name if ext in {".jpg", ".jpeg"} else jpg_output_name(p)

                dest_native = cwd / out_name
                dest_reduced = out_dir / out_name

                if ext in {".jpg", ".jpeg"}:
                    with Image.open(p) as im:
                        horod, orient = extract_datetime_orientation(im)
                        im2 = ensure_rgb(im)
                        save_with_metadata(im2, dest_reduced, QUALITY, SUBSAMPLING, PROGRESSIVE)
                else:
                    if (cwd / jpg_output_name(p)).exists():
                        print(f"[SKIP] {p.name} ignoré : JPG déjà présent")
                        continue

                    im2, horod, orient = load_media_as_image(p)

                    if not dest_native.exists():
                        save_plain_jpeg(im2, dest_native, quality=95, subs=0, progressive=False)

                    save_plain_jpeg(im2, dest_reduced, quality=QUALITY, subs=SUBSAMPLING, progressive=PROGRESSIVE)

                photo_rel_native = "/".join([
                    "AE_Expert_captations",
                    id_captation,
                    "photos",
                    "JPG",
                    out_name
                ])

                row_ui = {k: "" for k in ui_header}
                row_ui.update({
                    "photo_rel_native": photo_rel_native,
                    "nom_fichier_image": out_name,
                    "chemin_photo_native": chemin_native,
                    "chemin_photo_reduite": chemin_reduite,
                    "horodatage_photo": horod,
                    "orientation_photo": orient,
                    "annotation_validee": "0",
                    "ui_ts": "",
                    "dictee_asr_status": "",
                    "dictee_asr_text": "",
                })
                w_ui.writerow(row_ui)
                written_count += 1

                row_b = {k: "" for k in batch_header}
                row_b["photo_rel_native"] = photo_rel_native
                batch_rows.append(row_b)

                print(f"[OK] {p.name} -> {out_name}")

            except Exception as e:
                print(f"ERREUR sur {p.name}: {e}")
                continue

    batch_csv_path = csv_path.with_name(csv_path.stem + "_batch.csv")
    with open(batch_csv_path, "w", encoding="utf-8-sig", newline="") as f_b:
        w_b = csv.DictWriter(f_b, fieldnames=batch_header, delimiter=";")
        w_b.writeheader()
        for r in batch_rows:
            w_b.writerow(r)

    print(f"[OK] photos_batch.csv créé : {batch_csv_path}")
    print(f"[OK] {len(batch_rows)} lignes écrites dans photos_batch.csv")

    ok_xls = try_make_xls_with_excel(csv_path, xls_path)
    if not ok_xls:
        try:
            xls_path.write_bytes(csv_path.read_bytes())
            print("[Fallback] XLS copié depuis le CSV.")
        except Exception as e:
            print(f"[Fallback] Impossible de créer l'XLS: {e}")

    print(f"[INFO] id_captation détecté : {id_captation}")
    print(f"[OK] photos.csv généré : {csv_path}")
    print(f"[OK] {written_count} lignes écrites dans photos.csv")

    print("Terminé.")
    print("- Dossier réduit :", out_dir)
    print("- CSV           :", csv_path)
    print("- XLS           :", xls_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())