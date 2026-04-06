#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import os
from pathlib import Path
from datetime import datetime
import pandas as pd

REQUIRED_COLS = [
    "nom_fichier_image",
    "horodatage_photo",
    "orientation_photo",
    "chemin_photo_native",
    "chemin_photo_reduite",
    "horodatage_secondes",
    "t_audio",
    "decalage_individuel",
    "synchro_audio",
    "decalage_moyen",
]

OPTIONAL_SYNC_COLS = ["sec_of_day"]

BATCH_COLS = [
    "chemin_photo_native_pcfixe",
    "chemin_photo_reduite_pcfixe",
    "id_affaire",
    "id_captation",
    "photo_rel_native",
    "photo_rel_reduite",
    "photo_disponible_pcfixe",
    "date_copie_pcfixe",
]

def ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    for c in REQUIRED_COLS:
        if c not in df.columns:
            df[c] = ""
    for c in OPTIONAL_SYNC_COLS:
        if c not in df.columns:
            df[c] = ""
    for c in BATCH_COLS:
        if c not in df.columns:
            df[c] = ""
    return df

def calc_relpaths(id_captation: str, nom: str):
    rel_native  = f"AE_Expert_captations/{id_captation}/photos/JPG/{nom}"
    rel_reduite = f"AE_Expert_captations/{id_captation}/photos/JPG reduit/{nom}"
    return rel_native, rel_reduite

def _dir_with_trailing_backslash(p: Path) -> str:
    s = os.path.normpath(str(p))
    return s + "\\"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--photos_csv", required=True)
    ap.add_argument("--id_affaire", required=True)
    ap.add_argument("--id_captation", required=True)
    ap.add_argument("--root_pcfixe", required=True, help=r"Racine des affaires. Ex: \\192.168.0.155\Affaires ou C:\Affaires")
    args = ap.parse_args()

    csv_path = Path(args.photos_csv)
    if not csv_path.exists():
        raise FileNotFoundError(csv_path)

    df = pd.read_csv(csv_path, sep=";", encoding="utf-8-sig")
    df = ensure_columns(df)

    # Renseigner id_affaire / id_captation (si vide)
    df["id_affaire"] = df["id_affaire"].fillna("").astype(str)
    df.loc[df["id_affaire"].str.strip().eq(""), "id_affaire"] = args.id_affaire

    df["id_captation"] = df["id_captation"].fillna("").astype(str)
    df.loc[df["id_captation"].str.strip().eq(""), "id_captation"] = args.id_captation

    # racine affaire (root_pcfixe = racine des affaires)
    root_affaire = Path(args.root_pcfixe) / args.id_affaire

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    dispo = []
    for idx, row in df.iterrows():
        # id_captation (ligne) : fallback sur args
        id_capt = str(row.get("id_captation", "")).strip() or args.id_captation

        # nom_fichier_image : fallback depuis photo_rel_native
        nom = str(row.get("nom_fichier_image", "")).strip()
        if nom.lower() == "nan":
            nom = ""
        if not nom:
            prn = str(row.get("photo_rel_native", "")).strip().replace("\\", "/")
            if prn:
                nom = prn.split("/")[-1].strip()

        if not nom or not id_capt:
            dispo.append(0)
            continue

        # photo_rel_* : conserver si déjà présent, sinon calculer
        existing_prn = str(row.get("photo_rel_native", "")).strip()
        if existing_prn.lower() == "nan":
            existing_prn = ""

        if existing_prn:
            rel_nat = existing_prn.replace("\\", "/")
            rel_red = rel_nat.replace("/photos/JPG/", "/photos/JPG reduit/")
        else:
            rel_nat, rel_red = calc_relpaths(id_capt, nom)

        df.at[idx, "photo_rel_native"] = rel_nat
        df.at[idx, "photo_rel_reduite"] = rel_red
        df.at[idx, "nom_fichier_image"] = nom
        df.at[idx, "id_captation"] = id_capt

        abs_nat = root_affaire / rel_nat
        abs_red = root_affaire / rel_red

        df.at[idx, "chemin_photo_native_pcfixe"] = _dir_with_trailing_backslash(abs_nat.parent)
        df.at[idx, "chemin_photo_reduite_pcfixe"] = _dir_with_trailing_backslash(abs_red.parent)

        ok = 1 if abs_red.exists() or abs_nat.exists() else 0
        dispo.append(ok)
        if ok:
            df.at[idx, "date_copie_pcfixe"] = now_str

    df["photo_disponible_pcfixe"] = dispo

    df.to_csv(csv_path, sep=";", encoding="utf-8-sig", index=False)

    print("OK - photos.csv mis à jour :", csv_path)

if __name__ == "__main__":
    main()
