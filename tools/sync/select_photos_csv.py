from pathlib import Path
import streamlit as st
import requests
import os
from datetime import datetime

def list_csv_candidates(photos_dir: Path):
    cands = sorted(
        photos_dir.glob("*.csv"),
        key=lambda p: (p.name.lower() != "photos.csv", -p.stat().st_mtime)
    )
    items = []
    for p in cands:
        st_mtime = datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        size_kb = int(p.stat().st_size / 1024)
        label = f"{p.name}  —  {size_kb} KB  —  {st_mtime}"
        items.append((label, p))
    return items

id_affaire = st.text_input("ID affaire", value="2025-J40")
id_captation = st.text_input("ID captation", value="accedit-2026-01-23")

photos_dir = Path(rf"C:\Affaires\{id_affaire}\AE_Expert_captations\{id_captation}\photos")
jpg_dir = photos_dir / "JPG"

if not photos_dir.exists():
    st.error(f"Dossier introuvable : {photos_dir}")
else:
    candidates = list_csv_candidates(photos_dir)
    if not candidates:
        st.warning("Aucun .csv trouvé dans le dossier photos.")
    else:
        labels = [x[0] for x in candidates]
        selected_label = st.selectbox("Choisir le CSV des photos à recalculer", labels)
        selected_path = dict(candidates)[selected_label]

        st.write("Sélection :", str(selected_path))

        if st.button("Lancer propagation captation (seed)"):
            payload = {
                "id_affaire": id_affaire,
                "id_captation": id_captation,
                "cwd_jpg": str(jpg_dir),
                "photos_csv_path": str(selected_path),
                "root_dst": r"\\192.168.1.20\volume1\Affaires",
                "mode": "NAS",
            }
            # PC fixe = serveur Flask ; adaptez l'URL/route
            r = requests.post("http://PCFIXE:5000/seed_captation", json=payload, timeout=60)
            st.code(r.text)