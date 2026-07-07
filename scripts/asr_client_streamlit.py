# asr_client_streamlit.py
# Client Streamlit côté laptop pour tester la transcription Voxtral via le serveur Flask.
# - Upload audio -> copie sur partage réseau du PC fixe -> appel /asr_voxtral
# - OU saisie d'un chemin local déjà présent sur le serveur
# - Liste les modèles ASR via /asr_models
#
# Dépendances côté laptop :
#   pip install streamlit requests python-dotenv
#
# Lancer :
#   streamlit run asr_client_streamlit.py

import os
import io
import time
import shutil
import requests
import streamlit as st
from pathlib import Path
import sys
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from server_locator import resolve_flask_base_url

load_dotenv()  # charge .env local si présent

st.set_page_config(page_title="ASR Voxtral (client laptop)", page_icon="🎤", layout="wide")

# ============== Helpers ==============

def get_env_default(key: str, default: str = "") -> str:
    return os.environ.get(key, default)

def api_get(base_url: str, path: str, api_key: str, timeout: int = 30):
    url = base_url.rstrip("/") + path
    headers = {"x-api-key": api_key}
    r = requests.get(url, headers=headers, timeout=timeout)
    r.raise_for_status()
    return r.json()

def api_post(base_url: str, path: str, api_key: str, payload: dict, timeout: int = 600):
    url = base_url.rstrip("/") + path
    headers = {"x-api-key": api_key, "Content-Type": "application/json"}
    r = requests.post(url, headers=headers, json=payload, timeout=timeout)
    r.raise_for_status()
    return r.json()

def copy_to_network_share(file_bytes: bytes, filename: str, unc_dir: str) -> str:
    """
    Copie le fichier uploadé vers un partage réseau UNC (ex: \\PC-Fixe\Drop_transcrip\).
    Renvoie le chemin UNC complet écrit (utile debug), lève en cas d'échec.
    """
    # Normaliser le chemin UNC (finir par antislash)
    if not unc_dir.endswith("\\"):
        unc_dir = unc_dir + "\\"
    dest_unc = unc_dir + filename
    # Crée dossier si possible (sur certains partages, création auto interdite)
    parent = Path(unc_dir)
    try:
        parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        # possible que le partage ne permette pas mkdir → on ignore
        pass
    with open(dest_unc, "wb") as f:
        f.write(file_bytes)
    return dest_unc

# ============== UI ==============

st.title("🎤 ASR Voxtral — Client (Laptop)")

with st.expander("⚙️ Paramètres serveur", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        server_base_url = st.text_input(
            "URL du serveur Flask",
            value=get_env_default("ASR_SERVER_URL", resolve_flask_base_url()),
            help="Ex: http://PC-FIXE:5050"
        )
        api_key = st.text_input(
            "x-api-key",
            value=get_env_default("ASR_API_KEY", ""),
            type="password",
            help="Clé API attendue par le serveur"
        )
        timeout_s = st.number_input("⏱️ Timeout (sec)", min_value=10, max_value=3600, value=600, step=10)
    with col2:
        st.write("")

    if st.button("🔎 Tester /ping"):
        try:
            res = api_get(server_base_url, "/ping", api_key, timeout=15)
            st.success(f"Serveur OK : {res}")
        except Exception as e:
            st.error(f"Ping KO : {e}")

# Récupérer la liste de modèles ASR
models = []
model_details = {}
try:
    res_models = api_get(server_base_url, "/asr_models", api_key, timeout=15)
    models = res_models.get("models", [])
    model_details = res_models.get("details", {})
except Exception as e:
    st.warning(f"Impossible de lister /asr_models : {e} — Saisissez le nom manuellement si besoin.")

default_model_key = "Voxtral_Mini_3B_Transformers" if "Voxtral_Mini_3B_Transformers" in models else (models[0] if models else "")

with st.expander("🎚️ Options ASR", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        selected_model = st.selectbox("Modèle ASR", options=["(saisir manuellement)"] + models, index=(0 if not default_model_key else (models.index(default_model_key)+1 if default_model_key in models else 0)))
        if selected_model == "(saisir manuellement)":
            model_key = st.text_input("Clé modèle (models_index.json)", value=default_model_key or "Voxtral_Mini_3B_Transformers")
        else:
            model_key = selected_model

        lang = st.text_input("Langue (optionnelle)", value="fr", help="Ex: fr, en, es …")
        timestamps = st.checkbox("Inclure les timestamps", value=False)
    with col2:
        chunk = st.number_input("Chunk (s)", min_value=5, max_value=120, value=30, step=5)
        stride = st.number_input("Stride (s)", min_value=0, max_value=30, value=5, step=1)
    with col3:
        force_cpu = st.checkbox("Forcer CPU", value=False)
        no4bit = st.checkbox("Désactiver 4-bit", value=False)

st.markdown("---")

mode = st.radio("Mode d’accès au fichier audio sur le serveur", ["📁 Copier via partage réseau (recommandé)", "📝 J'ai déjà le chemin local sur le serveur"], index=0)

if mode == "📁 Copier via partage réseau (recommandé)":
    st.info("Le fichier sera copié sur un **partage réseau UNC** du PC fixe, puis la route `/asr_voxtral` sera appelée avec le **chemin local** vu par le serveur.")

    with st.form("form_upload"):
        uploaded = st.file_uploader("Fichier audio (wav/mp3/flac...)", type=["wav", "mp3", "flac", "m4a", "ogg"])
        unc_dir = st.text_input("Dossier **UNC** du partage (serveur) pour déposer le fichier", value="\\\\PC-Fixe\\Drop_transcrip\\", help="Ex: \\\\PC-Fixe\\Drop_transcrip\\")
        server_local_dir = st.text_input("Chemin **local (serveur)** correspondant", value="D:\\Drop_transcrip\\", help="Ex: D:\\Drop_transcrip\\ — le serveur doit voir le fichier à cette adresse.")
        submitted = st.form_submit_button("➡️ Transcrire")

        if submitted:
            if not uploaded:
                st.error("Veuillez sélectionner un fichier audio.")
            elif not api_key or not server_base_url:
                st.error("Renseignez URL du serveur et x-api-key.")
            else:
                try:
                    filename = uploaded.name
                    data_bytes = uploaded.read()

                    # 1) copie UNC
                    st.write("📤 Copie sur le partage réseau…")
                    t0 = time.time()
                    dest_unc = copy_to_network_share(data_bytes, filename, unc_dir)
                    st.success(f"Copie OK → {dest_unc} ({time.time()-t0:.1f}s)")

                    # 2) chemin local serveur
                    if not server_local_dir.endswith("\\"):
                        server_local_dir = server_local_dir + "\\"
                    audio_path_server = server_local_dir + filename

                    # 3) appel /asr_voxtral
                    st.write("🛰️ Appel /asr_voxtral…")
                    payload = {
                        "audio_path": audio_path_server,
                        "model_key": model_key,
                        "timestamps": timestamps,
                        "lang": (lang or None),
                        "chunk": int(chunk),
                        "stride": int(stride),
                        "cpu": bool(force_cpu),
                        "no4bit": bool(no4bit)
                    }
                    res = api_post(server_base_url, "/asr_voxtral", api_key, payload, timeout=timeout_s)
                    text = res.get("text", "")
                    st.success("✅ Transcription terminée.")
                    st.text_area("Texte", value=text, height=250)

                    if "chunks" in res:
                        with st.expander("📎 Chunks & timestamps"):
                            st.write(res["chunks"])

                    # Télécharger la transcription
                    st.download_button(
                        "💾 Télécharger .txt",
                        data=text,
                        file_name=Path(filename).with_suffix(".txt").name,
                        mime="text/plain",
                    )

                except Exception as e:
                    st.error(f"Erreur : {e}")

else:
    st.info("Si le fichier est déjà présent sur le serveur, indiquez le **chemin local** directement (ex: D:\\audios\\reunion.wav).")
    with st.form("form_manual"):
        audio_path_server = st.text_input("Chemin **local (serveur)** du fichier audio", value="D:\\audios\\exemple.wav")
        submitted2 = st.form_submit_button("➡️ Transcrire (sans copie)")

        if submitted2:
            if not api_key or not server_base_url:
                st.error("Renseignez URL du serveur et x-api-key.")
            elif not audio_path_server.strip():
                st.error("Chemin audio manquant.")
            else:
                try:
                    payload = {
                        "audio_path": audio_path_server.strip(),
                        "model_key": model_key,
                        "timestamps": timestamps,
                        "lang": (lang or None),
                        "chunk": int(chunk),
                        "stride": int(stride),
                        "cpu": bool(force_cpu),
                        "no4bit": bool(no4bit)
                    }
                    res = api_post(server_base_url, "/asr_voxtral", api_key, payload, timeout=timeout_s)
                    text = res.get("text", "")
                    st.success("✅ Transcription terminée.")
                    st.text_area("Texte", value=text, height=250)

                    if "chunks" in res:
                        with st.expander("📎 Chunks & timestamps"):
                            st.write(res["chunks"])

                    st.download_button(
                        "💾 Télécharger .txt",
                        data=text,
                        file_name=Path(audio_path_server).with_suffix(".txt").name,
                        mime="text/plain",
                    )

                except Exception as e:
                    st.error(f"Erreur : {e}")
