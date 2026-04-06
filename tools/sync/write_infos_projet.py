# write_infos_projet.py
# Usage:
#   python write_infos_projet.py <AFFAIRE> <CAPTATION> <PHOTOS_CSV> <PHOTOS_BATCH> <TRANSCRIPT_CSV> <WAV_MONO16> <WAV_SOURCE> <CTX_GENERAL> <ROOT_PC> <DST_TRANS> <INFOS_LAPTOP>

import json
import sys
from pathlib import Path

def p_or_empty(s: str) -> Path:
    s = (s or "").strip()
    return Path(s) if s else Path("")

def str_if_exists(p: Path) -> str:
    try:
        return str(p) if p and str(p) and p.exists() else ""
    except Exception:
        return ""

def main() -> int:
    if len(sys.argv) != 12:
        print(
            "Usage: python write_infos_projet.py <AFFAIRE> <CAPTATION> <PHOTOS_CSV> <PHOTOS_BATCH> "
            "<TRANSCRIPT_CSV> <WAV_MONO16> <WAV_SOURCE> <CTX_GENERAL> <ROOT_PC> <DST_TRANS> <INFOS_LAPTOP>",
            file=sys.stderr,
        )
        return 2

    aff = sys.argv[1]
    capt = sys.argv[2]
    photos_csv = Path(sys.argv[3])
    photos_batch = Path(sys.argv[4])
    tr_csv = p_or_empty(sys.argv[5])
    wav_mono = p_or_empty(sys.argv[6])
    wav_src = p_or_empty(sys.argv[7])
    ctx = p_or_empty(sys.argv[8])
    root_pc = Path(sys.argv[9])
    dst_trans = Path(sys.argv[10])
    infos_laptop = Path(sys.argv[11])

    # Dérivations PC fixe (structure canonique)
    dst_base = root_pc / aff / "AE_Expert_captations" / capt
    dst_ph = dst_base / "photos"
    dst_au = dst_base / "audio"
    dst_tr = root_pc / aff / "AF_Expert_ASR" / "transcriptions" / capt

    # proper_names_file : premier match dans DST_TRANS
    pn = ""
    try:
        cand = sorted(dst_trans.glob("*proper_names*.txt"))
        pn = str(cand[0]) if cand else ""
    except Exception:
        pn = ""

    d = {
        "fichier_transcription": str_if_exists(tr_csv),
        "fichier_photos": str_if_exists(photos_csv),
        "fichier_photos_batch": str_if_exists(photos_batch),
        "fichier_audio": r"C:\AnnotationPhotosGPT\data\temp\audio_compatible.wav",
        "audio_compat_source": str_if_exists(wav_mono),
        "fichier_audio_source": str_if_exists(wav_src),
        "fichier_audio_compatible": r"C:\AnnotationPhotosGPT\data\temp\audio_compatible.wav",
        "fichier_contexte_general": str_if_exists(ctx),
        "pcfixe": {
            "fichier_transcription": str(dst_tr / (tr_csv.name if tr_csv.name else "")),
            "fichier_photos": str(dst_ph / photos_csv.name),
            "fichier_photos_batch": str(dst_ph / (photos_batch.name if photos_batch.name else "photos_batch.csv")),
            "fichier_audio": str(dst_au / "audio_compatible.wav"),
            "audio_compat_source": str(dst_au / (wav_mono.name if wav_mono.name else "")),
            "fichier_audio_source": str(dst_au / (wav_src.name if wav_src.name else "")),
            "fichier_audio_compatible": str(dst_au / "audio_compatible.wav"),
            "fichier_contexte_general": str(dst_tr / (ctx.name if ctx.name else "")),
            "config_llm": str(dst_tr / "config_llm.json"),
            "out_dir": str(dst_au / "out"),
            "boost_file": r"D:\GPT4All_Local\config\boost_vocab.txt",
            "max_speakers": 6,
            "proper_names_file": pn,
        },
        "profil_execution": "pcfixe",
        "id_affaire": aff,
        "id_captation": capt,
    }

    infos_laptop.parent.mkdir(parents=True, exist_ok=True)
    dst_trans.mkdir(parents=True, exist_ok=True)

    payload = json.dumps(d, ensure_ascii=False, indent=2)
    infos_laptop.write_text(payload, encoding="utf-8")
    (dst_trans / "infos_projet.json").write_text(payload, encoding="utf-8")

    print(str(infos_laptop))
    print(str(dst_trans / "infos_projet.json"))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())