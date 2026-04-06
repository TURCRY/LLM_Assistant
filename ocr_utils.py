# ocr_utils.py — version corrigée
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Any, List, Tuple, Optional

import csv
import hashlib
import json
import math
import time
from statistics import median

import cv2
import fitz  # PyMuPDF
import numpy as np
import pytesseract
from docx import Document

# =========================
# Options d'OCR
# =========================

@dataclass
class OcrOptions:
    lang: str = "fra"                 # ex: "fra", "eng", "fra+eng"
    psm: Optional[int] = None         # ex: 3, 6, 11
    oem: Optional[int] = None         # 0..3
    dpi: Optional[int] = 300
    denoise: bool = True
    threshold: bool = True
    threshold_method: str = "adaptive"  # "adaptive" | "otsu"
    deskew: bool = True
    return_hocr: bool = False

# =========================
# Aides scoring / profiling
# =========================

def _word_stats(text: str, lang="fra", use_wordfreq=True) -> Tuple[float, float]:
    """
    Retourne (fr_word_ratio, bad_char_ratio).
    - fr_word_ratio: part de tokens plausiblement français (approx).
    - bad_char_ratio: proportion de caractères non imprimables/suspects.
    """
    if not text:
        return 0.0, 1.0
    import re
    tokens = re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ'-]{2,}", text)
    total = max(len(tokens), 1)
    fr_hits = 0
    if use_wordfreq:
        try:
            from wordfreq import zipf_frequency
            for w in tokens:
                if zipf_frequency(w.lower(), "fr") > 2.5:
                    fr_hits += 1
        except Exception:
            use_wordfreq = False
    if not use_wordfreq:
        # Heuristique simple si wordfreq indisponible
        for w in tokens:
            if any(ch in "éèêàùçâîôû" for ch in w.lower()):
                fr_hits += 1
    fr_ratio = fr_hits / total if total else 0.0

    bad_chars = sum(1 for c in text if ord(c) < 9 or (ord(c) < 32 and c not in "\n\t\r"))
    bad_ratio = bad_chars / max(len(text), 1)
    return fr_ratio, bad_ratio

def _tess_conf_stats(img: np.ndarray, lang: str, cfg: str) -> Tuple[List[float], str]:
    """
    Utilise image_to_data pour récupérer les confidences (exclut -1) et le texte agrégé.
    """
    data = pytesseract.image_to_data(img, lang=lang, config=cfg, output_type=pytesseract.Output.DICT)
    confs = [float(c) for c in data.get("conf", []) if c not in ("-1", -1)]
    words = data.get("text", [])
    text = "\n".join([w for w in words if w and w.strip()])
    return confs, text

def _apply_threshold(gray: np.ndarray, method: str) -> np.ndarray:
    if method == "otsu":
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        return th
    # "adaptive" (défaut)
    th = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 35, 11
    )
    return th

def _preprocess_variant(bgr: np.ndarray, variant: Dict[str, Any]) -> np.ndarray:
    """
    Applique denoise/deskew + threshold choisi pour un essai de variante.
    """
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    if variant.get("denoise", True):
        gray = cv2.fastNlMeansDenoising(gray, h=7)
    if variant.get("deskew", True):
        coords = np.column_stack(np.where(gray < 255))
        if coords.size:
            rect = cv2.minAreaRect(coords)
            angle = rect[-1]
            angle = -(90 + angle) if angle < -45 else -angle
            (h, w) = gray.shape[:2]
            M = cv2.getRotationMatrix2D((w // 2, h // 2), angle, 1.0)
            gray = cv2.warpAffine(gray, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    gray = _apply_threshold(gray, variant.get("threshold", "adaptive"))
    return gray

def _tess_cfg(psm=None, oem=None, dpi=None) -> str:
    cfg = []
    if psm is not None:
        cfg.append(f"--psm {psm}")
    if oem is not None:
        cfg.append(f"--oem {oem}")
    if dpi is not None:
        cfg.append(f"--dpi {dpi}")
    return " ".join(cfg)

def _doc_hash(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()[:16]

def _bucketize(bright, contrast, skew, pages) -> str:
    b = "dark" if bright < 110 else "bright" if bright > 170 else "mid"
    c = "lowc" if contrast < 30 else "midc" if contrast < 60 else "highc"
    s = "skew" if abs(skew) > 1.5 else "flat"
    p = "short" if pages < 5 else "long"
    return f"{b}-{c}-{s}-{p}"

def _doc_profile(images: List[np.ndarray]) -> dict:
    """
    Profil minimal: nb pages, luminance/contraste moyens, skew approx. (échantillon).
    Sert à "bucketiser" les docs pour réutiliser l'expérience.
    """
    n = len(images)
    sample = [0, n // 2] if n >= 2 else [0]
    br, ct, skew = [], [], []
    for i in sample:
        bgr = images[i]
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        br.append(float(gray.mean()))
        ct.append(float(gray.std()))
        edges = cv2.Canny(gray, 50, 150)
        lines = cv2.HoughLines(edges, 1, np.pi/180, 200)
        ang = []
        if lines is not None:
            for rho, theta in lines[:, 0]:
                deg = (theta * 180 / np.pi) % 180
                if 5 < deg < 85:
                    ang.append(deg - 45)  # très grossier
        skew.append(float(np.median(ang)) if ang else 0.0)
    profile = {
        "pages": n,
        "brightness": float(np.mean(br)) if br else 0.0,
        "contrast": float(np.mean(ct)) if ct else 0.0,
        "skew_deg": float(np.mean(skew)) if skew else 0.0,
    }
    profile["bucket"] = _bucketize(profile["brightness"], profile["contrast"], profile["skew_deg"], n)
    return profile

def _variant_key(v: dict) -> str:
    keys = ["dpi", "psm", "oem", "denoise", "threshold", "deskew"]
    return "|".join(f"{k}={v.get(k)!r}" for k in keys)

def _load_history(path: Path) -> list:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []

def _experience_rerank(variants: list, history_path: Path, grid_name: str, bucket: str) -> list:
    """
    Récupère les meilleures variantes passées pour le même 'bucket',
    pondérées par récence (décroissance exponentielle), et les place en tête.
    """
    hist = _load_history(history_path)
    winners = {}  # key -> score agrégé
    now = time.time()
    for h in hist:
        if h.get("grid") != grid_name:
            continue
        ch = h.get("chosen", {}) or {}
        b = (ch.get("variant", {}) or {}).get("bucket", h.get("bucket"))
        if b != bucket:
            continue
        v = ch.get("variant", {})
        key = _variant_key(v)
        base = float(ch.get("score", 0.0))
        age = now - int(h.get("timestamp", now))
        decay = math.exp(-age / (30 * 24 * 3600))  # demi-vie ~ 20 j ; à ajuster
        winners[key] = winners.get(key, 0.0) + base * decay

    prefix = []
    for k, _ in sorted(winners.items(), key=lambda x: x[1], reverse=True):
        found = next((v for v in variants if _variant_key(v) == k), None)
        if found:
            prefix.append(found)

    seen = set(_variant_key(v) for v in prefix)
    tail = [v for v in variants if _variant_key(v) not in seen]
    return prefix + tail

# =========================
# Chargement images
# =========================

def _tess_config(opts: OcrOptions) -> str:
    cfg = []
    if opts.psm is not None:
        cfg.append(f"--psm {opts.psm}")
    if opts.oem is not None:
        cfg.append(f"--oem {opts.oem}")
    if opts.dpi is not None:
        cfg.append(f"--dpi {opts.dpi}")
    return " ".join(cfg)

def _pdf_to_images(pdf_path: Path, dpi: int = 300) -> List[np.ndarray]:
    pages = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, 3)
            pages.append(img[:, :, ::-1])  # BGR
    return pages

def _load_images_any(path: Path, dpi: int = 300) -> List[np.ndarray]:
    if path.suffix.lower() in {".pdf"}:
        return _pdf_to_images(path, dpi=dpi)
    img = cv2.imread(str(path), cv2.IMREAD_COLOR)
    if img is None:
        raise FileNotFoundError(f"Impossible de lire le fichier image: {path}")
    return [img]

# =========================
# OCR complet (DOCX + CSV)
# =========================

def _opencv_preprocess(img_bgr: np.ndarray, opts: OcrOptions) -> np.ndarray:
    """Prétraitement du run final: denoise -> deskew -> threshold (suivant opts)."""
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    if opts.denoise:
        gray = cv2.fastNlMeansDenoising(gray, h=7)
    if opts.deskew:
        coords = np.column_stack(np.where(gray < 255))
        if coords.size:
            rect = cv2.minAreaRect(coords)
            angle = rect[-1]
            angle = -(90 + angle) if angle < -45 else -angle
            (h, w) = gray.shape[:2]
            M = cv2.getRotationMatrix2D((w // 2, h // 2), angle, 1.0)
            gray = cv2.warpAffine(gray, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    if opts.threshold:
        if opts.threshold_method == "otsu":
            _, gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        else:
            gray = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                         cv2.THRESH_BINARY, 35, 11)
    return gray

def run_ocr(input_path: str, out_dir: str, opts: OcrOptions) -> Dict[str, Any]:
    """
    1) charge le document (PDF/image) -> images
    2) prétraite (OpenCV) -> gray/binaire
    3) Tesseract -> texte/pages
    4) DOCX + CSV (+ HOCR si demandé)
    """
    src = Path(input_path)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    stem = src.stem

    if not src.exists():
        raise FileNotFoundError(f"Fichier introuvable: {src}")

    images = _load_images_any(src, dpi=opts.dpi or 300)
    cfg = _tess_config(opts)

    pages_text: List[str] = []
    pages_hocr: List[str] = []

    for idx, bgr in enumerate(images, 1):
        proc = _opencv_preprocess(bgr, opts)
        txt = pytesseract.image_to_string(proc, lang=opts.lang, config=cfg) or ""
        pages_text.append(txt)
        if opts.return_hocr:
            hocr = pytesseract.image_to_pdf_or_hocr(proc, extension="hocr", lang=opts.lang, config=cfg)
            pages_hocr.append(hocr.decode("utf-8", errors="ignore"))

    # DOCX
    docx_path = out / f"{stem}.docx"
    doc = Document()
    for i, t in enumerate(pages_text, 1):
        if i > 1:
            doc.add_page_break()
        doc.add_paragraph(t)
    doc.save(docx_path)

    # CSV (page;line;text)
    csv_path = out / f"{stem}.csv"
    with csv_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["page", "line", "text"])
        for p, t in enumerate(pages_text, 1):
            for ln, line in enumerate(t.splitlines(), 1):
                w.writerow([p, ln, line])

    # HOCR par page (optionnel)
    hocr_paths: List[str] = []
    if opts.return_hocr:
        for i, h in enumerate(pages_hocr, 1):
            hp = out / f"{src.stem}_p{i:03d}.hocr.html"
            hp.write_text(h, encoding="utf-8", errors="ignore")
            hocr_paths.append(str(hp))

    return {
        "ok": True,
        "pages": len(pages_text),
        "docx_path": str(docx_path),
        "csv_path": str(csv_path),
        "hocr_paths": hocr_paths,
    }

# =========================
# Auto-tuning par grille
# =========================

def run_ocr_auto(
    input_path: str,
    out_dir: str,
    grid_path: str,
    lang: str = "fra",
    return_hocr: bool = False
) -> Dict[str, Any]:
    """
    Échantillonne, score les variantes de la grille, choisit la meilleure,
    lance l'OCR complet et journalise un historique.
    """
    t0 = time.time()

    src = Path(input_path)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    if not src.exists():
        raise FileNotFoundError(f"Fichier introuvable: {src}")

    grid = json.loads(Path(grid_path).read_text(encoding="utf-8"))
    variants = grid.get("variants", []) or []
    sampling = grid.get("sampling", {"pages": 2, "strategy": "head+middle"})
    weights = (grid.get("scoring", {}) or {}).get("weights", {}) or {}
    use_wordfreq = bool((grid.get("scoring", {}) or {}).get("use_wordfreq", True))

    # Images source (dpi max parmi les variantes pour l'échantillonnage)
    base_dpi = max([v.get("dpi", 300) for v in variants] or [300])
    imgs = _load_images_any(src, dpi=base_dpi)
    n = len(imgs)
    if n == 0:
        raise RuntimeError("Aucune page/image à traiter.")

    if sampling.get("strategy") == "head+middle" and n >= 2:
        sample_pages = [0, n // 2]
    else:
        sample_pages = list(range(min(n, int(sampling.get("pages", 2)))))

    # Profil & réordonnancement expérience
    profile = _doc_profile(imgs)
    grid_name = Path(grid_path).name
    history_path = Path(grid_path).with_name("ocr_history.json")
    variants = _experience_rerank(variants, history_path=history_path, grid_name=grid_name, bucket=profile["bucket"])

    # Essais
    trials: List[Dict[str, Any]] = []
    for vidx, v in enumerate(variants):
        psm, oem, dpi = v.get("psm"), v.get("oem"), v.get("dpi", 300)
        cfg = _tess_cfg(psm=psm, oem=oem, dpi=dpi)
        t_start = time.time()
        confs_all: List[float] = []
        text_all: List[str] = []
        for pi in sample_pages:
            pre = _preprocess_variant(imgs[pi], v)
            confs, text = _tess_conf_stats(pre, lang=lang, cfg=cfg)
            confs_all += confs
            text_all.append(text)
        runtime = time.time() - t_start
        med_conf = float(median(confs_all)) if confs_all else 0.0
        fr_ratio, bad_ratio = _word_stats("\n".join(text_all), lang=lang, use_wordfreq=use_wordfreq)

        score = (
            weights.get("median_conf", 0.5) * med_conf +
            weights.get("fr_word_ratio", 0.3) * (fr_ratio * 100) +
            weights.get("bad_char_ratio", -0.1) * (bad_ratio * 100) +
            weights.get("runtime_s", -0.1) * runtime
        )
        trials.append({
            "variant_index": vidx,
            "variant": v,
            "median_conf": med_conf,
            "fr_word_ratio": fr_ratio,
            "bad_char_ratio": bad_ratio,
            "runtime_s": runtime,
            "score": score
        })

    if not trials:
        raise RuntimeError("Aucune variante n'a pu être évaluée.")

    best = max(trials, key=lambda t: t["score"])

    # OCR complet avec la meilleure variante
    best_opts = OcrOptions(
        lang=lang,
        psm=best["variant"].get("psm"),
        oem=best["variant"].get("oem"),
        dpi=best["variant"].get("dpi", 300),
        denoise=best["variant"].get("denoise", True),
        threshold=True,
        threshold_method=best["variant"].get("threshold", "adaptive"),
        deskew=best["variant"].get("deskew", True),
        return_hocr=return_hocr,
    )

    res = run_ocr(input_path=str(src), out_dir=str(out), opts=best_opts)

    # Journalisation
    try:
        hist = json.loads(history_path.read_text(encoding="utf-8"))
        if not isinstance(hist, list):
            hist = []
    except Exception:
        hist = []
    hist.append({
        "timestamp": int(time.time()),
        "file": str(src),
        "doc_hash": _doc_hash(src),
        "grid": grid_name,
        "bucket": profile["bucket"],
        "profile": profile,
        "trials": trials,
        "chosen": {**best, "variant": {**best["variant"], "bucket": profile["bucket"]}},
        "pages": res.get("pages", 0),
        "duration_total_s": time.time() - t0,
    })
    history_path.write_text(json.dumps(hist, indent=2, ensure_ascii=False), encoding="utf-8")

    return {
        "ok": True,
        "selected_variant": best,
        **res,
    }
