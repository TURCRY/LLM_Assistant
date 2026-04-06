# mk_photos_batch_min.py
# Usage: python mk_photos_batch_min.py <ui_csv> <out_csv>

import csv
import sys
from pathlib import Path

HDR = [
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
    "sujets_ids",
    "sujets_scores",
    "sujets_method",
    "sujets_justif",
]

def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: python mk_photos_batch_min.py <ui_csv> <out_csv>", file=sys.stderr)
        return 2

    ui = Path(sys.argv[1])
    out = Path(sys.argv[2])

    if not ui.exists():
        print(f"ERREUR: UI CSV introuvable: {ui}", file=sys.stderr)
        return 3

    # Lecture UI (UTF-8-SIG, séparateur ;)
    with ui.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))

    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=HDR, delimiter=";")
        w.writeheader()
        for r in rows:
            base = {h: "" for h in HDR}
            base["photo_rel_native"] = (r.get("photo_rel_native") or "").strip()
            w.writerow(base)

    print(str(out))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())