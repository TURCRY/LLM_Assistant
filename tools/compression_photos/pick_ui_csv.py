import sys
from pathlib import Path

def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python pick_ui_csv.py <photos_dir>", file=sys.stderr)
        return 2

    d = Path(sys.argv[1])
    if not d.exists():
        print("", end="")
        return 0

    candidates = []

    for p in d.glob("*.csv"):
        name = p.name
        if "_GTP_" in name:
            continue
        if "_batch" in p.stem.lower():
            continue
        if name.lower() == "photos_batch.csv":
            continue
        candidates.append(p)

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)

    print(str(candidates[0]) if candidates else "")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())