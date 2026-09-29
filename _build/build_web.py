"""Pack data/Databel - Data.csv into docs/data.js for the interactive web version.

Categorical fields are stored as an index into a level list, so the whole dataset
is a few hundred KB and loads with a plain <script> tag (works on file:// too).

    python _build/build_web.py
"""
import csv
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "data" / "Databel - Data.csv"
DOCS = ROOT / "docs"

# output key -> (source column, kind)
FIELDS = {
    "churn": ("Churn Label", "bool"),
    "cat": ("Churn Category", "cat"),
    "reason": ("Churn Reason", "cat"),
    "contract": ("Contract Type", "cat"),
    "tenure": ("Account Length (in months)", "int"),
    "payment": ("Payment Method", "cat"),
    "calls": ("Customer Service Calls", "int"),
    "intlPlan": ("Intl Plan", "cat"),
    "intlActive": ("Intl Active", "cat"),
    "unlimited": ("Unlimited Data Plan", "cat"),
    "device": ("Device Protection & Online Backup", "cat"),
    "age": ("Age", "int"),
    "senior": ("Senior", "cat"),
    "gender": ("Gender", "cat"),
    "state": ("State", "cat"),
    "group": ("Group", "cat"),
    "charge": ("Monthly Charge", "int"),
}


def main():
    with SRC.open(encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    out = {"n": len(rows), "levels": {}, "cols": {}}
    for key, (col, kind) in FIELDS.items():
        values = [r[col].strip() for r in rows]
        if col == "Intl Plan":
            values = [v.title() for v in values]  # same as the Power Query step
        if kind == "bool":
            out["cols"][key] = [1 if v == "Yes" else 0 for v in values]
        elif kind == "int":
            out["cols"][key] = [int(v) for v in values]
        else:
            levels = sorted({v for v in values if v})
            index = {v: i for i, v in enumerate(levels)}
            out["levels"][key] = levels
            out["cols"][key] = [index.get(v, -1) for v in values]  # -1 = blank
    DOCS.mkdir(exist_ok=True)
    payload = json.dumps(out, separators=(",", ":"))
    (DOCS / "data.js").write_text(f"window.DATABEL={payload};\n", encoding="utf-8", newline="\n")
    print(f"docs/data.js: {len(rows)} rows, {len(payload) // 1024} KB")

    shots = ROOT / "_build" / "screenshots"
    if shots.exists():
        (DOCS / "img").mkdir(exist_ok=True)
        for png in shots.glob("*.png"):
            shutil.copyfile(png, DOCS / "img" / png.name)


if __name__ == "__main__":
    main()
