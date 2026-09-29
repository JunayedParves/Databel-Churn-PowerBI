"""Run `powerbi-report-author validate` and print a compact summary, then check
that every field/measure referenced by a visual exists in the TMDL model."""
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REPORT = ROOT / "Databel Churn.Report"
MODEL = ROOT / "Databel Churn.SemanticModel" / "definition"


def cli_validate():
    exe = shutil.which("powerbi-report-author") or shutil.which("powerbi-report-author.cmd")
    out = subprocess.run([exe, "validate", str(REPORT)], capture_output=True, text=True,
                         encoding="utf-8-sig").stdout
    data = json.loads(out)["data"]
    print(f"CLI validate: {data['result']}  errors={data['errorCount']} warnings={data['warningCount']}")
    for code, info in data.get("diagnostics", {}).items():
        items = info["items"]
        print(f"  [{info['severity']}] {code} x{len(items)}")
        for it in items[:3]:
            msg = it["message"].replace(str(REPORT) + "\\", "")
            print("     -", msg[:300])
    return data["errorCount"] == 0


def model_objects():
    objs = set()
    for f in (MODEL / "tables").glob("*.tmdl"):
        text = f.read_text(encoding="utf-8")
        table = re.search(r"^table (.+)$", text, re.M).group(1).strip().strip("'")
        for kind, name in re.findall(r"^\t(column|measure) ('(?:[^']|'')+'|[^\s=]+)", text, re.M):
            objs.add((table, name.strip("'").replace("''", "'"), kind))
    return objs


def field_refs(node, acc):
    if isinstance(node, dict):
        for k in ("Column", "Measure"):
            v = node.get(k)
            if isinstance(v, dict) and "Property" in v:
                src = v.get("Expression", {}).get("SourceRef", {})
                acc.append((k.lower(), src.get("Entity") or src.get("Source"), v["Property"]))
        for v in node.values():
            field_refs(v, acc)
    elif isinstance(node, list):
        for v in node:
            field_refs(v, acc)


def check_fields():
    objs = model_objects()
    names = {(t, n): k for t, n, k in objs}
    aliases = {"d": "Databel", "m": "_Measures"}
    bad, total = [], 0
    for f in REPORT.glob("definition/**/*.json"):
        data = json.loads(f.read_text(encoding="utf-8"))  # also proves every file parses
        refs = []
        field_refs(data, refs)
        for kind, ent, prop in refs:
            total += 1
            ent = aliases.get(ent, ent)
            if names.get((ent, prop)) != kind:
                bad.append(f"{f.relative_to(REPORT)}: {kind} {ent}[{prop}]")
    print(f"Field references checked: {total}; missing: {len(bad)}")
    for x in bad:
        print("   -", x)
    return not bad


def check_json():
    n = 0
    for f in ROOT.rglob("*.json"):
        json.loads(f.read_text(encoding="utf-8"))
        n += 1
    for f in ROOT.rglob("*.pb*"):
        if f.is_file():
            json.loads(f.read_text(encoding="utf-8"))
            n += 1
    for f in ROOT.rglob(".platform"):
        json.loads(f.read_text(encoding="utf-8"))
        n += 1
    print(f"JSON files parsed: {n}")


if __name__ == "__main__":
    check_json()
    ok = check_fields()
    ok = cli_validate() and ok
    sys.exit(0 if ok else 1)
