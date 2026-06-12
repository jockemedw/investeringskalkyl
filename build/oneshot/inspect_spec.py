"""Snabbinspektion av spec_iter9.json: `python inspect_spec.py <flik> [ref|rad-intervall]`."""
import io
import json
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
spec = json.load(open(Path(__file__).parent / "spec_iter9.json", encoding="utf-8"))
sheets = {s["name"]: s for s in spec["sheets"]}

name = sys.argv[1]
sel = sys.argv[2] if len(sys.argv) > 2 else None
cols = sys.argv[3].split(",") if len(sys.argv) > 3 else None
s = sheets[name]

import re
def rowno(ref):
    return int(re.sub(r"[A-Z]+", "", ref))

for c in s["cells"]:
    if not (set(c) & {"f", "v", "link"}):
        continue
    if cols and re.sub(r"\d+", "", c["ref"]) not in cols:
        continue
    if sel:
        if "-" in sel and sel.replace("-", "").isdigit():
            lo, hi = map(int, sel.split("-"))
            if not (lo <= rowno(c["ref"]) <= hi):
                continue
        elif c["ref"] != sel:
            continue
    out = {k: v for k, v in c.items() if k != "style"}
    st = c.get("style", {})
    brief = []
    if "font" in st:
        f = st["font"]
        brief.append(f"font={f['name']},{f['size']},{'b' if f['bold'] else ''},{f['color']}")
    if "fill" in st:
        brief.append(f"fill={st['fill']}")
    if "numfmt" in st:
        brief.append(f"fmt={st['numfmt']}")
    print(out, "|", "; ".join(brief))
