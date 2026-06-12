"""Lista alla TEXT()-formler (locale-känsliga) i spec_iter9.json."""
import io, json, sys
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
spec = json.load(open(Path(__file__).parent / "spec_iter9.json", encoding="utf-8"))
for s in spec["sheets"]:
    for c in s["cells"]:
        f = c.get("f", "")
        if "TEXT(" in f:
            print(f"{s['name']}!{c['ref']}: {f}")
            print(f"   cached: {c.get('cached')}")
