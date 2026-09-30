"""Kör apply.ps1 (Excel COM) med en ops-lista och returnerar probes/scan-resultat."""
from __future__ import annotations
import json
import subprocess
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
APPLY = HERE / "apply.ps1"


def run(source: Path, ops: list[dict], out: Path | None = None) -> dict:
    # Range.Replace matchar mot FormulaLocal (svenska funktionsnamn, ';' som separator) och
    # avvisar tyst ersättningar som blir ogiltiga formler. Sök/ersätt får därför varken
    # innehålla ',' eller funktionsnamn — sätt hela formeln med .Formula i stället.
    for op in ops:
        if op.get("op") == "replace" and ("," in op["find"] or "," in op["repl"] or "(" in op["repl"]):
            raise ValueError(f"replace-op med separator/funktion (fungerar inte mot FormulaLocal): {op}")
    with tempfile.TemporaryDirectory() as td:
        ops_p = Path(td) / "ops.json"
        probes_p = Path(td) / "probes.json"
        ops_p.write_text(json.dumps(ops, ensure_ascii=False), encoding="utf-8")
        cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(APPLY),
               "-Source", str(source), "-Ops", str(ops_p), "-Probes", str(probes_p)]
        if out is not None:
            cmd += ["-Out", str(out)]
        res = subprocess.run(cmd, capture_output=True, encoding="utf-8", errors="replace", timeout=900)
        if res.returncode != 0 or not probes_p.exists():
            raise RuntimeError(f"apply.ps1 misslyckades:\n{res.stdout}\n{res.stderr}")
        return json.loads(probes_p.read_text(encoding="utf-8-sig"))


def probe(sheet: str, ref: str) -> dict:
    return {"op": "probe", "sheet": sheet, "ref": ref}


def scan(sheet: str, ref: str) -> dict:
    return {"op": "scan", "sheet": sheet, "ref": ref}
