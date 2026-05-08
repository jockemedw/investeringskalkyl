"""Iter 9 — bygger ovanpå iter 8 med §10-uppgifter.

Varje funktion = en §10-uppgift. main() kör i sekvens. Idempotent: kan
köras om utan att stega framåt mer än en gång (text-replaces matchar inte
redan-patchade celler).

Genomförda uppgifter (commit-historik förklarar):
- round E: 'år N' / 'år N+1' → 'år 20' / 'år 21' (explicit horisont)
"""
from __future__ import annotations
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook
from tools.patches import load_iter, save_iter, ITER8
from tools.recalc import recalc
from tests.regression import check_baseline


def round_e_year_n(wb: Workbook) -> int:
    """§10 punkt 4: 'år N' → 'år 20', 'år N+1' → 'år 21'.

    Kalkylperioden är 20 år (Indata!C16) och refereras med variabeln N i text.
    Round E gör horisonten explicit för läsaren.

    OBS: 'år N+1' måste replaceas FÖRE 'år N' (annars matchar 'år N' delvis).
    Lämnar formelvariabeln 'N × avskrivnings-...' i Dokumentation!B31 orörd —
    den är en algebraisk variabel, inte en horisontetikett.
    """
    replacements = [
        ("år N+1", "år 21"),
        ("år N",   "år 20"),
        ("År N",   "år 20"),
    ]
    n_changes = 0
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for c in row:
                if not isinstance(c.value, str):
                    continue
                original = c.value
                new_val = original
                for old, new in replacements:
                    new_val = new_val.replace(old, new)
                if new_val != original:
                    c.value = new_val
                    n_changes += 1
    return n_changes


def main() -> int:
    out = ROOT / "build" / "iter9.xlsx"
    out.parent.mkdir(exist_ok=True)

    print(f"1. Laddar {ITER8.name}")
    wb = load_iter(ITER8)

    print("2. Round E: 'år N' → 'år 20' / 'år N+1' → 'år 21'")
    n = round_e_year_n(wb)
    print(f"   {n} celler patchade")

    print(f"\n3. Sparar → {out.name}")
    save_iter(wb, out)

    print("\n4. Recalc (Excel COM / LibreOffice)")
    recalc(out)

    print("\n5. Regressionstest:")
    res = check_baseline(out)
    if res["rent"] is not None:
        print(f"   Hyra={res['rent']:,.2f}  IRR={res['irr']:.4%}  margin={res['margin']*100:+.2f} pp")
    if res["ok"]:
        print("   ✓ Regression OK")
        return 0
    print("   ✗ FAIL:", res["fails"])
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
