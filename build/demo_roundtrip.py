"""Demo: round-trip iter8 → iter8_roundtrip.xlsx via patches.py.

Verifierar att pipelinen kan ladda, spara och bevara cached values utan
att röra något — sanity check innan verkliga §10-patches körs.
"""
from __future__ import annotations
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from tools.patches import load_iter, save_iter, diff_workbooks, ITER8
from tools.recalc import recalc
from tests.regression import check_baseline


def main() -> int:
    out = ROOT / "build" / "iter8_roundtrip.xlsx"
    out.parent.mkdir(exist_ok=True)

    print(f"1. Laddar {ITER8.name}")
    wb = load_iter(ITER8)
    print(f"   Flikar: {wb.sheetnames}")

    print(f"\n2. Sparar utan ändringar → {out.name}")
    save_iter(wb, out)
    size_in  = ITER8.stat().st_size
    size_out = out.stat().st_size
    print(f"   Storlek: in={size_in:,} out={size_out:,} bytes")

    print("\n3. Cell-diff mellan original och round-trip:")
    diffs = diff_workbooks(ITER8, out)
    if not diffs:
        print("   ✓ Inga celldiffar")
    else:
        for sheet, sdiffs in diffs.items():
            print(f"   {sheet}: {len(sdiffs)} skillnader")
            for coord, va, vb in sdiffs[:3]:
                print(f"     {coord}: {va!r} -> {vb!r}")

    print("\n4. Recalc via Excel/LibreOffice (uppdaterar cached values)")
    recalc(out)
    print("   ✓ Recalc klar")

    print("\n5. Regressionstest mot recalcad round-trip-fil:")
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
