"""Regressionstest mot iter 8-baslinjen.

Två lägen:
  - cached_only:  läser cached values från xlsx (ingen omräkning) — snabbt,
                  funkar bara på filer som öppnats/sparats av Excel/LibreOffice
  - recalc:       kör LibreOffice headless först → läser cached — säkrast efter
                  openpyxl-skrivning (openpyxl rör inte cached values)

Baslinje (iter 8 mot iter 7):
  Resultat!D14            = 11 631 221,72  (bindande kravhyra mål-utfall)
  Lönsamhetskontroll!C45  = 0,07645        (faktisk IRR EK)
  IRR-marginal            = +1,35 pp       (IRR EK − IRR-krav 6,3 %)
"""
from __future__ import annotations
import sys
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
ITER8 = ROOT / "Investeringskalkyl_iter8.xlsx"

EXPECTED = {
    "binding_rent_kr_per_year": 11_631_222,    # Resultat!D14, mål-utfall
    "actual_equity_irr":        0.0765,        # Lönsamhetskontroll!C45
    "irr_margin_pp":            0.0135,        # IRR EK − Indata!C65
}
TOLERANCE_RENT_KR = 1.0
TOLERANCE_PCT     = 0.0001  # 1 bp


def check_baseline(xlsx_path: Path | str = ITER8) -> dict:
    """Läs cached values och jämför mot baslinjen. Returnerar resultatdict."""
    wb = load_workbook(Path(xlsx_path), data_only=True)
    rent = wb["Resultat"]["D14"].value
    irr  = wb["Lönsamhetskontroll"]["C45"].value
    irr_krav = wb["Indata"]["C65"].value
    margin = (irr - irr_krav) if (irr is not None and irr_krav is not None) else None

    fails: list[str] = []
    if rent is None:
        fails.append("Resultat!D14 saknar cached value (kör LibreOffice-recalc?)")
    elif abs(rent - EXPECTED["binding_rent_kr_per_year"]) > TOLERANCE_RENT_KR:
        fails.append(
            f"Hyra: fick {rent:,.2f}, väntade {EXPECTED['binding_rent_kr_per_year']:,}"
        )
    if irr is None:
        fails.append("Lönsamhetskontroll!C45 saknar cached value")
    elif abs(irr - EXPECTED["actual_equity_irr"]) > TOLERANCE_PCT:
        fails.append(
            f"IRR EK: fick {irr:.4%}, väntade {EXPECTED['actual_equity_irr']:.2%}"
        )
    if margin is None:
        fails.append("IRR-marginal: ej beräkningsbar")
    elif abs(margin - EXPECTED["irr_margin_pp"]) > TOLERANCE_PCT:
        fails.append(
            f"IRR-marginal: fick {margin*100:+.2f} pp, väntade +{EXPECTED['irr_margin_pp']*100:.2f} pp"
        )

    return {
        "ok": not fails,
        "rent": rent,
        "irr": irr,
        "irr_krav": irr_krav,
        "margin": margin,
        "fails": fails,
    }


def main(argv: list[str] | None = None) -> int:
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    argv = argv if argv is not None else sys.argv[1:]
    target = Path(argv[0]) if argv else ITER8
    print(f"Regressionstest mot: {target.name}")
    res = check_baseline(target)
    if res["rent"] is not None:
        print(f"  Bindande kravhyra:  {res['rent']:>15,.2f} kr/år   (väntade {EXPECTED['binding_rent_kr_per_year']:,})")
    if res["irr"] is not None:
        print(f"  Faktisk IRR EK:     {res['irr']:>15.4%}              (väntade {EXPECTED['actual_equity_irr']:.2%})")
    if res["margin"] is not None:
        print(f"  Marginal mot krav:  {res['margin']*100:>+14.2f} pp           (väntade +{EXPECTED['irr_margin_pp']*100:.2f} pp)")
    if res["ok"]:
        print("\n✓ Regression OK")
        return 0
    print("\n✗ REGRESSION FAILED:")
    for f in res["fails"]:
        print(f"  - {f}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
