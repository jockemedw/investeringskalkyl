"""Iter 9 — bygger ovanpå iter 8 med §10-uppgifter.

Varje funktion = en §10-uppgift. main() kör i sekvens. Idempotent: kan
köras om utan att stega framåt mer än en gång (text-replaces matchar inte
redan-patchade celler).

Genomförda uppgifter (commit-historik förklarar):
- round E: 'år N' / 'år N+1' → 'år 20' / 'år 21' (explicit horisont)
- round B: pedagogisk omskrivning av Beräkningslogik-text (användarcentrerad)
- round C: Indata-fält renamning till branschterminologi
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


def round_c_indata_rename(wb: Workbook) -> int:
    """§10: Indata-fält renamning till branschterminologi.

    Tre celler byter etikett:
    - Indata!B18: 'Avskrivnings-% per år' → 'Avskrivningstakt'
    - Indata!B62: 'Långsiktig ränta' → 'Räntenivå' (matchar Finansiering!B14)
    - Lönsamhetskontroll!B26: 'Avskrivnings-%' → 'Avskrivningstakt'
    """
    renames = [
        ("Indata", "B18", "Avskrivnings-% per år", "Avskrivningstakt"),
        ("Indata", "B62", "Långsiktig ränta", "Räntenivå"),
        ("Lönsamhetskontroll", "B26", "Avskrivnings-%", "Avskrivningstakt"),
    ]
    n = 0
    for sheet, coord, old, new in renames:
        cell = wb[sheet][coord]
        if cell.value == old:
            cell.value = new
            n += 1
    return n


def round_b_pedagogisk(wb: Workbook) -> int:
    """§10 punkt 1: pedagogisk omskrivning av Beräkningslogik-text (round B).

    Inriktning: användarcentrerat — vad mallen ger användaren, inte bara
    metodbeskrivning. Behåller B33-44 (Goal Seek-jämförelse + fördelar) och
    block-rubriker B59/B121/B164/B179/B197 oförändrade.

    Körs efter round_e så 'år 20' i ny text inte stör replace-logiken.
    """
    new_text = {
        "B3": (
            "Du fyller i projektets förutsättningar i Indata. Mallen räknar baklänges "
            "till den årshyra som precis uppfyller respektive lönsamhetskrav: NPV ≥ 0, "
            "IRR EK ≥ avkastningskrav, samt marknadsvärde ≥ bokfört värde år 20. "
            "Den högsta av de tre kraven blir bindande kravhyra — golvet under vilket "
            "projektet inte är lönsamt. Lösningen är analytisk (ett steg, exakt) — "
            "beskrivet nedan."
        ),
        "B6":  "STEG 1 — Räkna NPV vid två testhyror (0 kr och 1 Mkr/år)",
        "B10": "STEG 2 — Härled hur mycket NPV ändras per krona hyra (lutningen b)",
        "B14": "STEG 3 — Lös linjärt: kravhyran är där NPV korsar noll, dvs −NPV(0) / b",
        "B47": (
            "Här under räknas de tre kravhyrorna ut. Block A ger NPV vid 0 kr hyra, "
            "block B vid 1 Mkr/år — differensen ger lutningen b för NPV-kravet. "
            "Block D-E gör samma för IRR EK. Block F gör marknadsvärdes-kravet. "
            "Detta är inte projektets faktiska kassaflöde — det visas på Kassaflöde-fliken "
            "vid bindande kravhyra."
        ),
    }
    ws = wb["Beräkningslogik"]
    for coord, text in new_text.items():
        ws[coord] = text
    return len(new_text)


def main() -> int:
    out = ROOT / "build" / "iter9.xlsx"
    out.parent.mkdir(exist_ok=True)

    print(f"1. Laddar {ITER8.name}")
    wb = load_iter(ITER8)

    print("2. Round E: 'år N' → 'år 20' / 'år N+1' → 'år 21'")
    n = round_e_year_n(wb)
    print(f"   {n} celler patchade")

    print("3. Round B: pedagogisk omskrivning av Beräkningslogik-text")
    n = round_b_pedagogisk(wb)
    print(f"   {n} celler patchade")

    print("4. Round C: Indata-fält renamning (branschterminologi)")
    n = round_c_indata_rename(wb)
    print(f"   {n} celler patchade")

    print(f"\n5. Sparar → {out.name}")
    save_iter(wb, out)

    print("\n6. Recalc (Excel COM / LibreOffice)")
    recalc(out)

    print("\n7. Regressionstest:")
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
