"""Iter 9 — bygger ovanpå iter 8 med §10-uppgifter.

Varje funktion = en §10-uppgift. main() kör i sekvens. Idempotent: kan
köras om utan att stega framåt mer än en gång (text-replaces matchar inte
redan-patchade celler).

Genomförda uppgifter (commit-historik förklarar):
- round E: 'år N' / 'år N+1' → 'år 20' / 'år 21' (explicit horisont)
- round B: pedagogisk omskrivning av Beräkningslogik-text (användarcentrerad)
- round C: Indata-fält renamning till branschterminologi
- round F: designcleanup Lönsamhetskontroll (outline grouping av hjälprader)
- round G: Översikt om till beslutsdokument (projektinfo + kalkylantaganden)
"""
from __future__ import annotations
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from tools.patches import load_iter, save_iter, patch_summary_below, ITER8
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


def round_f_lonsamhetskontroll_cleanup(wb: Workbook) -> int:
    """§10 round F: Designcleanup av Lönsamhetskontroll.

    EK-cashflöde hjälprader (rad 43, 64-70) görs åtkomliga men visuellt
    separerade via Excel outline grouping (toggle [+] i radmarginalen).
    summaryBelow=False XML-patch körs av main() efter save_iter() så
    toggle hamnar OVANFÖR gruppen (rad 42 resp. rad 63).
    """
    ws = wb["Lönsamhetskontroll"]
    n = 0

    # Rad 43: etikett (synlig när expanderad), toggle hamnar på rad 42
    if ws["B43"].value is None:
        ws["B43"] = "EK-cashflöde år 0–20 (beräkningsunderlag för IRR ovan)"
        n += 1
    ws.row_dimensions[43].outlineLevel = 1
    ws.row_dimensions[43].hidden = True
    n += 1

    # Rad 63: synlig rubrik — indikator för gruppen nedanför
    if ws["B63"].value is None:
        ws["B63"] = "EK-cashflöde per scenario — klicka [+] för att expandera"
        n += 1

    # Rader 64-70: gruppera och kollapsa (toggle på rad 63)
    for row in range(64, 71):
        ws.row_dimensions[row].outlineLevel = 1
        ws.row_dimensions[row].hidden = True
    n += 7

    # Rader 35-41: tomma — dölj för att komprimera gap
    for row in range(35, 42):
        ws.row_dimensions[row].hidden = True
    n += 7

    return n


def round_g_oversikt_redesign(wb: Workbook) -> int:
    """§10 round G: Översikt om till beslutsdokument.

    Lägger till PROJEKTINFORMATION (redigerbara Fyll i-celler) och
    KALKYLANTAGANDEN (formler mot Indata) ovanför befintliga resultatblock.
    Innehållsförteckning skrivs om med korrekta hyperlänkar och förskjuts ner.
    Alla befintliga formler (Resultat, Lönsamhetskontroll) bevaras.
    """
    ws = wb["Översikt"]

    # Rensa rader 2-50 — unmerga först, sedan nolla värden
    for merged in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged))
    for r in range(2, 51):
        for c in range(1, 11):
            cell = ws.cell(r, c)
            cell.value = None
            cell.hyperlink = None

    def set_val(coord, val):
        ws[coord] = val

    # ── Titel ────────────────────────────────────────────────────────────────
    set_val("B2", "INVESTERINGSKALKYL — LEJONFASTIGHETER AB")
    set_val("B3", '=IF(Indata!B26="","(projekt ej namngivet)","Projekt: "&Indata!B26)')

    # ── A. PROJEKTINFORMATION ────────────────────────────────────────────────
    set_val("B5", "PROJEKTINFORMATION")
    set_val("B6", "Fastighet")
    set_val("C6", "Fyll i")
    set_val("E6", "Utförd av")
    set_val("F6", "Fyll i")
    set_val("B7", "Kalkylstart")
    set_val("C7", "=Indata!C5")
    set_val("E7", "Kalkylperiod")
    set_val("F7", '=Indata!C16&" år"')

    # ── B. KALKYLANTAGANDEN ──────────────────────────────────────────────────
    set_val("B9", "KALKYLANTAGANDEN")
    set_val("B10", "Kalkylränta driftnetto")
    set_val("C10", "=Indata!C7")
    set_val("E10", "IRR-krav EK")
    set_val("F10", "=Indata!C65")
    set_val("B11", "Direktavkastning marknad")
    set_val("C11", "=Indata!C6")
    set_val("E11", "Belåningsgrad")
    set_val("F11", "=Indata!C64")
    set_val("B12", "Inflation")
    set_val("C12", "=Indata!C8")

    # ── C. BINDANDE KRAVHYRA ─────────────────────────────────────────────────
    set_val("B14", "BINDANDE KRAVHYRA")
    set_val("B15", "Årshyra (mål-utfall)")
    set_val("C15", "=Resultat!D14")
    set_val("E15", "Bindande krav")
    set_val("F15", "=Resultat!D15")
    set_val("B16", "Per kvm/år")
    set_val("C16", "=Resultat!D16")
    set_val("E16", "Total area")
    set_val("F16", '=Indata!F31&" m²"')
    set_val("B17", "Total investering")
    set_val("C17", "=Indata!R31")
    set_val("E17", "Investering per kvm")
    set_val("F17", "=IFERROR(Indata!R31/Indata!F31,0)")

    # ── D. HYRESSPANN ────────────────────────────────────────────────────────
    set_val("B19", "HYRESSPANN VID INVESTERINGSUTFALL")
    set_val("B20", "Lägsta utfall (-X%)")
    set_val("C20", "=Resultat!C14")
    set_val("E20", "=Resultat!C7")
    set_val("B21", "Mål-utfall (bindande)")
    set_val("C21", "=Resultat!D14")
    set_val("E21", "=Resultat!D7")
    set_val("B22", "Högsta utfall (+X%)")
    set_val("C22", "=Resultat!E14")
    set_val("E22", "=Resultat!E7")

    # ── E. LÖNSAMHETSKRAV ────────────────────────────────────────────────────
    set_val("B24", "LÖNSAMHETSKRAV — STATUS VID BINDANDE KRAVHYRA")
    set_val("B25", "1. NPV ≥ 0")
    set_val("C25", "=Lönsamhetskontroll!C9")
    set_val("E25", "=Lönsamhetskontroll!D9")
    set_val("B26", "2. IRR EK ≥ avkastningskrav")
    set_val("C26", "=Lönsamhetskontroll!C18")
    set_val("E26", "=Lönsamhetskontroll!D18")
    set_val("B27", "3. MV år 20 ≥ Bokfört värde år 20")
    set_val("C27", "=Lönsamhetskontroll!C30")
    set_val("E27", "=Lönsamhetskontroll!D30")
    set_val("B28", "Faktisk IRR EK")
    set_val("C28", "=Lönsamhetskontroll!C45")
    set_val("E28", "Marginal mot krav")
    set_val("F28", "=Lönsamhetskontroll!C47")

    # ── F. INNEHÅLLSFÖRTECKNING ──────────────────────────────────────────────
    set_val("B30", "INNEHÅLLSFÖRTECKNING")

    toc = [
        (32, "Indata",              "Inmatning",            "Indata",
         "Marknadsförutsättningar, investering, hyresobjekt, drift, finansiering."),
        (34, "Resultat",            "Vad blir hyran?",      "Resultat",
         "Kravhyra för de tre lönsamhetskraven, bindande hyra och hyresspann."),
        (36, "Kassaflöde",          "Driftnetto år för år", "Kassaflöde",
         "Driftnetto-projektion vid bindande hyra (bruttohyra → driftnetto)."),
        (38, "Finansiering",        "Lån och avskrivning",  "Finansiering",
         "Lånebild: lån, amortering, räntekostnad, eget kapital, avskrivning."),
        (40, "Lönsamhetskontroll",  "Är kraven uppfyllda?", "Lönsamhetskontroll",
         "Verifiering av tre lönsamhetskrav, faktisk IRR EK och känslighetsanalys."),
        (42, "Beräkningslogik",     "Hur räknar mallen?",   "Beräkningslogik",
         "Delta-metoden + tekniska beräkningsblock A–F bakom kravhyran."),
        (44, "Dokumentation",       "Varför så här?",       "Dokumentation",
         "Designprinciper, val och paritetstest mot LM 371."),
    ]

    for row, name, tagline, target_sheet, desc in toc:
        cell = ws.cell(row, 2)
        cell.value = name
        cell.hyperlink = Hyperlink(
            ref=cell.coordinate,
            location=f"'{target_sheet}'!A1",
            display=name,
        )
        ws.cell(row, 3).value = tagline
        ws.cell(row, 5).value = desc

    return 1


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

    print("5. Round F: designcleanup Lönsamhetskontroll")
    n = round_f_lonsamhetskontroll_cleanup(wb)
    print(f"   {n} rader/celler patchade")

    print("6. Round G: Översikt om till beslutsdokument")
    n = round_g_oversikt_redesign(wb)
    print(f"   {n} sektioner skrivna")

    print(f"\n7. Sparar → {out.name}")
    save_iter(wb, out)

    print("\n8. XML-patch summaryBelow=False (Lönsamhetskontroll)")
    patch_summary_below(out, ["sheet6.xml"])
    print("   patch klar")

    print("\n9. Recalc (Excel COM / LibreOffice)")
    recalc(out)

    print("\n10. Regressionstest:")
    res = check_baseline(out)
    if res["rent"] is not None:
        print(f"    Hyra={res['rent']:,.2f}  IRR={res['irr']:.4%}  margin={res['margin']*100:+.2f} pp")
    if res["ok"]:
        print("    ✓ Regression OK")
        return 0
    print("    ✗ FAIL:", res["fails"])
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
