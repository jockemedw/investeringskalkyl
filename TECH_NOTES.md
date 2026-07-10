# Tekniska anteckningar — openpyxl, XML-patch, regressionstest

Konkreta tekniska mönster från projektet, i kodbar form. Avsedd att kunna kopieras direkt in i nya scripts.

---

## 1. openpyxl XML-patch för `summaryBelow=False`

**Problem:** `outlinePr.summaryBelow = False` persisterar inte genom vanlig openpyxl-tilldelning. Excel öppnar arket med summary nedanför grupperna, inte ovanför.

**Lösning:** Patcha XML direkt i den sparade `.xlsx`-filen.

```python
import zipfile
import shutil
import re
from pathlib import Path

def patch_summary_below(xlsx_path: Path, sheet_filenames: list[str]):
    """
    Patcha sheet-XML så att outlinePr summaryBelow="0" (= False) sätts korrekt.
    
    sheet_filenames: t.ex. ["sheet1.xml", "sheet3.xml"] — bara de blad
                     som har row groups.
    """
    tmp_path = xlsx_path.with_suffix(".xlsx.tmp")
    
    with zipfile.ZipFile(xlsx_path, "r") as zin:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                base = item.split("/")[-1]
                if base in sheet_filenames:
                    text = data.decode("utf-8")
                    # Sätt summaryBelow="0" på outlinePr om det finns
                    if "<outlinePr" in text:
                        text = re.sub(
                            r'<outlinePr([^/>]*)summaryBelow="[01]"',
                            r'<outlinePr\1summaryBelow="0"',
                            text,
                        )
                        # Om attributet saknas, lägg till det
                        if 'summaryBelow="0"' not in text:
                            text = text.replace(
                                "<outlinePr",
                                '<outlinePr summaryBelow="0"',
                                1,
                            )
                    else:
                        # Lägg till hela outlinePr-elementet
                        text = text.replace(
                            "<sheetPr",
                            '<sheetPr><outlinePr summaryBelow="0"/>',
                            1,
                        )
                    data = text.encode("utf-8")
                zout.writestr(item, data)
    
    shutil.move(tmp_path, xlsx_path)
```

**Notera:** Detta måste köras EFTER `wb.save(...)`, eftersom det opererar på den sparade filen.

---

## 2. Row grouping (kollapserbara sektioner)

```python
from openpyxl.worksheet.worksheet import Worksheet

def add_collapsible_section(ws: Worksheet, start_row: int, end_row: int, level: int = 1):
    """
    Markerar rader [start_row, end_row] som collapsible på given outline level.
    Kombineras med patch_summary_below() för korrekt UX.
    """
    for row in range(start_row, end_row + 1):
        ws.row_dimensions[row].outlineLevel = level
        ws.row_dimensions[row].hidden = False  # Default: visad
```

**Användning:** Etappinmatning för investeringar — användaren kan vika ihop sektionen när den inte är aktuell.

---

## 3. Hyperlänkar (interna, mellan flikar)

```python
from openpyxl.worksheet.hyperlink import Hyperlink

def add_internal_link(cell, target_sheet: str, target_cell: str = "A1", display: str = None):
    """
    Lägger till hyperlänk till annan flik i samma arbetsbok.
    
    VIKTIGT: Använd Hyperlink-objekt, INTE sträng-tilldelning.
    """
    location = f"'{target_sheet}'!{target_cell}"
    cell.hyperlink = Hyperlink(
        ref=cell.coordinate,
        location=location,
        display=display or cell.value,
    )
    # Optional: ge länken länk-stil
    cell.style = "Hyperlink"
```

**FELAKTIGT (funkar inte):**
```python
cell.hyperlink = "'Indata'!A1"   # ← persisterar inte korrekt
```

---

## 4. Regressionstest med LibreOffice headless

```python
import subprocess
import openpyxl
from pathlib import Path

EXPECTED = {
    "binding_required_rent": 11_631_222,
    "actual_equity_irr": 0.0765,
    "irr_margin_pp": 0.0135,
}
TOLERANCE_RENT = 1.0           # kr
TOLERANCE_PERCENT = 0.0001     # 1 bp

def evaluate_with_libreoffice(xlsx_in: Path, xlsx_out: Path):
    """
    Tvinga LibreOffice att räkna alla formler och spara med cached values,
    så openpyxl kan läsa data_only=True.
    """
    subprocess.run(
        [
            "libreoffice", "--headless",
            "--calc", "--convert-to", "xlsx",
            "--outdir", str(xlsx_out.parent),
            str(xlsx_in),
        ],
        check=True,
        capture_output=True,
    )

def regression_check(xlsx_path: Path):
    evaluated = xlsx_path.with_name(xlsx_path.stem + "_evaluated.xlsx")
    evaluate_with_libreoffice(xlsx_path, evaluated)
    
    wb = openpyxl.load_workbook(evaluated, data_only=True)
    # Anpassa cellreferenserna till den faktiska placeringen i iter 8:
    rent = wb["Resultat"]["C5"].value           # exempel — verifiera!
    irr  = wb["Lönsamhetskontroll"]["D11"].value # exempel — verifiera!
    margin = wb["Lönsamhetskontroll"]["E11"].value
    
    fails = []
    if abs(rent - EXPECTED["binding_required_rent"]) > TOLERANCE_RENT:
        fails.append(f"Rent: got {rent}, expected {EXPECTED['binding_required_rent']}")
    if abs(irr - EXPECTED["actual_equity_irr"]) > TOLERANCE_PERCENT:
        fails.append(f"IRR: got {irr}, expected {EXPECTED['actual_equity_irr']}")
    if abs(margin - EXPECTED["irr_margin_pp"]) > TOLERANCE_PERCENT:
        fails.append(f"Margin: got {margin}, expected {EXPECTED['irr_margin_pp']}")
    
    if fails:
        raise AssertionError("Regression failed:\n  " + "\n  ".join(fails))
    print("✓ Regression OK")
```

**Cellreferenserna ovan är platshållare** — de exakta cellerna i iter 8 måste verifieras mot den faktiska filen. Hitta dem genom att söka på etikettsträngarna i fliken.

---

## 5. Indata i iter 8 — verifierat 2026-05-08

Värdena nedan är hämtade direkt ur `Investeringskalkyl_iter8.xlsx` och producerar baslinjen 11 631 222 kr/år / 7,65 % IRR EK.

### Sektion 1 — Marknads-/fastighetsförutsättningar (r4-13)
| Cell | Parameter | Värde |
|------|-----------|-------|
| C5  | Kalkylstart | 2026 |
| C6  | Direktavkastning marknad | 5,0 % |
| C7  | Kalkylränta driftnetto | 4,0 %  *(D-19 omprövar D-12)* |
| C8  | Inflation | 2,0 % |
| C9  | Kalkylränta restvärde | =C6+C8 (7,0 %) |
| C10 | Bolagsskatt | 20,6 % |
| C11 | Långsiktig vakansnivå | 0 % |
| C12 | Momsregistreringsgrad före | 100 % |
| C13 | Momsregistreringsgrad efter | 100 % |

### Sektion 2 — Investering & avskrivning (r15-22)
| Cell | Parameter | Värde |
|------|-----------|-------|
| C16 | Kalkylperiod | 20 år |
| C17 | Byggnadsvärde (avskrivningsbas) | 200 000 000 kr |
| C18 | Avskrivnings-% per år | 3,0 % |
| C19 | Avskrivningstid | =ROUND(1/C18,0) → 33 år |
| C20 | Fastighetsarea (mark) | 10 000 kvm |
| C21 | Markvärde | 500 kr/kvm |
| C22 | Minsta markvärde (totalt) | =C20*C21 (5 Mkr) |
| C23 | Hyresläge | "Beräkna" |

### Sektion 3 — Hyresobjekt (r24-31, kol B-R)
Bred matris med upp till 5 objekt. Kolumner: Objektnamn, Typ (Nyb/Bef), Bef area, Nyb area, Total area, Index, Andel, Avtalsstart, Avtalsslut, Hyra efter avtal, Vakans %, Prod start, Prod slut, Budget kr/kvm, Investering. Baslinjefall: 1 objekt "Skola" Nyb 5 000 kvm × 40 000 kr/kvm = 200 Mkr, indexandel 70 %, avtal 2026-2045.

### Sektion 4 — Drift & underhåll (r34-41) *(D-18 omprövar D-11)*
Matris kr/kvm/år. Defaults för Nybyggnad: FS 50, Rep 30, PU 40, Media 150, Övr 130 → summa 400 kr/kvm. Befintligt = 0 i baslinjen.

### Sektion 6 — Övriga poster (r52-58)
| Cell | Parameter | Värde |
|------|-----------|-------|
| C53 | Central administration | 80 kr/kvm |
| C54 | Fastighetsskatt (% av taxv) | 0 % |
| C55 | Taxeringsvärde | 0 kr |
| C56 | Tomträttsavgäld | 0 kr/år |
| C57 | Bokfört värde (befintligt) | 0 kr |
| C58 | Bokfört värde projektbelastning | 0 kr |

### Sektion 7 — Finansiering (r60-65)
| Cell | Parameter | Värde |
|------|-----------|-------|
| C61 | Byggnadskreditivränta | 3,0 % |
| C62 | Långsiktig ränta | 4,0 % |
| C63 | Amortering | 2,0 % /år |
| C64 | Belåningsgrad | 37 % |
| C65 | Avkastningskrav eget kapital | 6,3 % |

### Sektion 8 — Hyresspann (r67-68) *(D-16)*
| Cell | Parameter | Värde |
|------|-----------|-------|
| C68 | Intervall på investeringen (±) | 10 % |

### Nyckeloutput
| Cell | Vad | Cached value |
|------|-----|--------------|
| Resultat!D14 | Bindande kravhyra (mål) | 11 631 221,72 kr/år |
| Resultat!D16 | Per kvm/år | 2 326 kr |
| Lönsamhetskontroll!C45 | Faktisk IRR EK | 7,645 % |
| Indata!R31 | Total investering | 200 000 000 kr |
| Indata!F31 | Total area | 5 000 kvm |

---

## 6. Reverse engineering — Formelkatalog

Från fas 2 (~72 700 formler från LM 371 extraherade). Den fullständiga Formelkatalogen ligger sannolikt lokalt hos Joakim. Om den behöver återskapas:

```python
import openpyxl

wb = openpyxl.load_workbook("LM_371_Investeringskalkyl_2.xlsx", data_only=False)
catalog = []
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and isinstance(cell.value, str) and cell.value.startswith("="):
                catalog.append({
                    "sheet": sheet_name,
                    "cell": cell.coordinate,
                    "formula": cell.value,
                })

# Skriv till CSV / JSON / xlsx för analys
```

---

## 7. Build-script-skelett (rekommenderad struktur)

```
investeringskalkyl/
├── build.py                  # huvudentrypoint
├── sheets/
│   ├── oversikt.py
│   ├── indata.py
│   ├── kassafloede.py        # stannar vid driftnetto (D-05)
│   ├── finansiering.py       # block C (D-06)
│   ├── resultat.py
│   ├── lonsamhetskontroll.py
│   ├── beraekningslogik.py   # delta-metod (D-08)
│   └── dokumentation.py
├── styling/
│   └── theme.py              # färger, fonts, kantlinjer
├── postprocess/
│   ├── xml_patch.py          # summaryBelow-patchen (TECH §1)
│   └── hyperlinks.py
├── tests/
│   ├── test_regression.py    # 11 631 222 / 7,65 % / +1,35 pp
│   └── test_synthetic.py     # bug-fix-verifieringar (HANDOFF §6)
└── output/
    └── iter8.xlsx
```

## §9 — POLISH-fasens fallgropar (2026-07-10)

1. **Excel målar aldrig om rutnätet vid programmatisk scroll i bakgrundsfönster.**
   ScrollRow/Goto/Range.Select uppdaterar modellen (VisibleRange ändras) men
   pixlarna förblir öppningsvyn; zoom-ändringar målar däremot om. Lösning i
   `screenshot_sheets.py`: baka in scrollpositionen i en temporär kopia via
   XML-patch av sheetView/pane@topLeftCell (vid ren kolumnfrys styr
   sheetView@topLeftCell radscrollen — pane@topLeftCell räcker inte), så
   öppningsritningen hamnar rätt. Fönstret fångas med PrintWindow
   (PW_RENDERFULLCONTENT) — fungerar bakom andra fönster, stjäl inte fokus.

2. **PowerShell-stdout kan bli None trots exit 0.** OEM-kodade å/ä/ö i PS-output
   kraschar subprocess-lästråden (UnicodeDecodeError i tråd → res.stdout=None)
   — ser ut som intermittent flakiness. Fix: `[Console]::OutputEncoding = UTF8`
   i PS-skriptet + `encoding="utf-8", errors="replace"` i subprocess.run.
   Kritisk data går via fil, inte stdout.

3. **openpyxl rapporterar fantombredd 13 för intervall-lagrade kolumner.**
   Excel skriver `<col min="3" max="5" width="18"/>` — openpyxl expanderar inte
   intervallet vid läsning; `column_dimensions['D']` ger en NY dimension med
   defaultbredd 13. Läs XML:en vid tvivel.

4. **ignoredErrors saknar openpyxl-API.** Gröna felkontrollstrianglar släcks
   via XML-patch i `post_save` (v2_rounds): `<ignoredErrors>` efter colBreaks
   i schemaordningen (före drawing). Excel bevarar elementet vid recalc-save.

5. **xlsx-bladskydd blockerar outline-expandering** (D-24). COM:s
   EnableOutlining/UserInterfaceOnly persisteras inte i ren xlsx → flikar med
   kollapsade block lämnas oskyddade.

---

*Slut på tekniska anteckningar.*
