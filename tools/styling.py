"""Design system för Investeringskalkyl — finansmodell-standarder (XLSX-skill).

Färgkodning (branschstandard):
  BLÅ  text (0,0,255)   — hårdkodade indata, "Fyll i"-celler
  SVART text (0,0,0)    — formler och beräkningar
  GRÖN  text (0,128,0)  — cross-sheet-referenser

Rubrikpalett:
  PRIMARY  #1B3A6B — sektion-rubriker (mörkblå)
  SECONDARY #EBF0F8 — data-bakgrund (ljusblå)
  INPUT_BG  #EFF6FF — redigerbara fält
  STATUS_OK #E8F5E9 — ✓ Uppfyllt
"""
from __future__ import annotations
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.worksheet import Worksheet

# ── Palett ─────────────────────────────────────────────────────────────────
PRIMARY   = "1B3A6B"   # mörkblå — sektion-rubriker
WHITE     = "FFFFFF"
LIGHT     = "EBF0F8"   # ljusblå — data-bakgrund
INPUT_BG  = "EFF6FF"   # blåvit  — Fyll i-celler
STATUS_BG = "E8F5E9"   # ljusgrön — ✓-celler
RULE_CLR  = "C5D3E8"   # kantstreckets färg

# ── Textfärger (finansmodell-standard) ────────────────────────────────────
BLUE  = "0000FF"   # hårdkodad indata
GREEN = "008000"   # cross-sheet-referens
BLACK = "000000"   # formel / beräkning

# ── Talformat ──────────────────────────────────────────────────────────────
FMT_KR     = '#,##0 "kr";-#,##0 "kr";"-"'
FMT_KR_AR  = '#,##0 "kr/år";-#,##0 "kr/år";"-"'
FMT_KR_KVM = '#,##0 "kr/kvm";-#,##0 "kr/kvm";"-"'
FMT_PCT    = "0.0%"
FMT_PCT2   = "0.00%"
FMT_INT    = "#,##0"
FMT_AR     = '0 "år"'


# ── Hjälpfunktioner ────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _border_bottom(color: str = RULE_CLR) -> Border:
    side = Side(style="thin", color=color)
    return Border(bottom=side)

def _font(bold=False, size=11, color=BLACK, italic=False) -> Font:
    return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)

def _align(h="left", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# ── Stilapplikatorer ───────────────────────────────────────────────────────

def section_header(cell, level: int = 1) -> None:
    """Sektion-rubrik: mörkblå bakgrund, vit bold text."""
    if level == 1:
        cell.font = _font(bold=True, size=10, color=WHITE)
        cell.fill = _fill(PRIMARY)
    else:
        cell.font = _font(bold=True, size=10, color=PRIMARY)
        cell.fill = _fill(LIGHT)
    cell.alignment = _align(h="left", v="center")

def title_style(cell) -> None:
    cell.font = _font(bold=True, size=14, color=PRIMARY)
    cell.alignment = _align(h="left", v="center")

def subtitle_style(cell) -> None:
    cell.font = _font(bold=False, size=11, color="4A5568", italic=True)
    cell.alignment = _align(h="left", v="center")

def label_style(cell) -> None:
    cell.font = _font(bold=False, size=10, color="2D3748")
    cell.alignment = _align(h="left", v="center")

def input_style(cell) -> None:
    """Blå text + ljusblå bakgrund — användarens fält."""
    cell.font = _font(bold=False, size=10, color=BLUE)
    cell.fill = _fill(INPUT_BG)
    cell.alignment = _align(h="left", v="center")

def formula_style(cell) -> None:
    """Svart text — formel/beräkning."""
    cell.font = _font(bold=False, size=10, color=BLACK)
    cell.alignment = _align(h="right", v="center")

def crossref_style(cell) -> None:
    """Grön text — cross-sheet-referens."""
    cell.font = _font(bold=False, size=10, color=GREEN)
    cell.alignment = _align(h="right", v="center")

def crossref_bold(cell) -> None:
    cell.font = _font(bold=True, size=11, color=GREEN)
    cell.alignment = _align(h="right", v="center")

def status_ok_style(cell) -> None:
    cell.font = _font(bold=True, size=10, color="1A5E20")
    cell.fill = _fill(STATUS_BG)
    cell.alignment = _align(h="center", v="center")

def toc_link_style(cell) -> None:
    cell.font = Font(name="Calibri", bold=True, size=10,
                     color=PRIMARY, underline="single")
    cell.alignment = _align(h="left", v="center")

def toc_tag_style(cell) -> None:
    cell.font = _font(bold=False, size=10, color="6B7280", italic=True)
    cell.alignment = _align(h="left", v="center")

def toc_desc_style(cell) -> None:
    cell.font = _font(bold=False, size=9, color="6B7280")
    cell.alignment = _align(h="left", v="center", wrap=True)


# ── Oversikt ───────────────────────────────────────────────────────────────

def style_oversikt(ws: Worksheet) -> None:
    # Kolumnbredder
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 2
    ws.column_dimensions["E"].width = 24
    ws.column_dimensions["F"].width = 16

    # Radhöjder
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 16
    for r in [5, 9, 14, 19, 24, 30]:
        ws.row_dimensions[r].height = 20
    for r in [6, 7, 10, 11, 12, 15, 16, 17, 20, 21, 22, 25, 26, 27, 28]:
        ws.row_dimensions[r].height = 18

    # Titel
    title_style(ws["B2"])
    subtitle_style(ws["B3"])

    # A. PROJEKTINFORMATION
    section_header(ws["B5"])
    for r in [6, 7]:
        label_style(ws[f"B{r}"])
        crossref_style(ws[f"C{r}"])
        label_style(ws[f"E{r}"])
    input_style(ws["C6"])   # Fastighet — Fyll i
    input_style(ws["F6"])   # Utförd av — Fyll i
    crossref_style(ws["C7"])
    ws["C7"].number_format = "0"
    crossref_style(ws["F7"])
    ws["F7"].number_format = FMT_AR

    # B. KALKYLANTAGANDEN
    section_header(ws["B9"])
    for r in [10, 11, 12]:
        label_style(ws[f"B{r}"])
        crossref_style(ws[f"C{r}"])
        ws[f"C{r}"].number_format = FMT_PCT
    label_style(ws["E10"])
    label_style(ws["E11"])
    crossref_style(ws["F10"])
    crossref_style(ws["F11"])
    ws["F10"].number_format = FMT_PCT
    ws["F11"].number_format = FMT_PCT

    # C. BINDANDE KRAVHYRA
    section_header(ws["B14"])
    for r in [15, 16, 17]:
        label_style(ws[f"B{r}"])
        crossref_bold(ws[f"C{r}"])
        label_style(ws[f"E{r}"])
        crossref_style(ws[f"F{r}"])
    ws["C15"].number_format = FMT_KR_AR
    ws["C16"].number_format = FMT_KR_KVM
    ws["C17"].number_format = FMT_KR
    ws["F17"].number_format = FMT_KR_KVM

    # D. HYRESSPANN
    section_header(ws["B19"])
    for r in [20, 21, 22]:
        label_style(ws[f"B{r}"])
        crossref_style(ws[f"C{r}"])
        crossref_style(ws[f"E{r}"])
        ws[f"C{r}"].number_format = FMT_KR_AR
        ws[f"E{r}"].number_format = FMT_KR

    # E. LÖNSAMHETSKRAV
    section_header(ws["B24"])
    for r in [25, 26, 27]:
        label_style(ws[f"B{r}"])
        crossref_style(ws[f"C{r}"])
        ws[f"C{r}"].number_format = FMT_KR
        status_ok_style(ws[f"E{r}"])
    ws["C25"].number_format = FMT_KR   # NPV
    ws["C26"].number_format = FMT_KR   # NPV EK
    ws["C27"].number_format = FMT_KR   # MV diff
    label_style(ws["B28"])
    crossref_bold(ws["C28"])
    ws["C28"].number_format = FMT_PCT2
    label_style(ws["E28"])
    crossref_style(ws["F28"])
    ws["F28"].number_format = FMT_PCT2

    # F. INNEHÅLLSFÖRTECKNING
    section_header(ws["B30"])
    for r in [32, 34, 36, 38, 40, 42, 44]:
        toc_link_style(ws[f"B{r}"])
        toc_tag_style(ws[f"C{r}"])
        toc_desc_style(ws[f"E{r}"])
        ws.row_dimensions[r].height = 18


# ── Indata (sektion-rubriker och input-celler) ────────────────────────────

def style_indata(ws: Worksheet) -> None:
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 16

    # Kända sektion-rubriker i iter 8 Indata
    HEADERS = [3, 15, 24, 33, 51, 60, 67]
    for r in HEADERS:
        cell = ws[f"B{r}"]
        if cell.value:
            section_header(cell, level=2)

    # Input-celler (C-kolumnen, de flesta rader med siffervärden)
    for r in range(4, 70):
        c = ws[f"C{r}"]
        if c.value is not None and not (isinstance(c.value, str) and c.value.startswith("=")):
            input_style(c)


# ── Resultat (sektion-rubriker) ───────────────────────────────────────────

def style_resultat(ws: Worksheet) -> None:
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16

    for r in range(1, ws.max_row + 1):
        cell = ws[f"B{r}"]
        if isinstance(cell.value, str) and cell.value.isupper() and len(cell.value) > 4:
            section_header(cell, level=2)
