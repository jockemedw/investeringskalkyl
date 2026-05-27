"""Legacy-fasad mot tools.theme — alla stilar går via theme.apply(cell, role).

styling.py existerar bara för bakåtkompatibilitet med tidigare rounds.
Inga nya stilar — använd theme.apply() direkt.
"""
from __future__ import annotations
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment
from tools.theme import (
    apply, bottom_rule, clear_format,
    INK, MUTED, PAPER, SURFACE, ACCENT, POSITIVE, RULE, POSITIVE_TINT,
    FAMILY, SIZE_D, SIZE_H, SIZE_B, SIZE_C,
    FMT_KR, FMT_KR_AR, FMT_KR_KVM, FMT_INT, FMT_PCT1, FMT_PCT2, FMT_AR, FMT_M2,
    ROW_BANNER, ROW_SECTION, ROW_DATA, ROW_HERO, ROW_GAP,
)

# Exponerade alias för iter9.py (rensar successivt)
PRIMARY      = INK
SECONDARY    = INK         # eliminerad — peka till INK
NAVY_TEXT    = INK
WHITE        = PAPER
LIGHT        = SURFACE
INPUT_BG     = SURFACE
STATUS_BG    = POSITIVE_TINT
ACCENT_BG    = ACCENT
POSITIVE_BG  = POSITIVE_TINT
RULE_CLR     = RULE
BLUE         = INK
GREEN        = POSITIVE
BLACK        = INK
FONT_FAM     = FAMILY
FONT_FAM_BOLD = FAMILY      # samma family, bold via Font(bold=True)


# ── Style-applikatorer per flik ────────────────────────────────────────────

def style_oversikt(ws: Worksheet) -> None:
    """Översikt — banner + 6 sektioner + TOC. All styling via apply()."""
    # Kolumnbredder
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 2
    ws.column_dimensions["E"].width = 26
    ws.column_dimensions["F"].width = 20

    # Banner
    apply(ws["B2"], "title")
    ws.row_dimensions[2].height = ROW_BANNER
    apply(ws["B3"], "caption")
    ws.row_dimensions[3].height = ROW_GAP + 4

    # Sektion-headers
    for r in [5, 9, 14, 19, 24, 30]:
        apply(ws[f"B{r}"], "section")
        ws.row_dimensions[r].height = ROW_SECTION

    # Datarader: B=label, C/F=value, E=label (höger block)
    data_rows = {
        "block_a": [6, 7],                # PROJEKTINFORMATION
        "block_b": [10, 11, 12],          # KALKYLANTAGANDEN
        "block_c": [15, 16, 17],          # BINDANDE KRAVHYRA
        "block_d": [20, 21, 22],          # HYRESSPANN
        "block_e": [25, 26, 27, 28],      # LÖNSAMHETSKRAV
    }
    for rows in data_rows.values():
        for r in rows:
            apply(ws[f"B{r}"], "label")
            apply(ws[f"C{r}"], "value")
            apply(ws[f"E{r}"], "label")
            apply(ws[f"F{r}"], "value")
            ws.row_dimensions[r].height = ROW_DATA

    # Talformat på block_b (procent), block_c (kr), block_d (kr / kr), block_e (kr+pct)
    ws["C10"].number_format = FMT_PCT1; ws["F10"].number_format = FMT_PCT1
    ws["C11"].number_format = FMT_PCT1; ws["F11"].number_format = FMT_PCT1
    ws["C12"].number_format = FMT_PCT1
    ws["C15"].number_format = FMT_KR_AR
    ws["C16"].number_format = FMT_KR_KVM
    ws["C17"].number_format = FMT_KR
    ws["F17"].number_format = FMT_KR_KVM
    for r in [20, 21, 22]:
        ws[f"C{r}"].number_format = FMT_KR_AR
        ws[f"E{r}"].number_format = FMT_KR
    for r in [25, 26, 27]:
        ws[f"C{r}"].number_format = FMT_KR
    ws["C28"].number_format = FMT_PCT2
    ws["F28"].number_format = FMT_PCT2

    # Hero — bindande kravhyra
    apply(ws["B15"], "hero_label")
    apply(ws["C15"], "hero")
    ws["C15"].number_format = FMT_KR_AR
    ws.row_dimensions[15].height = ROW_HERO

    # Status-pillar på 25, 26, 27 (kolumn E)
    for r in [25, 26, 27]:
        apply(ws[f"E{r}"], "status")

    # TOC
    apply(ws["B30"], "section")
    for r in [32, 34, 36, 38, 40, 42, 44]:
        apply(ws[f"B{r}"], "value")  # länken stylas särskilt nedan
        ws[f"B{r}"].font = Font(name=FAMILY, size=SIZE_B, bold=True,
                                color=INK, underline="single")
        apply(ws[f"C{r}"], "caption")
        apply(ws[f"E{r}"], "caption")
        ws[f"E{r}"].alignment = Alignment(horizontal="left", vertical="center",
                                          wrap_text=True, indent=1)
        ws.row_dimensions[r].height = 30


def style_indata(ws: Worksheet) -> None:
    """Indata — sektion-rubriker + input-celler."""
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 16

    HEADERS = [3, 15, 24, 33, 51, 60, 67]
    for r in HEADERS:
        cell = ws[f"B{r}"]
        if cell.value:
            apply(cell, "section")
            ws.row_dimensions[r].height = ROW_SECTION

    # Input-celler i C-kolumnen
    for r in range(4, 70):
        cell = ws[f"C{r}"]
        if cell.value is not None and not (isinstance(cell.value, str)
                                            and cell.value.startswith("=")):
            apply(cell, "value")
            # Input-feel: tunn underline via bottom-rule
            bottom_rule(cell)
        # Etiketter i B-kolumnen
        b = ws[f"B{r}"]
        if b.value and r not in HEADERS:
            apply(b, "label")


def style_resultat(ws: Worksheet) -> None:
    """Resultat — sektion-rubriker + hero-rad."""
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18

    for r in range(1, ws.max_row + 1):
        cell = ws[f"B{r}"]
        if isinstance(cell.value, str) and cell.value.isupper() and len(cell.value) > 4:
            apply(cell, "section")
            ws.row_dimensions[r].height = ROW_SECTION


# ── Legacy-funktioner — dirigeras till apply() ────────────────────────────

def section_header(cell, level: int = 1) -> None:
    apply(cell, "section")

def title_style(cell) -> None:
    apply(cell, "title")

def subtitle_style(cell) -> None:
    apply(cell, "caption")

def label_style(cell) -> None:
    apply(cell, "label")

def input_style(cell) -> None:
    apply(cell, "value")
    bottom_rule(cell)

def formula_style(cell) -> None:
    apply(cell, "value")

def crossref_style(cell) -> None:
    apply(cell, "value")

def crossref_bold(cell) -> None:
    apply(cell, "value")
    cell.font = Font(name=FAMILY, size=SIZE_B, bold=True, color=INK)

def accent_value(cell) -> None:
    apply(cell, "hero")

def status_ok_style(cell) -> None:
    apply(cell, "status")

def toc_link_style(cell) -> None:
    cell.font = Font(name=FAMILY, size=SIZE_B, bold=True, color=INK, underline="single")
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

def toc_tag_style(cell) -> None:
    apply(cell, "caption")

def toc_desc_style(cell) -> None:
    apply(cell, "caption")
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
