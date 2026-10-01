"""Steg 5: känslighet på restvärdet (yield ±1,00 pp) på flik 5 + redovisning på flik 1.

Flik 5 rad 53–57: tre scenarier (optimistisk/bedömt/pessimistisk) med direktavkastning,
exit-värde, IRR före skatt, NPV tillkommande och status mot IRR-kravet. Rad 59–61: scenario-
cashflöden = rad 37 + Δexit i exit-årets kolumn (vikten A7 skalar MV-delen). Bedömt == rad 37.
Flik 1: sju rader infogas efter kommentarsrutan i Marknadsituation (rad 64–70) med tabellen.
"""
from __future__ import annotations
from step4_bef import _set, _fmt, YEAR_COLS, K, KD, NPV, IRR, FS, N_PVDN, N_PVREI, N_RV, N_NPV

# rad 42–49 är en kollapsad outline-grupp i mallen — blocket läggs nedanför
S_HDR, S_COLS, S_OPT, S_BED, S_PES = 53, 54, 55, 56, 57
CF_OPT, CF_BED, CF_PES = 59, 60, 61
IRR_PRINT_LAST = 57
FS_INSERT = "64:70"          # nya rader på flik 1; gamla rad 64 (blank) → 71, Kalkylantaganden → 72
F_SPC1, F_TITLE, F_COLS, F_OPT, F_BED, F_PES, F_SPC2 = range(64, 71)
PCT = "0,00%"                # NumberFormatLocal — Excel på denna maskin tolkar NumberFormat lokalt


def _pct(sheet, ref):
    return {"op": "numfmt_local", "sheet": sheet, "ref": ref, "fmt": PCT}


def ops_kanslighet() -> list[dict]:
    o: list[dict] = []
    # ── Flik 5 ──
    o += [_fmt(IRR, "B14:W14", f"B{S_HDR}:W{S_HDR}"),
          _set(IRR, f"B{S_HDR}", v="KÄNSLIGHET RESTVÄRDE — direktavkastning ±1,00 pp (tillkommande flöde)"),
          _fmt(IRR, "B15:G15", f"B{S_COLS}:G{S_COLS}"),
          _set(IRR, f"B{S_COLS}", v="Scenario"), _set(IRR, f"C{S_COLS}", v="Direktavk."),
          _set(IRR, f"D{S_COLS}", f='="Exit-värde år "&$C$5'), _set(IRR, f"E{S_COLS}", v="IRR före skatt"),
          _set(IRR, f"F{S_COLS}", v="NPV tillk."), _set(IRR, f"G{S_COLS}", v="Mot krav"),
          {"op": "wrap", "sheet": IRR, "ref": f"B{S_COLS}:G{S_COLS}"}, {"op": "rowheight", "sheet": IRR, "row": S_COLS, "h": 30},
          _fmt(IRR, "B37:G37", f"B{S_OPT}:G{S_PES}"),
          _set(IRR, f"B{S_OPT}", v="Optimistisk (−1,00 pp)"),
          _set(IRR, f"B{S_BED}", v="Bedömt (Indata sektion 5)"),
          _set(IRR, f"B{S_PES}", v="Pessimistisk (+1,00 pp)"),
          _set(IRR, f"C{S_OPT}", f=f"={K}$M$11-0.01"), _set(IRR, f"C{S_BED}", f=f"={K}$M$11"),
          _set(IRR, f"C{S_PES}", f=f"={K}$M$11+0.01"),
          _pct(IRR, f"C{S_OPT}:C{S_PES}"), _pct(IRR, f"E{S_OPT}:E{S_PES}"),
          _fmt(IRR, "B37:BA37", f"B{CF_OPT}:BA{CF_PES}"),
          _set(IRR, f"B{CF_OPT}", v="Cashflow exkl skatt — optimistisk"),
          _set(IRR, f"B{CF_BED}", v="Cashflow exkl skatt — bedömt (= rad 37)"),
          _set(IRR, f"B{CF_PES}", v="Cashflow exkl skatt — pessimistisk"),
          {"op": "print_area", "sheet": IRR, "area": f"$A$1:$W${IRR_PRINT_LAST}"},
          ]
    for r, cf in ((S_OPT, CF_OPT), (S_BED, CF_BED), (S_PES, CF_PES)):
        o += [# exit-värde vid scenarioyield: MV-delen (C7, räknad på M11) skalas om
              _set(IRR, f"D{r}", f=f"=$C$7*{K}$M$11/C{r}"),
              _set(IRR, f"E{r}", f=f'=IFERROR(IRR(D{cf}:BA{cf}),"N/A")'),
              # NPV tillkommande vid scenarioyield: restvärdet skalas, PV med scenarioyield + inflation
              _set(IRR, f"F{r}", f=(f"='4. NPV'!$E${N_PVDN}+'4. NPV'!$E${N_PVREI}"
                                    f"+ROUND(ROUND('4. NPV'!$E${N_RV}*{K}$M$11/C{r}/10000,0)*10000"
                                    f"/((1+C{r}+{K}$M$9)^'4. NPV'!$E$47)/10000,0)*10000+'4. NPV'!$E$56")),
              _set(IRR, f"G{r}", f=f'=IF(ISNUMBER(E{r}),IF(E{r}>={K}$F$14,"✓ klarar krav","⚠ under krav"),"")'),
              ]
        for X in YEAR_COLS:
            o.append(_set(IRR, f"{X}{cf}",
                          f=f'=IF({X}$37="","",IF({X}$14=$C$5,{X}$37+$A$7*($D{r}-$C$7),{X}$37))'))
    # ── Flik 1: infoga rader efter kommentarsrutan i Marknadsituation ──
    # OBS: efter infogningen har alla rader ≥64 flyttats 7 steg — källor för fmt_from anges i NYA adresser.
    o += [{"op": "insert_rows", "sheet": FS, "ref": FS_INSERT},
          {"op": "clear_fmt", "sheet": FS, "ref": f"B{F_SPC1}:H{F_SPC2}"},     # ärvd blå fyllning från kommentarsrutan
          _fmt(FS, "B56:H56", f"B{F_SPC1}:H{F_SPC1}"), _fmt(FS, "B56:H56", f"B{F_SPC2}:H{F_SPC2}"),
          {"op": "rowheight", "sheet": FS, "row": F_SPC1, "h": 8.1}, {"op": "rowheight", "sheet": FS, "row": F_SPC2, "h": 8.1},
          _fmt(FS, "B55:H55", f"B{F_TITLE}:H{F_TITLE}"),
          _set(FS, f"B{F_TITLE}", f=(f'="Restvärdesbedömning — direktavkastning "&ROUND({K}$M$11*100,2)&" % '
                                     f'(extern "&ROUND({K}$M$6*100,2)&" % "&IF({K}$M$10>=0,"+","")&ROUND({K}$M$10*100,2)&" pp)"')),
          _set(FS, f"G{F_TITLE}", v=""),
          _fmt(FS, "B105:H105", f"B{F_COLS}:H{F_COLS}"),                        # = gamla rad 98 (IRR-raden, mörkt band)
          _set(FS, f"B{F_COLS}", v="Scenario"), _set(FS, f"C{F_COLS}", v="Direktavk."),
          _set(FS, f"D{F_COLS}", v="IRR före skatt"), _set(FS, f"E{F_COLS}", v="Exit-värde"),
          _set(FS, f"F{F_COLS}", f=f'="Krav "&ROUND({K}$F$14*100,2)&" %"'), _set(FS, f"G{F_COLS}", v="NPV tillkommande"),
          _fmt(FS, "B88:H88", f"B{F_OPT}:H{F_PES}"),                             # normal 11 pt (= gamla rad 81)
          _fmt(FS, "G107", f"E{F_OPT}:E{F_PES}"), _fmt(FS, "G107", f"G{F_OPT}:G{F_PES}"),   # #,##0 kr (= gamla rad 100)
          _pct(FS, f"C{F_OPT}:D{F_PES}"),
          {"op": "align", "sheet": FS, "ref": f"C{F_COLS}:G{F_PES}", "h": -4152},
          ]
    # B Scenario | C yield | D IRR | E exit | F status | G NPV  (kolumnbredder på flik 1 styr ordningen)
    for fr, sr in ((F_OPT, S_OPT), (F_BED, S_BED), (F_PES, S_PES)):
        for fcol, scol in (("B", "B"), ("C", "C"), ("D", "E"), ("E", "D"), ("F", "G"), ("G", "F")):
            o.append(_set(FS, f"{fcol}{fr}", f=f"='5. IRR'!{scol}{sr}"))
    return o
