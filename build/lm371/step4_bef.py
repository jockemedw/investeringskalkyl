"""Steg 4: Bef = nollalternativ (beslut 8–10).

Flik 4 får ett block (rad 64–82) som särskiljer Bef-objektens driftnetto och räknar NPV på det
tillkommande flödet (krav 1). Blocket rad 47–62 behålls för hela fastigheten (krav 3).
IRR-flödet, CA-basen och Grafers MV-rad går på tillkommande. Befintligt bokfört värde (P9)
får egen återstående avskrivningstid (Q18).
"""
from __future__ import annotations
from openpyxl.utils import get_column_letter

KD, NPV, IRR, GRAF, FS, FIN = "2. Kalkyldata", "4. NPV", "5. IRR", "7.Grafer", "1. Framskrivningsunderlag", "3. Finansiering"
K = "'2. Kalkyldata'!"

B_HDR, B_AREA, B_INT, B_VAK, B_DOU, B_REI, B_DN_BEF, B_DN_TILLK, B_REI_TILLK, B_PV, B_PV_REI = range(64, 75)
N_HDR, N_PVDN, N_PVREI, N_RV, N_PVRV, N_PVTOT, N_NPV = range(76, 83)
NPV_PRINT_LAST = 82
YEAR_COLS = [get_column_letter(c) for c in range(4, 54)]      # D..BA
# (Bef-suffix, TioÅr/FemtonÅr-suffix) — namngivna områden i mallen har olika versalisering
DOU_NAMES = [("Fastighetsskötsel", "Fastighetsskötsel"), ("Reparationer", "Reparationer"),
             ("Planeratunderhåll", "PlaneratUnderhåll"), ("Media", "Media"),
             ("Övrigaförvaltningskostnader", "ÖvrigaFörvaltningskostnader")]


def _set(sheet, ref, *, f=None, v=None):
    d = {"op": "set", "sheet": sheet, "ref": ref}
    if f is not None:
        d["formula"] = f
    else:
        d["value"] = v
    return d


def _fmt(sheet, src, dst):
    return {"op": "fmt_from", "sheet": sheet, "src": src, "dst": dst}


def _rnd(expr: str) -> str:
    """Avrundning till 10 000 utan MROUND (MROUND ger #NUM! för negativa tal)."""
    return f"ROUND(({expr})/10000,0)*10000"


def _year_formula(template: str, X: str) -> str:
    """Byt kolumnplatshållaren 'X' mot årskolumnens bokstav. 'X' följs alltid av $ eller siffra."""
    out = []
    i = 0
    while i < len(template):
        ch = template[i]
        nxt = template[i + 1] if i + 1 < len(template) else ""
        if ch == "X" and (nxt == "$" or nxt.isdigit()):
            out.append(X)
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def ops_bef() -> list[dict]:
    o: list[dict] = []
    # ── Flik 2: återstående avskrivningstid för befintligt bokfört värde (beslut 10) ──
    o += [_fmt(KD, "N17", "N13"), _set(KD, "N13", v="Avskr.tid befintligt"),
          _fmt(KD, "P17", "P13"), _set(KD, "P13", f="=IF(Q13<>0,Q13,P16)"),
          _fmt(KD, "Q17", "Q13"),
          {"op": "replace", "sheet": FIN, "ref": "D40:BA40",
           "find": "'2. Kalkyldata'!$P$16", "repl": "'2. Kalkyldata'!$P$13"},
          _fmt(KD, "C42", "I18:M18"), {"op": "merge", "sheet": KD, "ref": "I18:M18"},
          {"op": "wrap", "sheet": KD, "ref": "I18"}, {"op": "rowheight", "sheet": KD, "row": 18, "h": 30},
          _set(KD, "I18", f='=IF(COUNTIF(E23:E27,"Bef")+(P9>0)>0,'
                            '"Bef-objekt och befintligt bokfört värde ingår bara i krav 3 (hela fastigheten). '
                            'Krav 1–2 räknar tillkommande.","")'),
          {"op": "italic", "sheet": KD, "ref": "I18"},
          ]
    # ── Flik 4: block tillkommande vs hela ──
    o += [_fmt(NPV, "B47:E47", f"B{B_HDR}:E{B_HDR}"),
          _set(NPV, f"B{B_HDR}", v="TILLKOMMANDE FLÖDE (Bef = nollalternativ)"),
          _fmt(NPV, "B20", f"B{B_AREA}:B{B_PV_REI}"), _fmt(NPV, "D20:BA20", f"D{B_AREA}:BA{B_PV_REI}"),
          _fmt(NPV, "B29", f"B{B_DN_TILLK}"), _fmt(NPV, "D29:BA29", f"D{B_DN_TILLK}:BA{B_DN_TILLK}"),
          {"op": "colwidth_from", "sheet": NPV, "src": "C", "dst": "E"},
          _set(NPV, f"B{B_AREA}", v="Area Bef-objekt"), _set(NPV, f"B{B_INT}", v="Intäkter Bef-objekt"),
          _set(NPV, f"B{B_VAK}", v="Vakans Bef-objekt"), _set(NPV, f"B{B_DOU}", v="Drift och underhåll Bef-objekt"),
          _set(NPV, f"B{B_REI}", v="Re-investering Bef-objekt"), _set(NPV, f"B{B_DN_BEF}", v="Driftnetto Bef-objekt"),
          _set(NPV, f"B{B_DN_TILLK}", v="Driftnetto tillkommande"),
          _set(NPV, f"B{B_REI_TILLK}", v="Re-investering tillkommande"),
          _set(NPV, f"B{B_PV}", v="PV tillkommande"), _set(NPV, f"B{B_PV_REI}", v="PV Extraordinärt UH tillkommande"),
                    _set(NPV, f"B{N_HDR - 1}", v="Utan Bef-rader är tillkommande = hela fastigheten. "
                                       "Vikter 0/0,4/0,6 på 5. IRR återskapar den tidigare mallens IRR."),
          {"op": "italic", "sheet": NPV, "ref": f"B{N_HDR - 1}"},
          ]
    bef = f'--({K}$E$23:$E$27="Bef")'
    dou_terms = "+".join(
        f"Bef{bs}*(1+IF($D$4+Höjnin1<=X$4,TioÅr{ts},0)+IF($D$4+Höjning2<=X$4,FemtonÅr{ts},0))"
        for bs, ts in DOU_NAMES)
    templates = {
        B_AREA: f'=IF(X$4="",0,SUMIFS({K}$H$23:$H$27,{K}$E$23:$E$27,"Bef",{K}$M$23:$M$27,X$3))',
        B_INT: f'=IF(X$4="",0,SUMPRODUCT({bef},X12:X16))',
        B_VAK: (f'=IF(X$4="",0,-SUMPRODUCT({bef},1-(X$4>={K}$M$23:$M$27)*(X$4<={K}$N$23:$N$27),'
                f'X12:X16,{K}$P$23:$P$27))'),
        B_DOU: f'=IF(X$4="",0,-X{B_AREA}*(1+Inflation)^(X$5-1)*({dou_terms}))',
        B_REI: (f'=IF(X$4="",0,-SUMPRODUCT({bef},SUMIFS({K}$W$34:$W$39,{K}$R$34:$R$39,X$4,'
                f'{K}$Q$34:$Q$39,{K}$C$23:$C$27)))'),
        B_DN_BEF: f"=X{B_INT}+X{B_VAK}+X{B_DOU}+X{B_REI}",
        B_DN_TILLK: f"=X29-X{B_DN_BEF}",
        B_REI_TILLK: f"=X25-X{B_REI}",
        B_PV: f"=(X{B_DN_TILLK}-X{B_REI_TILLK})/((1+{K}$M$7)^X$5)",
        B_PV_REI: f"=-(X{B_REI_TILLK}/((1+{K}$M$7)^X$5))",
    }
    for X in YEAR_COLS:
        for r, tmpl in templates.items():
            o.append(_set(NPV, f"{X}{r}", f=_year_formula(tmpl, X)))
        # IRR: central administration på tillkommande area (Bef-arean är nollalternativets kostnad).
        # Fas 2 F2: exponent = årsindex − 1 (mallen hade ^0 år 1 och ^2 år 2).
        ca = f"=IF({X}15=\"\",0,-{K}$P$14*('4. NPV'!{X}$8-'4. NPV'!{X}${B_AREA}))"
        if X != "D":
            ca += f"*(1+{K}$M$9)^('5. IRR'!{X}14-1)"      # fas 2 F2: år 2 = ^1 (var ^2)
        o.append(_set(IRR, f"{X}26", f=ca))
    # ── NPV-sammanfattning tillkommande (krav 1), 10 år (C) / kalkylperiod (E) — speglar rad 47–59 ──
    idx_dn, idx_rei = B_DN_TILLK - 3, B_REI_TILLK - 3      # HLOOKUP-index i $D$4:$BA$74
    o += [_fmt(NPV, "B47:E47", f"B{N_HDR}:E{N_HDR}"),
          _set(NPV, f"B{N_HDR}", v="NPV TILLKOMMANDE, krav 1"),
          _set(NPV, f"C{N_HDR}", f="=C47"), _set(NPV, f"E{N_HDR}", f="=E47"),
          _fmt(NPV, "B48:E48", f"B{N_PVDN}:E{N_PVTOT}"), _fmt(NPV, "B59:E59", f"B{N_NPV}:E{N_NPV}"),
          _set(NPV, f"B{N_PVDN}", v="PV Driftnetto tillkommande"), _set(NPV, f"B{N_PVREI}", v="PV Extraordinärt UH"),
          _set(NPV, f"B{N_RV}", v="Restvärde tillkommande"), _set(NPV, f"B{N_PVRV}", v="PV Restvärde"),
          _set(NPV, f"B{N_PVTOT}", v="PV TOTALT tillkommande"), _set(NPV, f"B{N_NPV}", v="NPV tillkommande"),
          _set(NPV, f"C{N_PVDN}", f="=" + _rnd(f"SUMIFS($D${B_PV}:$BA${B_PV},$D$4:$BA$4,$C$44)")),
          _set(NPV, f"E{N_PVDN}", f="=" + _rnd(f"SUMIFS($D${B_PV}:$BA${B_PV},$D$4:$BA$4,$D$44)")),
          _set(NPV, f"C{N_PVREI}", f="=-" + _rnd(f"SUM($D${B_PV_REI}:$BA${B_PV_REI})")),
          _set(NPV, f"E{N_PVREI}", f="=-" + _rnd(f"SUM($D${B_PV_REI}:$BA${B_PV_REI})")),
          _set(NPV, f"C{N_RV}", f="=" + _rnd(f"HLOOKUP(C$43+1,$D$4:$BA${B_PV_REI},{idx_dn},FALSE)/{K}$M$11")),
          _set(NPV, f"E{N_RV}", f="=" + _rnd(f"(HLOOKUP(D$43+1,$D$4:$BA${B_PV_REI},{idx_dn},FALSE)"
                                            f"-HLOOKUP(D$43+1,$D$4:$BA${B_PV_REI},{idx_rei},FALSE))/{K}$M$11")),
          _set(NPV, f"C{N_PVRV}", f="=" + _rnd(f"C{N_RV}/((1+{K}$M$8)^(C47))")),
          _set(NPV, f"E{N_PVRV}", f="=" + _rnd(f"E{N_RV}/((1+{K}$M$8)^(E47))")),
          _set(NPV, f"C{N_PVTOT}", f=f"=IFERROR(C{N_PVDN}+C{N_PVREI}+C{N_PVRV},0)"),
          _set(NPV, f"E{N_PVTOT}", f=f"=IFERROR(E{N_PVDN}+E{N_PVREI}+E{N_PVRV},0)"),
          _set(NPV, f"C{N_NPV}", f=f"=C{N_PVTOT}+$C$56"), _set(NPV, f"E{N_NPV}", f=f"=E{N_PVTOT}+$E$56"),
          _set(NPV, "B47", v="NET PRESENT VALUE (NPV) — hela fastigheten, används i krav 3 (PV mot bokfört värde)"),
          {"op": "print_area", "sheet": NPV, "area": f"$B$1:$Y${NPV_PRINT_LAST}"},
          # Krav 1 och flik 1 läser tillkommande
          _set(KD, "D13", f=f"=IF(G13<>\"\",\"\",'4. NPV'!$E${N_NPV}/1000)"),
          _set(FS, "G100", f=f"='4. NPV'!$E${N_NPV}"), _set(FS, "K100", f=f"='4. NPV'!$E${N_NPV}"),
          _set(FS, "G101", f=f"='4. NPV'!$E${N_PVTOT}"), _set(FS, "K101", f=f"='4. NPV'!$E${N_PVTOT}"),
          _set(FS, "B103", v="PV jmf bokfört värde inkl investering (hela fastigheten)"),
          # IRR och Grafer på tillkommande driftnetto
          {"op": "replace", "sheet": IRR, "ref": "D23:BA23", "find": "$29)", "repl": f"${B_DN_TILLK})"},
          _set(IRR, "B23", v="Driftnetto tillkommande"),
          {"op": "replace", "sheet": GRAF, "ref": "C40:BA40", "find": "$29-", "repl": f"${B_DN_TILLK}-"},
          {"op": "replace", "sheet": GRAF, "ref": "C40:BA40", "find": "$25)", "repl": f"${B_REI_TILLK})"},
          _set(GRAF, "B40", v="Driftnetto tillkommande"),
          ]
    return o
