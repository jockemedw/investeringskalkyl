"""LM 371 v3 — kirurgisk patch av mallen via Excel COM.

Källa:  build/lm371/source/LM371_source.xlsx  (orörd mall, ifylld med Skola-exemplet)
Utdata: build/lm371/LM371_v3.xlsx

Kör:  python build/lm371/build.py
Fas 1 steg 1–3: restvärdesmodell (justerad direktavkastning, rent MV), investeringsmatris.
Se ANALYS_OMTAG.md för beslut 1–10.
"""
from __future__ import annotations
import io
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from comrun import run  # noqa: E402

SOURCE = HERE / "source" / "LM371_source.xlsx"
OUT = HERE / "LM371_v3.xlsx"

KD = "2. Kalkyldata"
NPV = "4. NPV"
IRR = "5. IRR"
GRAF = "7.Grafer"
FS = "1. Framskrivningsunderlag"

# Excel-konstanter
LEFT, CENTER, RIGHT = -4131, -4108, -4152

# 14 kostnadsposter — samma lista som de gamla rutorna (AF5:AF18)
POSTER = [
    "Tomt- och fastighetskostnader", "Projekteringskostnader", "Byggherrekostnader",
    "Entreprenadkostnader", "Tillkommande arbeten", "Avgifter & byggherrekostnader",
    "Finansiering", "Index", "Övriga kostnader", "Bidrag och intäkter",
    "Ersättningslokaler", "Övrigt 1", "Övrigt 2", "Övrigt 3",
]
OBJ_COLS = ["G", "H", "I", "J", "K"]          # matrisens objektkolumner
OBJ_ROWS = [23, 24, 25, 26, 27]               # motsvarande objektrader i hyrestabellen
SKALA = ["Mycket positivt", "Något positivt", "Neutralt", "Något negativt", "Mycket negativt"]
STEG = 0.0025                                 # pp per steg (beslut 3)

# Radlayout nya sektioner på flik 2 (print_area var C3:U51; rad 42–50 = antaganden-text)
S4_HDR, S4_TBL = 52, 54
S4_FIRST = S4_TBL + 1                          # 55
S4_LAST = S4_FIRST + len(POSTER) - 1           # 68
S4_SUM, S4_KVM = S4_LAST + 1, S4_LAST + 2      # 69, 70
S5_HDR, S5_TBL = 72, 74
S5_FIRST, S5_LAST = 75, 78
S5_SUM, S5_ADJ = 79, 80
PRINT_LAST = 81


def _set(sheet, ref, *, f=None, v=None):
    d = {"op": "set", "sheet": sheet, "ref": ref}
    if f is not None:
        d["formula"] = f
    else:
        d["value"] = v
    return d


def _fmt(sheet, src, dst):
    return {"op": "fmt_from", "sheet": sheet, "src": src, "dst": dst}


def ops_restvarde() -> list[dict]:
    """Steg 2–3: justerad direktavkastning + ompekning + vikter."""
    o: list[dict] = []
    # ── Flik 2, förutsättningar marknad: två nya rader under kalkylräntorna ──
    o += [_fmt(KD, "K6:L6", "K10:L10"), _fmt(KD, "K6:L6", "K11:L11"),
          {"op": "merge", "sheet": KD, "ref": "K10:L10"}, {"op": "merge", "sheet": KD, "ref": "K11:L11"},
          _set(KD, "K10", v="Justering direktavkastning"),
          _set(KD, "K11", v="Justerad direktavkastning"),
          _fmt(KD, "M8", "M10:M11"),
          _set(KD, "M10", f=f"=$H${S5_SUM}"),
          _set(KD, "M11", f="=$M$6+$M$10"),
          # Kalkylränta restvärde följer justerad yield (beslut 2)
          _set(KD, "M8", f="=$M$11+$M$9"),
          # Varningsrutan: visa yield-justering bredvid dagens restvärdesvarning
          _set(KD, "G14", f='=IF(\'5. IRR\'!C11<>"","Eget restvärde ifyllt i 5. IRR. ","")'
                            '&IF(M10<>0,"Yield "&IF(M10>0,"+","")&ROUND(M10*100,2)&" pp → "&ROUND(M11*100,2)&" %","")'),
          # Tvingande-vakten: S23:S27 är nu formler (COUNTA räknar dem som ifyllda) → explicit
          # kontroll att varje icke-Bef-objekt med area har en investering (matris eller U-överstyrning)
          _set(KD, "G13", f='=IF(COUNTA(M3,M6,M7,M9,Q15,Q17,C23:C27,E23:E27,F23:F27,G23:G27,I23:I27,J23:J27,'
                            'M23:M27,N23:N27,O23:O27,Q23:Q27,S23:S27,R23:R27)<17,"Fyll i alla tvingande fält",'
                            'IF(SUMPRODUCT(--($H$23:$H$27>0),--($E$23:$E$27<>"Bef"),--($T$23:$T$27=0))>0,'
                            '"Investering saknas för objekt med area (sektion 4)",'
                            'IF($AD$12<>1,"Hyresmultiplikator AD12 ≠ 1","")))'),
          {"op": "replace", "sheet": KD, "ref": "D13:D15", "find": 'G13="Fyll i alla tvingande fält"', "repl": 'G13<>""'},
          ]
    # ── Sektion 5: Restvärdesbedömning ──
    o += [_fmt(KD, "C31", f"C{S5_HDR}"), _set(KD, f"C{S5_HDR}", v="RESTVÄRDESBEDÖMNING"),
          _fmt(KD, "A33", f"A{S5_TBL}"), _set(KD, f"A{S5_TBL}", v=5),
          {"op": "merge", "sheet": KD, "ref": f"A{S5_TBL}:A{S5_ADJ}"},
          _fmt(KD, "C33", f"C{S5_TBL}:F{S5_TBL}"), _fmt(KD, "G33", f"G{S5_TBL}:U{S5_TBL}"),
          _set(KD, f"C{S5_TBL}", v="Faktor"),
          _set(KD, f"G{S5_TBL}", v="Bedömning"), _set(KD, f"H{S5_TBL}", v="Effekt"),
          _set(KD, f"I{S5_TBL}", v="Motivering"),
          _set(KD, f"C{S5_HDR + 1}", v="Extern direktavkastning (M6) är utgångsläget. Bedömningarna nedan justerar "
                                        "den ±0,25/±0,50 pp per faktor och används i alla tre kraven."),
          {"op": "italic", "sheet": KD, "ref": f"C{S5_HDR + 1}"},
          # skala för dropdown i hjälpområdet (gamla rutorna rensas i ops_matris)
          _set(KD, "AC4", v="Skala restvärdesbedömning"),
          ]
    for i, s in enumerate(SKALA):
        o.append(_set(KD, f"AC{5 + i}", v=s))
    params = [
        ("Läge (centralitet, infrastruktur)", "Kollektivtrafik, serviceunderlag, demografi, planerad infrastruktur"),
        ("Långsiktig vakansrisk", "Hyresgästens beroende av lokalen, alternativa hyresgäster"),
        ("Lokalflexibilitet (omställbarhet)", "Planlösning, bjälklagshöjd, bärande stomme, installationskapacitet"),
        ("Byggnadsteknisk standard", "Stomme, klimatskal, installationer, energiprestanda mot framtida krav"),
    ]
    for i, (label, hint) in enumerate(params):
        r = S5_FIRST + i
        o += [_fmt(KD, "C34:F34", f"C{r}:F{r}"), _set(KD, f"C{r}", v=label),
              _fmt(KD, "M6", f"G{r}"), _set(KD, f"G{r}", v="Neutralt"), {"op": "align", "sheet": KD, "ref": f"G{r}", "h": CENTER},
              {"op": "dv_list", "sheet": KD, "ref": f"G{r}", "source": "=$AC$5:$AC$9"},
              _fmt(KD, "M8", f"H{r}"),
              _set(KD, f"H{r}", f=f"=IFERROR((MATCH(G{r},$AC$5:$AC$9,0)-3)*{STEG},0)"),
              _fmt(KD, "S23", f"I{r}:U{r}"), {"op": "merge", "sheet": KD, "ref": f"I{r}:U{r}"},
              {"op": "align", "sheet": KD, "ref": f"I{r}", "h": LEFT},
              _set(KD, f"I{r}", v=""),
              _set(KD, f"V{r}", v=hint), {"op": "italic", "sheet": KD, "ref": f"V{r}"},
              ]
    o += [_fmt(KD, "C34:F34", f"C{S5_SUM}:F{S5_SUM}"), _set(KD, f"C{S5_SUM}", v="Summa justering av direktavkastning"),
          {"op": "bold", "sheet": KD, "ref": f"C{S5_SUM}"},
          _fmt(KD, "M8", f"H{S5_SUM}"), _set(KD, f"H{S5_SUM}", f=f"=SUM(H{S5_FIRST}:H{S5_LAST})"),
          {"op": "bold", "sheet": KD, "ref": f"H{S5_SUM}"},
          _fmt(KD, "C34:F34", f"C{S5_ADJ}:F{S5_ADJ}"), _set(KD, f"C{S5_ADJ}", v="Justerad direktavkastning (M11) — används i NPV, PV-krav och IRR"),
          _fmt(KD, "M8", f"H{S5_ADJ}"), _set(KD, f"H{S5_ADJ}", f="=$M$11"),
          ]
    # ── Ompekning av yield: alla ställen som läste M6 för marknadsvärde ──
    find, repl = "'2. Kalkyldata'!$M$6", "'2. Kalkyldata'!$M$11"
    o += [{"op": "replace", "sheet": GRAF, "ref": "C44:BA44", "find": find, "repl": repl},
          {"op": "replace", "sheet": NPV, "ref": "C50:E50", "find": find, "repl": repl},
          {"op": "replace", "sheet": FS, "ref": "G68:K68", "find": find, "repl": repl},
          _set(FS, "B68", v="Direktavkastningskrav, marknad (justerad)"),
          ]
    # ── Flik 5: vikter 0/1/0 = rent marknadsvärde (beslut 1), kvar som inmatning ──
    o += [_set(IRR, "A7", v=1), _set(IRR, "A8", v=0),
          _set(IRR, "B7", f="=\"Nettokap marknadsvärde, yield \"&ROUND('2. Kalkyldata'!$M$11*100,2)&\" %\""),
          _set(IRR, "B13", v="Vikter (kolumn A): 0 / 1 / 0 = rent marknadsvärde vid justerad direktavkastning (v3). "
                             "0 / 0,4 / 0,6 återger tidigare viktning mot bokfört värde."),
          {"op": "italic", "sheet": IRR, "ref": "B13"},
          ]
    return o


def ops_matris() -> list[dict]:
    """Steg 2b: investeringsspecifikation som matris, automatiskt länkad till S23:S27."""
    o: list[dict] = [
        _fmt(KD, "C31", f"C{S4_HDR}"), _set(KD, f"C{S4_HDR}", v="INVESTERINGSSPECIFIKATION"),
        _set(KD, f"C{S4_HDR + 1}", v="Projektbudget per objekt i dagens penningvärde, kr. Summan styr "
                                     "Budget kr/kvm (S23:S27). Manuell överstyrning görs i Justering (U23:U27)."),
        {"op": "italic", "sheet": KD, "ref": f"C{S4_HDR + 1}"},
        _fmt(KD, "A33", f"A{S4_TBL}"), _set(KD, f"A{S4_TBL}", v=4),
        {"op": "merge", "sheet": KD, "ref": f"A{S4_TBL}:A{S4_KVM}"},
        _fmt(KD, "C33", f"C{S4_TBL}:F{S4_TBL}"), _fmt(KD, "G33", f"G{S4_TBL}:L{S4_TBL}"),
        _set(KD, f"C{S4_TBL}", v="Kostnadspost"), _set(KD, f"L{S4_TBL}", v="Totalt"),
    ]
    for col, r in zip(OBJ_COLS, OBJ_ROWS):
        o.append(_set(KD, f"{col}{S4_TBL}", f=f'=IF($C${r}="","Objekt {r - 22}",$C${r})'))
    for i, post in enumerate(POSTER):
        r = S4_FIRST + i
        o += [_fmt(KD, "C34:F34", f"C{r}:F{r}"), _set(KD, f"C{r}", v=post),
              _fmt(KD, "S23", f"G{r}:K{r}"),
              _fmt(KD, "T23", f"L{r}"), _set(KD, f"L{r}", f=f"=SUM(G{r}:K{r})"),
              ]
    o += [_fmt(KD, "C34:F34", f"C{S4_SUM}:F{S4_SUM}"), _set(KD, f"C{S4_SUM}", v="Summa investering"),
          {"op": "bold", "sheet": KD, "ref": f"C{S4_SUM}:L{S4_SUM}"},
          _fmt(KD, "C34:F34", f"C{S4_KVM}:F{S4_KVM}"), _set(KD, f"C{S4_KVM}", v="kr/kvm (BRA)"),
          ]
    for col, r in zip(OBJ_COLS, OBJ_ROWS):
        o += [_fmt(KD, "T23", f"{col}{S4_SUM}"), _set(KD, f"{col}{S4_SUM}", f=f"=SUM({col}{S4_FIRST}:{col}{S4_LAST})"),
              _fmt(KD, "T23", f"{col}{S4_KVM}"), _set(KD, f"{col}{S4_KVM}", f=f"=IFERROR({col}{S4_SUM}/$H{r},0)"),
              # länken: Budget kr/kvm = matrisens summa / area. Blå inmatning flyttar till matrisen.
              _fmt(KD, "T23", f"S{r}"), _set(KD, f"S{r}", f=f"=IFERROR({col}${S4_SUM}/$H{r},0)"),
              ]
    o += [_fmt(KD, "T23", f"L{S4_SUM}"), _set(KD, f"L{S4_SUM}", f=f"=SUM(G{S4_SUM}:K{S4_SUM})"),
          _fmt(KD, "T23", f"L{S4_KVM}"), _set(KD, f"L{S4_KVM}", f=f"=IFERROR(L{S4_SUM}/$H$28,0)"),
          # Exemplet: 240 Mkr låg hårdkodat i S23 — läggs på entreprenadraden för objekt 1
          _set(KD, f"G{S4_FIRST + 3}", v=240_000_000),
          {"op": "print_area", "sheet": KD, "area": f"$C$3:$U${PRINT_LAST}"},
          ]
    return o


def ops_historik() -> list[dict]:
    """Steg 7: rad i flik 1:s uppdateringshistorik (kol R, dold outline-grupp). Rad 1–23 påverkas
    inte av radinfogningarna (64–70, 107)."""
    return [_fmt(FS, "R23", "R24"), _fmt(FS, "S23", "S24"),
            _set(FS, "R24", v="16. 2026-09-30 v3: Restvärde = rent marknadsvärde vid justerad direktavkastning "
                             "(flik 2 sektion 5, vikter 0/1/0 på flik 5). Bef-objekt = nollalternativ: krav 1–2 på "
                             "tillkommande flöde (flik 4 rad 64–82), krav 3 på hela fastigheten, egen avskrivningstid "
                             "för befintligt bokfört värde. Investeringsspecifikation som matris (sektion 4) styr "
                             "Budget kr/kvm. Känslighet restvärde ±1,00 pp (flik 5 rad 53–61, flik 1). Kravhyra "
                             "som output (flik 2 rad 17 / kolumn V, datatabell AC14:AI16)."),
            _set(FS, "S24", v="Claude/JW")]


def main() -> int:
    from step4_bef import ops_bef
    from step5_kanslighet import ops_kanslighet
    from step6_kravhyra import ops_kravhyra
    from fas2_fixar import ops_fas2
    from regression import PROBES, check_source_baseline, run_all
    # gamla fem rutorna rensas FÖRST — hjälpområdet AC4:AC9 återanvänds för dropdown-skalan.
    # goto sist: flik 2 öppnas på rad 1, boken öppnas på flik 1.
    ops = ([{"op": "clear", "sheet": KD, "ref": "AC1:AP20"}]
           + ops_restvarde() + ops_matris() + ops_bef() + ops_kanslighet() + ops_kravhyra() + ops_fas2()
           + ops_historik()
           + [{"op": "goto", "sheet": IRR, "ref": "A1"}, {"op": "goto", "sheet": NPV, "ref": "A1"},
              {"op": "goto", "sheet": KD, "ref": "A1"}, {"op": "goto", "sheet": FS, "ref": "A1"}])
    print("Baslinje källa …", end=" ")
    check_source_baseline(SOURCE)
    print("ok")
    print(f"Bygger {OUT.name} ({len(ops)} ops) …", end=" ")
    res = run(SOURCE, ops + PROBES, OUT)
    print("ok")
    run_all(OUT, res)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
