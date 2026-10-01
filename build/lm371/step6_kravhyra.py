"""Steg 6: kravhyra som output (beslut 4, 6, 9).

Hyresmultiplikator m (2. Kalkyldata!AD12, teknisk cell = 1) skalar hyran på alla icke-Bef-rader
(L23:L27 och marknadshyran O23:O27). En envariabels datatabell (AC14:AI16, m = 0 och m = 1) ger
NPV tillkommande och NPV vid IRR-kravet (= krav 2 som nuvärde) för basfall och yield-scenarier.
Eftersom allt är linjärt i m löses kravhyran analytiskt: m_krav = −NPV(0)/(NPV(1)−NPV(0)),
m* = MAX(krav 1, krav 2). Krav 3 binder inte (extern parameter, beslut 9).
Redovisas: nyckeltalsrutan rad 17, kolumn V i hyrestabellen, flik 5 rad 55–57 (H), flik 1.
"""
from __future__ import annotations
from step4_bef import _set, _fmt, K, KD, IRR, FS
from step5_kanslighet import S_COLS, S_OPT, S_BED, S_PES, CF_OPT, CF_PES, F_COLS, F_OPT, F_BED, F_PES, _pct

M_CELL = "$AD$12"
FS_ROW = 107                     # ny rad på flik 1 (infogas; rader ≥107 flyttas +1)


def _nl(sheet, ref, fmt):
    return {"op": "numfmt_local", "sheet": sheet, "ref": ref, "fmt": fmt}


KRKVM = '# ##0" kr/kvm"'      # NumberFormatLocal; mellanslag före citattecknet tolkas som tusentalsskalning


def ops_kravhyra() -> list[dict]:
    o: list[dict] = []
    # ── Hyresmultiplikator + datatabell i hjälpområdet ──
    o += [_set(KD, "AC12", v="Hyresmultiplikator m — teknisk cell för datatabellen, lämna 1"),
          _set(KD, "AD12", v=1),
          _set(KD, "AC13", v="m"), _set(KD, "AD13", v="NPV tillk."), _set(KD, "AE13", v="NPV vid IRR-krav"),
          _set(KD, "AF13", v="NPV tillk. opt"), _set(KD, "AG13", v="NPV tillk. pess"),
          _set(KD, "AH13", v="NPV vid krav, opt"), _set(KD, "AI13", v="NPV vid krav, pess"),
          _set(KD, "AD14", f="='4. NPV'!$E$82"),
          _set(KD, "AE14", f=f"='5. IRR'!$D$37+NPV({K}$F$14,'5. IRR'!$E$37:$BA$37)"),
          _set(KD, "AF14", f=f"='5. IRR'!$F${S_OPT}"), _set(KD, "AG14", f=f"='5. IRR'!$F${S_PES}"),
          _set(KD, "AH14", f=f"='5. IRR'!$D${CF_OPT}+NPV({K}$F$14,'5. IRR'!$E${CF_OPT}:$BA${CF_OPT})"),
          _set(KD, "AI14", f=f"='5. IRR'!$D${CF_PES}+NPV({K}$F$14,'5. IRR'!$E${CF_PES}:$BA${CF_PES})"),
          _set(KD, "AC15", v=0), _set(KD, "AC16", v=1),
          {"op": "datatable", "sheet": KD, "ref": "AC14:AI16", "col_input": "AD12"},
          # analytisk lösning per krav och scenario
          _set(KD, "AC18", v="m per krav →"),
          _set(KD, "AD18", f="=IFERROR(-AD15/(AD16-AD15),1)"), _set(KD, "AE18", f="=IFERROR(-AE15/(AE16-AE15),1)"),
          _set(KD, "AF18", f="=IFERROR(-AF15/(AF16-AF15),1)"), _set(KD, "AG18", f="=IFERROR(-AG15/(AG16-AG15),1)"),
          _set(KD, "AH18", f="=IFERROR(-AH15/(AH16-AH15),1)"), _set(KD, "AI18", f="=IFERROR(-AI15/(AI16-AI15),1)"),
          _set(KD, "AC19", v="m* = MAX(krav 1, krav 2): bedömt | bindande | opt | pess"),
          _set(KD, "AD19", f="=MAX(AD18,AE18)"), _set(KD, "AE19", f='=IF(AD18>=AE18,"NPV binder","IRR binder")'),
          _set(KD, "AF19", f="=MAX(AF18,AH18)"), _set(KD, "AG19", f="=MAX(AG18,AI18)"),
          _set(KD, "AC20", v="Angiven hyra, viktad kr/kvm (icke-Bef)"),
          _set(KD, "AD20", f='=IFERROR(SUMPRODUCT(--($E$23:$E$27<>"Bef"),$I$23:$I$27,$H$23:$H$27)'
                             '/SUMPRODUCT(--($E$23:$E$27<>"Bef"),$H$23:$H$27),0)'),
          ]
    # ── Multiplikatorn in i modellen: totalhyra och marknadshyra på icke-Bef-rader ──
    for r in range(23, 28):
        o += [_set(KD, f"L{r}", f=f'=((H{r}*I{r}*IF($E{r}="Bef",1,{M_CELL}))+(K{r}*H{r}))'),
              _set(KD, f"O{r}", f=f'=I{r}*IF($E{r}="Bef",1,{M_CELL})')]
    # ── Nyckeltalsrutan rad 17 ──
    o += [_fmt(KD, "C16", "C17"), _set(KD, "C17", v="Kravhyra 1–2, kr/kvm"),
          _fmt(KD, "D16:E16", "D17:E17"), {"op": "merge", "sheet": KD, "ref": "D17:E17"},
          _set(KD, "D17", f='=IF(G13<>"","",CEILING($AD$19*$AD$20,1))'), _nl(KD, "D17", '# ##0'),
          _fmt(KD, "F16", "F17"), _set(KD, "F17", f='=IF(G13<>"","",LEFT($AE$19,3))'),
          ]
    # ── Flik 5: kravhyra per yield-scenario ──
    o += [_fmt(IRR, f"G{S_COLS}", f"H{S_COLS}"), _set(IRR, f"H{S_COLS}", v="Kravhyra kr/kvm"),
          _fmt(IRR, f"G{S_OPT}:G{S_PES}", f"H{S_OPT}:H{S_PES}"),
          _set(IRR, f"H{S_OPT}", f=f"=CEILING({K}$AD$20*{K}$AF$19,1)"),
          _set(IRR, f"H{S_BED}", f=f"=CEILING({K}$AD$20*{K}$AD$19,1)"),
          _set(IRR, f"H{S_PES}", f=f"=CEILING({K}$AD$20*{K}$AG$19,1)"),
          _nl(IRR, f"H{S_OPT}:H{S_PES}", '# ##0'), {"op": "align", "sheet": IRR, "ref": f"H{S_OPT}:H{S_PES}", "h": -4152},
          ]
    # ── Flik 1: ny rad i Resultat & avkastning + kravhyra i scenariotabellen ──
    o += [{"op": "insert_rows", "sheet": FS, "ref": f"{FS_ROW}:{FS_ROW}"},
          {"op": "clear_fmt", "sheet": FS, "ref": f"B{FS_ROW}:H{FS_ROW}"},
          _fmt(FS, "B109:H109", f"B{FS_ROW}:H{FS_ROW}"),                 # = gamla PV-raden (vit), nu 109
          {"op": "rowheight", "sheet": FS, "row": FS_ROW, "h": 15},
          _set(FS, f"B{FS_ROW}", v="Kravhyra krav 1–2, kr/kvm (tillkommande objekt, viktad)"),
          _set(FS, f"D{FS_ROW}", f=f"={K}$AE$19"), _set(FS, f"G{FS_ROW}", f=f"={K}$D$17"),
          _nl(FS, f"G{FS_ROW}", KRKVM),
          _set(FS, f"F{F_COLS}", v="Kravhyra kr/kvm"),
          _set(FS, f"F{F_OPT}", f=f"='5. IRR'!H{S_OPT}"), _set(FS, f"F{F_BED}", f=f"='5. IRR'!H{S_BED}"),
          _set(FS, f"F{F_PES}", f=f"='5. IRR'!H{S_PES}"),
          _nl(FS, f"F{F_OPT}:F{F_PES}", KRKVM),
          ]
    return o
