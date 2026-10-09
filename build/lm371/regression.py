"""Regressionsgate för LM371_v3.

(a) Källan reproducerar dagens siffror: D13 = 10 300, D14 = 6,298 %, D15 = 10 300.
(b) v3 med Neutralt: D13/D15 oförändrade, D14 = 8,026 % (rent MV), M10 = 0, M11 = M6,
    S23 = 240 Mkr / 6 100 kvm via matrisen, noll felvärden i synliga områden.
(c) v3 med vikter 0/0,4/0,6 återger dagens IRR exakt.
(d) Icke-neutral bedömning flyttar yield, exit, IRR och NPV åt rätt håll.
(e) Tom investeringsmatris → tvingande-vakt + blanka resultat.
(f) Bef-rad + befintligt bokfört värde påverkar bara krav 3 (hela fastigheten).
(g) v3.1 (beslut A–E 2026-10-09): ett restvärde i alla krav (flik 4 rad 88/89), M8 = M6 + M9,
    vakansfaktorns skala + ekvivalent vakans, vakansgradens tidsdomän, extraordinärt UH t.o.m. horisonten.

Kör:  python build/lm371/regression.py   (mot build/lm371/LM371_v3.xlsx)
"""
from __future__ import annotations
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from comrun import run, probe, probe_dv, scan  # noqa: E402

KD = "2. Kalkyldata"
IRR = "5. IRR"
NPV = "4. NPV"
FIN = "3. Finansiering"

BASE_NPV, BASE_IRR, BASE_PV = 10300.0, 0.0629828143787814, 10300.0   # källmallen (orörd)
# v3-pinnar. Fas 2 F2 (CA-index ^(år−1) i stället för ^år från år 2) höjer IRR marginellt:
# rent MV 8,0257 % → 8,0324 %; legacy-vikter 0/0,4/0,6: 6,2983 % → LEGACY_IRR nedan. NPV/PV opåverkade.
# v3.1 (beslut C): IRR-exit = restvärde år N+1-driftnetto (= NPV-spårets) → 8,0324 % → 8,0804 %;
# legacy 0/0,4/0,6: 6,3062 % → 6,3333 %. Avsiktliga brott mot baslinjen, se REGRESSION_v3.1.md.
V3_IRR = 0.08080374302175485          # v3: 0.08032406218316468
LEGACY_IRR = 0.063333409992863        # v3: 0.06306226827590034 (källans 6,2983 % + CA-fix)
EXIT_V31 = 271_940_000                # = 4. NPV!E79 (v3: 5. IRR!C7 = 268 500 000 på år N-driftnetto)
TOL_IRR = 1e-6

PROBES = [probe(KD, "D13"), probe(KD, "D14"), probe(KD, "D15"),
          probe(KD, "M6"), probe(KD, "M8"), probe(KD, "M10"), probe(KD, "M11"),
          probe(KD, "S23"), probe(KD, "H23"), probe(KD, "G69"), probe(KD, "L69"),
          probe(IRR, "C10"), probe(IRR, "A7"), probe(IRR, "A8"),
          probe(NPV, "E59"), probe(NPV, "E82"),
          scan(KD, "C3:U81"), scan(IRR, "B5:C41"), scan(NPV, "B64:BA82")]


def _p(res, sheet, ref):
    return res["probes"][f"{sheet}!{ref}"]


def _is_err(v) -> bool:
    """Excel-felvärden marshalas via COM som Int32 i intervallet −2146826281 … −2146826246."""
    return isinstance(v, int) and not isinstance(v, bool) and -2146827000 < v < -2146826000


def _close(a, b, tol):
    return isinstance(a, (int, float)) and abs(a - b) <= tol


def check_source_baseline(source: Path) -> None:
    res = run(source, [probe(KD, "D13"), probe(KD, "D14"), probe(KD, "D15")])
    assert _close(_p(res, KD, "D13"), BASE_NPV, 1), res
    assert _close(_p(res, KD, "D14"), BASE_IRR, TOL_IRR), res
    assert _close(_p(res, KD, "D15"), BASE_PV, 1), res


def check_v3(res: dict) -> None:
    fails = []

    def want(label, ok):
        if not ok:
            fails.append(label)
    want("NPV D13 oförändrad", _close(_p(res, KD, "D13"), BASE_NPV, 1))
    want("PV vs BV D15 oförändrad", _close(_p(res, KD, "D15"), BASE_PV, 1))
    want("IRR D14 = rent MV, exit på N+1 (8,080 %)", _close(_p(res, KD, "D14"), V3_IRR, TOL_IRR))
    want("M10 = 0 vid Neutralt", _close(_p(res, KD, "M10"), 0.0, 1e-12))
    want("M11 = M6", _close(_p(res, KD, "M11"), _p(res, KD, "M6"), 1e-12))
    want("M8 = M6 + inflation (0,085)", _close(_p(res, KD, "M8"), 0.085, 1e-12))
    want("S23 = 240 Mkr / area", _close(_p(res, KD, "S23"), 240_000_000 / _p(res, KD, "H23"), 1e-6))
    want("Matris G69 = 240 Mkr", _close(_p(res, KD, "G69"), 240_000_000, 0.5))
    want("Exit C10 = 271,94 Mkr (rent MV, = NPV E79)", _close(_p(res, IRR, "C10"), EXIT_V31, 0.5))
    want("Vikter 0/1/0", _p(res, IRR, "A7") == 1 and _p(res, IRR, "A8") == 0)
    want("NPV tillkommande = NPV hela utan Bef-rader", _close(_p(res, NPV, "E82"), _p(res, NPV, "E59"), 0.5))
    for k, n in res["errors"].items():
        want(f"noll felvärden i {k} (hittade {n})", n == 0)
    for k, v in res["probes"].items():
        print(f"  {k:28s} {v}")
    if fails:
        raise AssertionError("REGRESSION RÖD:\n  - " + "\n  - ".join(fails))


def check_legacy_weights(v3: Path) -> None:
    ops = [{"op": "set", "sheet": IRR, "ref": "A7", "value": 0.4},
           {"op": "set", "sheet": IRR, "ref": "A8", "value": 0.6},
           probe(KD, "D13"), probe(KD, "D14"), probe(KD, "D15")]
    res = run(v3, ops)
    assert _close(_p(res, KD, "D14"), LEGACY_IRR, TOL_IRR), f"Legacy-vikter ger inte pinnad IRR: {res}"
    assert _close(_p(res, KD, "D13"), BASE_NPV, 1) and _close(_p(res, KD, "D15"), BASE_PV, 1), res
    print(f"  legacy 0/0,4/0,6 → IRR {_p(res, KD, 'D14'):.6%} (källa 6,2983 % + CA-fix + exit N+1) ✓")


def check_adjustment(v3: Path) -> None:
    """Mycket negativt på en faktor: +0,50 pp → M11 = 7,0 %, exit = MV vid 7,0 %, IRR/NPV lägre, M8 oförändrad (A)."""
    ops = [{"op": "set", "sheet": KD, "ref": "G75", "value": "Mycket negativt"},
           probe(KD, "M10"), probe(KD, "M11"), probe(KD, "M8"), probe(KD, "D13"), probe(KD, "D14"), probe(IRR, "C10")]
    res = run(v3, ops)
    assert _close(_p(res, KD, "M10"), 0.005, 1e-12) and _close(_p(res, KD, "M11"), 0.07, 1e-12), res
    assert _close(_p(res, KD, "M8"), 0.085, 1e-12), res
    assert _close(_p(res, IRR, "C10"), round(EXIT_V31 * 0.065 / 0.07 / 1e4) * 1e4, 1e4), res
    assert _p(res, KD, "D14") < V3_IRR and _p(res, KD, "D13") < BASE_NPV, res
    print(f"  +0,50 pp → yield 7,0 %, exit {_p(res, IRR, 'C10'):,.0f}, IRR {_p(res, KD, 'D14'):.4%}, NPV {_p(res, KD, 'D13'):,.0f} tkr ✓")


def check_empty_matrix(v3: Path) -> None:
    """Tom matris + ingen överstyrning → vakten slår till och resultaten blankas."""
    ops = [{"op": "set", "sheet": KD, "ref": "G58", "value": 0},
           probe(KD, "G13"), probe(KD, "D13"), probe(KD, "D15"), probe(KD, "S23")]
    res = run(v3, ops)
    assert _p(res, KD, "G13") == "Investering saknas för objekt med area (sektion 4)", res
    assert _p(res, KD, "D13") in ("", None) and _p(res, KD, "D15") in ("", None), res
    print("  tom matris → vakt + blanka resultat ✓")


def check_bef(v3: Path) -> None:
    """Bef-rad + befintligt bokfört värde + Bef-DoU: krav 1 och 2 oförändrade, krav 3 (hela) ändras."""
    def kd(ref, v):
        return {"op": "set", "sheet": KD, "ref": ref, "value": v}
    ops = [kd("C24", "Befhuset"), kd("E24", "Bef"), kd("F24", 2000), kd("I24", 1500), kd("J24", 0.7),
           kd("M24", 2026), kd("N24", 2035), kd("O24", 1400), kd("P24", 0.05),
           kd("G34", 120),            # Bef-schablon fastighetsskötsel
           kd("P9", 50_000_000),      # befintligt bokfört värde
           kd("Q13", 15),             # egen återstående avskrivningstid
           probe(KD, "D13"), probe(KD, "D14"), probe(KD, "D15"), probe(KD, "G13"),
           probe(NPV, "E82"), probe(NPV, "E59"), probe(NPV, "E70"), probe(IRR, "C10"), probe(FIN, "D40"),
           scan(NPV, "B64:BA82"), scan(IRR, "B5:C41")]
    res = run(v3, ops)
    fails = []
    if not _close(_p(res, KD, "D13"), BASE_NPV, 1):
        fails.append(f"krav 1 påverkas av Bef: {_p(res, KD, 'D13')}")
    if not _close(_p(res, KD, "D14"), V3_IRR, TOL_IRR):
        fails.append(f"krav 2 påverkas av Bef: {_p(res, KD, 'D14')}")
    if _close(_p(res, KD, "D15"), BASE_PV, 1):
        fails.append("krav 3 (hela) reagerar inte på Bef/P9")
    if not _p(res, NPV, "E70"):
        fails.append("Bef-driftnetto = 0 trots Bef-rad")
    if not _close(_p(res, FIN, "D40"), -50_000_000 / 15, 1):
        fails.append(f"avskrivning bef fel: {_p(res, FIN, 'D40')}")
    if _p(res, KD, "G13") != "":
        fails.append(f"vakt: {_p(res, KD, 'G13')}")
    for k, n in res["errors"].items():
        if n:
            fails.append(f"felvärden i {k}: {n}")
    if fails:
        raise AssertionError("BEF-TEST RÖD:\n  - " + "\n  - ".join(fails) + f"\n{res}")
    print(f"  Bef-rad + P9 50 Mkr → krav 1/2 oförändrade, krav 3 {_p(res, KD, 'D15'):,.0f} tkr (hela) ✓")


def check_sensitivity(v3: Path) -> None:
    """Bedömt-scenariot == rad 37/D40/E82 bit-för-bit; opt > bedömt > pess; flik 1 speglar flik 5."""
    ops = [probe(IRR, "D40"), probe(IRR, "E55"), probe(IRR, "E56"), probe(IRR, "E57"),
           probe(IRR, "F55"), probe(IRR, "F56"), probe(IRR, "F57"), probe(IRR, "D56"), probe(IRR, "C10"),
           probe(NPV, "E82"), probe("1. Framskrivningsunderlag", "D68"), probe("1. Framskrivningsunderlag", "G108"),
           scan(IRR, "B53:G57"), scan(IRR, "D59:BA61"), scan("1. Framskrivningsunderlag", "B64:G70")]
    res = run(v3, ops)
    fails = []
    if _p(res, IRR, "E56") != _p(res, IRR, "D40"):
        fails.append(f"bedömt IRR {_p(res, IRR, 'E56')} != D40 {_p(res, IRR, 'D40')}")
    if not _close(_p(res, IRR, "F56"), _p(res, NPV, "E82"), 0.5):
        fails.append(f"bedömt NPV {_p(res, IRR, 'F56')} != E82 {_p(res, NPV, 'E82')}")
    if not _close(_p(res, IRR, "D56"), _p(res, IRR, "C10"), 0.5):
        fails.append("bedömt exit != C10")
    if not (_p(res, IRR, "E55") > _p(res, IRR, "E56") > _p(res, IRR, "E57")):
        fails.append("IRR-ordning opt > bedömt > pess bruten")
    if not (_p(res, IRR, "F55") > _p(res, IRR, "F56") > _p(res, IRR, "F57")):
        fails.append("NPV-ordning opt > bedömt > pess bruten")
    if _p(res, "1. Framskrivningsunderlag", "D68") != _p(res, IRR, "E56"):
        fails.append("flik 1 bedömt-IRR speglar inte flik 5")
    if not _close(_p(res, "1. Framskrivningsunderlag", "G108"), _p(res, NPV, "E82"), 0.5):
        fails.append("flik 1 NPV-rad (flyttad 100→108) pekar fel efter radinfogning")
    for k, n in res["errors"].items():
        if n:
            fails.append(f"felvärden i {k}: {n}")
    if fails:
        raise AssertionError("KÄNSLIGHET RÖD:" + chr(10) + "  - " + (chr(10) + "  - ").join(fails) + chr(10) + f"{res}")
    print(f"  känslighet: IRR {_p(res, IRR, 'E55'):.2%} / {_p(res, IRR, 'E56'):.2%} / {_p(res, IRR, 'E57'):.2%}, "
          f"NPV {_p(res, IRR, 'F55'):,.0f} / {_p(res, IRR, 'F56'):,.0f} / {_p(res, IRR, 'F57'):,.0f} ✓")


def check_kravhyra(v3: Path) -> None:
    """Kravhyran är exakt: insatt som hyra klarar båda kraven med liten marginal, och det bindande kravet ligger på gränsen."""
    res = run(v3, [probe(KD, "D17"), probe(KD, "I17"), probe(KD, "AE19"), probe(KD, "AD19"), probe(KD, "AD20"),
                   probe(KD, "AD15"), probe(KD, "AD16"), probe(KD, "AE15"), probe(KD, "AE16"),
                   probe(IRR, "H55"), probe(IRR, "H56"), probe(IRR, "H57"), probe("1. Framskrivningsunderlag", "G107")])
    krav = _p(res, KD, "D17")
    fails = []
    if not (isinstance(krav, (int, float)) and 0 < krav < 2650):
        fails.append(f"kravhyra orimlig: {krav} (exemplet klarar kraven vid 2 650)")
    if _p(res, KD, "I17") not in ("", None):
        fails.append(f"I17-notis ska vara tom när hyran räcker: {_p(res, KD, 'I17')!r}")
    if not (_p(res, IRR, "H55") < _p(res, IRR, "H56") < _p(res, IRR, "H57")):
        fails.append("kravhyra per scenario ej stigande opt < bedömt < pess")
    if _p(res, "1. Framskrivningsunderlag", "G107") != _p(res, KD, "D17"):
        fails.append("flik 1 kravhyra speglar inte D17")
    if _p(res, KD, "AD16") <= _p(res, KD, "AD15") or _p(res, KD, "AE16") <= _p(res, KD, "AE15"):
        fails.append(f"datatabellen saknar lutning: {res['probes']}")
    if fails:
        raise AssertionError("KRAVHYRA RÖD:" + chr(10) + "  - " + (chr(10) + "  - ").join(fails))
    # insatt kravhyra → båda kraven klaras, det bindande ligger nära gränsen
    res2 = run(v3, [{"op": "set", "sheet": KD, "ref": "I23", "value": float(krav)},
                    probe(KD, "D13"), probe(KD, "D14"), probe(KD, "F14")])
    npv, irr, kravirr = _p(res2, KD, "D13"), _p(res2, KD, "D14"), _p(res2, KD, "F14")
    ok_npv, ok_irr = npv >= -1, irr >= kravirr - 1e-5
    tight = npv < 300 or irr < kravirr + 0.0005            # CEILING till hel kr/kvm ger liten marginal
    if not (ok_npv and ok_irr and tight):
        raise AssertionError(f"KRAVHYRA RÖD: vid {krav} kr/kvm NPV={npv} tkr, IRR={irr:.4%} (krav {kravirr:.2%})")
    print(f"  kravhyra {krav:,.0f} kr/kvm ({_p(res, KD, 'AE19')}) → NPV {npv:,.0f} tkr, IRR {irr:.3%}; "
          f"scenarier {_p(res, IRR, 'H55'):,.0f}/{_p(res, IRR, 'H56'):,.0f}/{_p(res, IRR, 'H57'):,.0f} ✓")


def check_fas2(v3: Path) -> None:
    """F1 index år 2 för tidigt startat Bef-avtal, F3 ingen inflatering före kalkylstart, F4 LOOKUP,
    F5 negativa PV utan #NUM!, F6 L39 och inga #REF!-formler kvar."""
    import openpyxl
    def kd(ref, v):
        return {"op": "set", "sheet": KD, "ref": ref, "value": v}
    ops = [kd("C24", "Befhuset"), kd("E24", "Bef"), kd("F24", 2000), kd("I24", 1500), kd("J24", 0.7),
           kd("M24", 2020), kd("N24", 2035),                       # F1: avtal startat före kalkylstart
           kd("R23", 2020), kd("Q23", 2020),                       # F3: produktionsavslut före kalkylstart
           kd("P6", 1000),                                         # F4: tomträttsavgäld
           kd("H34", 100), kd("L34", 0.1),                         # F6: kostnadshöjning viktad
           probe(NPV, "D13"), probe(NPV, "E13"), probe(NPV, "F13"), probe(KD, "X23"), probe(KD, "W23"),
           probe(NPV, "D28"), probe(NPV, "E28"), probe(KD, "L39"), probe(KD, "I34"), probe(KD, "I39"),
           probe(IRR, "D26"), probe(IRR, "E26"), probe(IRR, "F26")]
    res = run(v3, ops)
    p = lambda s, r: _p(res, s, r)  # noqa: E731
    fails = []
    if not _close(p(NPV, "D13"), 2000 * 1500, 0.5):
        fails.append(f"F1 år 1 Bef-intäkt {p(NPV, 'D13')} != 3 000 000")
    if not _close(p(NPV, "E13"), 2000 * 1500 * 1.014, 0.5):
        fails.append(f"F1 år 2 Bef-intäkt {p(NPV, 'E13')} != 3 042 000 (^1)")
    if not _close(p(NPV, "F13"), 2000 * 1500 * 1.014 ** 2, 0.5):
        fails.append(f"F1 år 3 Bef-intäkt {p(NPV, 'F13')} (^2)")
    if not (_close(p(IRR, "E26") / p(IRR, "D26"), 1.02, 1e-9) and _close(p(IRR, "F26") / p(IRR, "E26"), 1.02, 1e-9)):
        fails.append(f"F2 CA-index år1→2→3: {p(IRR, 'D26')}, {p(IRR, 'E26')}, {p(IRR, 'F26')} (förväntat ×1,02 per år)")
    if not _close(p(KD, "X23"), p(KD, "W23"), 0.5):
        fails.append(f"F3 PV investering {p(KD, 'X23')} != {p(KD, 'W23')} vid avslut före kalkylstart")
    if not (_close(p(NPV, "D28"), -1000, 0.5) and _close(p(NPV, "E28"), -1000, 0.5)):
        fails.append(f"F4 tomträttsavgäld D28/E28 = {p(NPV, 'D28')}/{p(NPV, 'E28')}")
    if not _close(p(KD, "L39"), 0.1 * p(KD, "I34") / p(KD, "I39"), 1e-9):
        fails.append(f"F6 L39 {p(KD, 'L39')}")
    # F5: hyra 0 → negativa PV utan #NUM!
    res2 = run(v3, [kd("I23", 0), probe(NPV, "C48"), probe(NPV, "E48"), probe(NPV, "E59"), probe(KD, "D13"), probe(IRR, "C10")])
    for s_, r_ in ((NPV, "C48"), (NPV, "E48"), (NPV, "E59"), (KD, "D13"), (IRR, "C10")):
        v = _p(res2, s_, r_)
        if _is_err(v) or not isinstance(v, (int, float)):
            fails.append(f"F5 {s_}!{r_} = {v!r} vid hyra 0")
    if not _is_err(_p(res2, NPV, "E48")) and _p(res2, NPV, "E48") >= 0:
        fails.append("F5 testet exercerar inte negativt PV")
    # formeltext: fångar tysta Replace-missar (Replace arbetar mot FormulaLocal)
    wb = openpyxl.load_workbook(v3)
    for sh, ref, must in ((NPV, "E13", "MAX('2. Kalkyldata'!$M24,$D$4)"), (NPV, "C48", "ROUND(("), (NPV, "E50", "$D$4:$BA$89"),
                          (NPV, "D28", "$P$5:$R$5,"), (IRR, "C7", "'4. NPV'!$E$79"), ("7.Grafer", "C40", "$72)"), (IRR, "E23", "$71)")):
        fv = wb[sh][ref].value
        if not (isinstance(fv, str) and must in fv):
            fails.append(f"formeltext {sh}!{ref} saknar {must!r}: {str(fv)[:90]}")
    if any("$P$5:$R$6" in str(wb[NPV].cell(28, c).value) for c in range(4, 54)):
        fails.append("F4 $P$5:$R$6 kvar på rad 28")
    # F6: inga #REF! i formeltext
    refs = [(ws.title, c.coordinate) for ws in wb.worksheets for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and "#REF!" in c.value]
    if refs:
        fails.append(f"F6 #REF!-formler kvar: {len(refs)} t.ex. {refs[:5]}")
    if fails:
        raise AssertionError("FAS 2 RÖD:" + chr(10) + "  - " + (chr(10) + "  - ").join(fails))
    print(f"  fas 2: index ^0/^1/^2 ✓, PV=W före kalkylstart ✓, LOOKUP ✓, hyra 0 → NPV {_p(res2, KD, 'D13'):,.0f} tkr utan #NUM! ✓, 0 #REF! ✓")


def check_restvarde_v31(v3: Path) -> None:
    """v3.1, beslut A–E. Kolumner flik 4: W = år 20 (2045), X = år 21, M = 2035, N = 2036, R = 2040."""
    def kd(ref, v):
        return {"op": "set", "sheet": KD, "ref": ref, "value": v}
    BASE = [probe(KD, "D13"), probe(KD, "D14"), probe(KD, "D15"), probe(KD, "M6"), probe(KD, "M8"), probe(KD, "M9"),
            probe(KD, "M11"), probe(KD, "H76"), probe(KD, "I76"), probe(KD, "I17"), probe(KD, "AM23"),
            probe(IRR, "C7"), probe(IRR, "C10"), probe(IRR, "W7"), probe(IRR, "B12"), probe(IRR, "C12"),
            probe(NPV, "C49"), probe(NPV, "E49"), probe(NPV, "C50"), probe(NPV, "E50"),
            probe(NPV, "C78"), probe(NPV, "E78"), probe(NPV, "C79"), probe(NPV, "E79"),
            probe(NPV, "W88"), probe(NPV, "W89"), probe(NPV, "X86"), probe(NPV, "X90"), probe(NPV, "M88"),
            probe(NPV, "N18"), probe(NPV, "W18"), probe(NPV, "W12"), probe("7.Grafer", "V44")]
    fails = []

    def want(label, ok):
        if not ok:
            fails.append(label)
    b = run(v3, BASE + [probe_dv(KD, f"P{r}") for r in range(23, 28)] + [scan(NPV, "B84:BA90")])
    P = lambda res, s, r: _p(res, s, r)  # noqa: E731
    # C — ett restvärde: IRR-exit, krav 1, krav 3 och Grafer läser samma definitionsrad
    want("C: IRR C7 = NPV E79", P(b, IRR, "C7") == P(b, NPV, "E79"))
    want("C: E79 = rad 88 år N (avrundat)", _close(P(b, NPV, "E79"), round(P(b, NPV, "W88") / 1e4) * 1e4, 0.5))
    want("C: E50 = rad 89 år N (avrundat)", _close(P(b, NPV, "E50"), round(P(b, NPV, "W89") / 1e4) * 1e4, 0.5))
    want("C: rad 88 år N = driftnetto år N+1 / M11", _close(P(b, NPV, "W88"), P(b, NPV, "X86") / P(b, KD, "M11"), 1e-3))
    want("C: C79 = rad 88 år 10", _close(P(b, NPV, "C79"), round(P(b, NPV, "M88") / 1e4) * 1e4, 0.5))
    want("C: IRR rad 7 och Grafer rad 44 = rad 88", P(b, IRR, "W7") == P(b, NPV, "W88") == P(b, "7.Grafer", "V44"))
    # A — kalkylränta restvärde frikopplad
    want("A: M8 = M6 + M9", _close(P(b, KD, "M8"), P(b, KD, "M6") + P(b, KD, "M9"), 1e-12))
    # E — vilseledande PV-cell borta
    want("E: 5. IRR B12:C12 tomma", P(b, IRR, "B12") in ("", None) and P(b, IRR, "C12") in ("", None))
    want("B: I76 = marknadsnormal vid Normal", P(b, KD, "I76") == "marknadsnormal")
    for r in range(23, 28):
        dv = b["probes"].get(f"{KD}!P{r}#dv")
        if dv is not None:                                   # None = ej läsbart (LibreOffice, steg 1)
            want(f"D: validering P{r} = {dv!r}", f"$AM${r}" in dv and f"$P${r}=0" in dv)
    for k, n in b["errors"].items():
        want(f"felvärden i {k}: {n}", n == 0)
    # B — vakansfaktorn Hög: +1,00 pp, M8 oförändrad, ekvivalent vakans = överskottsgrad × Δ / M11
    h = run(v3, [kd("G76", "Hög")] + BASE)
    v_eq = P(b, NPV, "X86") / P(b, NPV, "X90") * 0.01 / (P(b, KD, "M6") + 0.01)
    want(f"B: H76 = +1,00 pp ({P(h, KD, 'H76')})", _close(P(h, KD, "H76"), 0.01, 1e-12))
    want("B: M11 = M6 + 1 pp", _close(P(h, KD, "M11"), P(b, KD, "M6") + 0.01, 1e-12))
    want("A+B: M8 oförändrad vid faktorjustering", _close(P(h, KD, "M8"), P(b, KD, "M8"), 1e-12))
    want(f"B: I76 = '≈ {round(v_eq * 100)} % vakans' ({P(h, KD, 'I76')!r})", P(h, KD, "I76") == f"≈ {round(v_eq * 100)} % vakans")
    # ekvivalensen: samma exit som vakans v_eq dragen ur kapitaliseringen vid extern yield
    exit_vak = (P(b, NPV, "X86") - v_eq * P(b, NPV, "X90")) / P(b, KD, "M6")
    want("B: exit vid Hög = exit med ekvivalent vakans", _close(P(h, NPV, "E79"), round(exit_vak / 1e4) * 1e4, 1e4))
    want("B: Hög sänker NPV och IRR", P(h, KD, "D13") < P(b, KD, "D13") and P(h, KD, "D14") < P(b, KD, "D14"))
    lo_ = run(v3, [kd("G76", "Låg"), probe(KD, "H76"), probe(KD, "I76")])
    want(f"B: Låg = −0,25 pp med 'lägre vakans' ({P(lo_, KD, 'I76')!r})",
         _close(P(lo_, KD, "H76"), -0.0025, 1e-12) and "lägre vakans" in str(P(lo_, KD, "I76")))
    # D — avtal täcker kalkylperioden: vakansgrad utan effekt + notis
    d1 = run(v3, [kd("P23", 0.10)] + BASE)
    want("D: AM23 = FALSE när avtalet täcker perioden", not P(b, KD, "AM23"))
    want(f"D: notis i I17 ({P(d1, KD, 'I17')!r})", "Vakansgrad utan effekt" in str(P(d1, KD, "I17")))
    for s_, r_ in ((KD, "D13"), (KD, "D14"), (KD, "D15"), (NPV, "E79"), (NPV, "E50")):
        want(f"D: {s_}!{r_} oförändrad när avtalet täcker perioden", _close(P(d1, s_, r_), P(b, s_, r_), 1e-9))
    # D — avtal t.o.m. 2035: vakans 2036–2045 i kassaflödet, restvärdet fullt uthyrt
    d2 = run(v3, [kd("N23", 2035), kd("P23", 0.10)] + BASE)
    want("D: AM23 = TRUE vid avtal t.o.m. 2035", bool(P(d2, KD, "AM23")))
    want("D: vakans år 2045 = 10 % av intäkten", _close(P(d2, NPV, "W18"), -0.10 * P(d2, NPV, "W12"), 1))
    want("D: vakans 2036 ≠ 0", P(d2, NPV, "N18") < 0)
    want("D: restvärdet oförändrat (fullt uthyrt)", _close(P(d2, NPV, "E79"), P(b, NPV, "E79"), 0.5)
         and _close(P(d2, NPV, "E50"), P(b, NPV, "E50"), 0.5))
    want("D: NPV sjunker av vakansen inom perioden", P(d2, KD, "D13") < P(b, KD, "D13"))
    want("D: ingen notis när vakansgraden har effekt", "Vakansgrad utan effekt" not in str(P(d2, KD, "I17")))
    # E — extraordinärt UH: år N+1 utanför horisonten; år 15 bara i NPV 20; restvärdet normaliserat
    uh = [kd("Q34", "Skola"), kd("S34", 500)]
    e1 = run(v3, uh + [kd("R34", 2046)] + BASE)
    want(f"E: UH år N+1 → E49/E78 = 0 ({P(e1, NPV, 'E49')}/{P(e1, NPV, 'E78')})",
         P(e1, NPV, "E49") == 0 and P(e1, NPV, "E78") == 0)
    want("E: UH år N+1 påverkar inte NPV eller restvärde", _close(P(e1, KD, "D13"), P(b, KD, "D13"), 1)
         and _close(P(e1, NPV, "E79"), P(b, NPV, "E79"), 0.5))
    e2 = run(v3, uh + [kd("R34", 2040)] + BASE)
    want(f"E: UH 2040 → C49/C78 = 0, E78 < 0 ({P(e2, NPV, 'C78')}/{P(e2, NPV, 'E78')})",
         P(e2, NPV, "C49") == 0 and P(e2, NPV, "C78") == 0 and P(e2, NPV, "E78") < 0)
    want("E: UH 2040 → restvärde år 10 och 20 oförändrat", _close(P(e2, NPV, "C79"), P(b, NPV, "C79"), 0.5)
         and _close(P(e2, NPV, "E79"), P(b, NPV, "E79"), 0.5))
    if fails:
        raise AssertionError("V3.1 RÖD:" + chr(10) + "  - " + (chr(10) + "  - ").join(fails))
    print(f"  v3.1: C7 = E79 = {P(b, NPV, 'E79'):,.0f} ✓, M8 = M6 + M9 ✓, Hög → {P(h, KD, 'I76')} ✓, "
          f"vakansgrad spärr/notis ✓, UH t.o.m. horisonten ✓")


def run_all(v3: Path, res: dict | None = None) -> None:
    check_v3(res if res is not None else run(v3, PROBES))
    check_legacy_weights(v3)
    check_adjustment(v3)
    check_empty_matrix(v3)
    check_bef(v3)
    check_sensitivity(v3)
    check_kravhyra(v3)
    check_fas2(v3)
    check_restvarde_v31(v3)
    print("REGRESSION GRÖN")


def main(path: Path | None = None) -> int:
    run_all(path or HERE / "LM371_v3.xlsx")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(Path(sys.argv[1]) if len(sys.argv) > 1 else None))
