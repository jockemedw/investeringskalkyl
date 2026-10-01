"""Regressionsgate för LM371_v3.

(a) Källan reproducerar dagens siffror: D13 = 10 300, D14 = 6,298 %, D15 = 10 300.
(b) v3 med Neutralt: D13/D15 oförändrade, D14 = 8,026 % (rent MV), M10 = 0, M11 = M6,
    S23 = 240 Mkr / 6 100 kvm via matrisen, noll felvärden i synliga områden.
(c) v3 med vikter 0/0,4/0,6 återger dagens IRR exakt.
(d) Icke-neutral bedömning flyttar yield, exit, IRR och NPV åt rätt håll.
(e) Tom investeringsmatris → tvingande-vakt + blanka resultat.
(f) Bef-rad + befintligt bokfört värde påverkar bara krav 3 (hela fastigheten).

Kör:  python build/lm371/regression.py   (mot build/lm371/LM371_v3.xlsx)
"""
from __future__ import annotations
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from comrun import run, probe, scan  # noqa: E402

KD = "2. Kalkyldata"
IRR = "5. IRR"
NPV = "4. NPV"
FIN = "3. Finansiering"

BASE_NPV, BASE_IRR, BASE_PV = 10300.0, 0.0629828143787814, 10300.0   # källmallen (orörd)
# v3-pinnar. Fas 2 F2 (CA-index ^(år−1) i stället för ^år från år 2) höjer IRR marginellt:
# rent MV 8,0257 % → 8,0324 %; legacy-vikter 0/0,4/0,6: 6,2983 % → LEGACY_IRR nedan. NPV/PV opåverkade.
V3_IRR = 0.08032406218316468
LEGACY_IRR = 0.06306226827590034      # fas 2 F2: källans 6,2983 % + CA-fix (COM 2026-09-30)
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
    want("IRR D14 = rent MV 8,026 %", _close(_p(res, KD, "D14"), V3_IRR, TOL_IRR))
    want("M10 = 0 vid Neutralt", _close(_p(res, KD, "M10"), 0.0, 1e-12))
    want("M11 = M6", _close(_p(res, KD, "M11"), _p(res, KD, "M6"), 1e-12))
    want("M8 = M11 + inflation (0,085)", _close(_p(res, KD, "M8"), 0.085, 1e-12))
    want("S23 = 240 Mkr / area", _close(_p(res, KD, "S23"), 240_000_000 / _p(res, KD, "H23"), 1e-6))
    want("Matris G69 = 240 Mkr", _close(_p(res, KD, "G69"), 240_000_000, 0.5))
    want("Exit C10 = 268,5 Mkr (rent MV)", _close(_p(res, IRR, "C10"), 268_500_000, 0.5))
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
    print(f"  legacy 0/0,4/0,6 → IRR {_p(res, KD, 'D14'):.6%} (källa 6,2983 % + CA-fix) ✓")


def check_adjustment(v3: Path) -> None:
    """Mycket negativt på en faktor: +0,50 pp → M11 = 7,0 %, exit = MV vid 7,0 %, IRR/NPV lägre."""
    ops = [{"op": "set", "sheet": KD, "ref": "G75", "value": "Mycket negativt"},
           probe(KD, "M10"), probe(KD, "M11"), probe(KD, "M8"), probe(KD, "D13"), probe(KD, "D14"), probe(IRR, "C10")]
    res = run(v3, ops)
    assert _close(_p(res, KD, "M10"), 0.005, 1e-12) and _close(_p(res, KD, "M11"), 0.07, 1e-12), res
    assert _close(_p(res, KD, "M8"), 0.09, 1e-12), res
    assert _close(_p(res, IRR, "C10"), round(268_500_000 * 0.065 / 0.07 / 1e5) * 1e5, 0.5), res
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
    for sh, ref, must in ((NPV, "E13", "MAX('2. Kalkyldata'!$M24,$D$4)"), (NPV, "C48", "ROUND(("), (NPV, "E50", "$M$11"),
                          (NPV, "D28", "$P$5:$R$5,"), (IRR, "C7", "ROUND(("), ("7.Grafer", "C40", "$72)"), (IRR, "E23", "$71)")):
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


def run_all(v3: Path, res: dict | None = None) -> None:
    check_v3(res if res is not None else run(v3, PROBES))
    check_legacy_weights(v3)
    check_adjustment(v3)
    check_empty_matrix(v3)
    check_bef(v3)
    check_sensitivity(v3)
    check_kravhyra(v3)
    check_fas2(v3)
    print("REGRESSION GRÖN")


def main(path: Path | None = None) -> int:
    run_all(path or HERE / "LM371_v3.xlsx")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(Path(sys.argv[1]) if len(sys.argv) > 1 else None))
