"""Regressionsgate för LM371_v3.

(a) Källan reproducerar dagens siffror: D13 = 10 300, D14 = 6,298 %, D15 = 10 300.
(b) v3 med Neutralt: D13/D15 oförändrade, D14 = 8,026 % (rent MV), M10 = 0, M11 = M6,
    S23 = 240 Mkr / 6 100 kvm via matrisen, noll felvärden i flik 2:s synliga område.
(c) v3 med vikter 0/0,4/0,6 återger dagens IRR exakt.

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

BASE_NPV, BASE_IRR, BASE_PV = 10300.0, 0.0629828143787814, 10300.0
V3_IRR = 0.0802573957499011          # rent marknadsvärde, exemplets hyra (COM-verifierat 2026-09-30)
TOL_IRR = 1e-6

PROBES = [probe(KD, "D13"), probe(KD, "D14"), probe(KD, "D15"),
          probe(KD, "M6"), probe(KD, "M8"), probe(KD, "M10"), probe(KD, "M11"),
          probe(KD, "S23"), probe(KD, "H23"), probe(KD, "G69"), probe(KD, "L69"),
          probe(IRR, "C10"), probe(IRR, "A7"), probe(IRR, "A8"),
          scan(KD, "C3:U81"), scan(IRR, "B5:C41")]


def _p(res, sheet, ref):
    return res["probes"][f"{sheet}!{ref}"]


def _close(a, b, tol):
    return a is not None and abs(a - b) <= tol


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
    assert _close(_p(res, KD, "D14"), BASE_IRR, TOL_IRR), f"Legacy-vikter ger inte dagens IRR: {res}"
    assert _close(_p(res, KD, "D13"), BASE_NPV, 1) and _close(_p(res, KD, "D15"), BASE_PV, 1), res
    print(f"  legacy 0/0,4/0,6 → IRR {_p(res, KD, 'D14'):.6%} = dagens ✓")


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


def main(path: Path | None = None) -> int:
    v3 = path or HERE / "LM371_v3.xlsx"
    res = run(v3, PROBES)
    check_v3(res)
    check_legacy_weights(v3)
    check_adjustment(v3)
    check_empty_matrix(v3)
    print("REGRESSION GRÖN")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(Path(sys.argv[1]) if len(sys.argv) > 1 else None))
