"""LM 371 utan Excel: applicera ops direkt i XML på en byggd version och fyll cachade värden.

    python build/lm371/build_xml.py <bas.xlsx> <ut.xlsx> [modul:funktion ...]

    v3 → v3.1:  build_xml.py v3.xlsx ut.xlsx step7_restvarde:ops_restvarde_v31 step7_restvarde:ops_historik_v31
    v3.1 → v3.2: build_xml.py v31.xlsx ut.xlsx step8_design:ops_design_v32 step8_design:ops_historik_v32
    v3.2 → v3.3: build_xml.py v32.xlsx ut.xlsx step9_hjalptexter:ops_hjalptexter_v33 step9_hjalptexter:ops_historik_v33

1. Ops från angivna stegfunktioner (samma som build.py kör via Excel COM) appliceras med apply_xml.py.
2. Filen räknas om i LibreOffice (lo_gate.LO, datatabellen emulerad) och resultaten skrivs in som cachade värden
   i alla formelceller — filen visar rätt siffror även i läsare som inte räknar om. fullCalcOnLoad gör att
   Excel ändå räknar om allt vid öppning.
3. Strukturkontroll + regressionsgaten (regression.py via lo_gate) körs mot resultatet.

build.py (Excel COM, från orörd mall) är fortsatt den kanoniska helbyggnaden; denna väg är inkrementell.
"""
from __future__ import annotations
import sys
import tempfile
import types
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from apply_xml import Book, addr, cells_in, q, validate  # noqa: E402
import lo_gate  # noqa: E402

LO_ERR = {502: "#NUM!", 503: "#NUM!", 504: "#VALUE!", 519: "#VALUE!", 522: "#REF!", 523: "#NUM!",
          524: "#REF!", 525: "#NAME?", 532: "#DIV/0!", 32767: "#N/A"}


def formula_cells(path: Path) -> dict[str, dict[str, str]]:
    """{blad: {ref: 'f'|'dt'|'b'}} — formelceller, datatabellens resultatceller och booleska formler."""
    book = Book(path)
    out = {}
    for name in book.sheet_names:
        cells = {}
        for row in book.sheet(name).rows.values():
            for c in row.findall(q("c")):
                f = c.find(q("f"))
                if f is None:
                    continue
                if f.get("t") == "dataTable":
                    for r_, c_ in cells_in(f.get("ref")):
                        cells[addr(r_, c_)] = "dt"
                elif c.get("t") == "b" or (f.text or "").startswith(("OR(", "AND(")):
                    cells[c.get("r")] = "b"
                else:
                    cells[c.get("r")] = "f"
        out[name] = cells
    return out


def lo_values(path: Path, wanted: dict[str, dict[str, str]]) -> dict[str, dict[str, object]]:
    """Räkna om i LibreOffice (datatabellen emulerad) och läs resultaten för de efterfrågade cellerna."""
    lo = lo_gate._lo()
    doc = lo.open(path)
    lo_gate.emulate_datatable(lo, doc)
    out = {}
    for name, cells in wanted.items():
        vals = {}
        for ref, kind in cells.items():
            c = lo.cell(doc, name, ref)
            typ = c.Type.value
            err = c.getError() if typ == "FORMULA" else 0
            if err:
                vals[ref] = LO_ERR.get(err, "#VALUE!")
            elif typ == "TEXT" or (typ == "FORMULA" and c.FormulaResultType2 == 2):
                vals[ref] = c.String
            elif typ == "EMPTY":
                vals[ref] = None
            elif kind == "b":
                vals[ref] = bool(c.Value)
            else:
                vals[ref] = c.Value
        out[name] = vals
    doc.close(True)
    return out


def load_ops(specs: list[str]) -> list[dict]:
    import importlib
    ops = []
    for spec in specs:
        mod, _, fn = spec.partition(":")
        ops += getattr(importlib.import_module(mod), fn)()
    return ops


def build(base: Path, out: Path, specs: list[str]) -> None:
    if any(s.startswith("step7_restvarde") for s in specs):
        b84 = Book(base).sheet("4. NPV").find(84, 2)
        if b84 is not None and b84.find(q("is")) is not None:
            raise SystemExit(f"{base.name} har redan v3.1-blocket (4. NPV!B84) — utgå från v3-filen.")
    ops = load_ops(specs)
    tmp = Path(tempfile.mkdtemp()) / "v31.xlsx"
    book = Book(base)
    book.apply(ops)
    book.save(tmp)
    print(f"ops applicerade: {len(ops)}")
    problems = validate(tmp)
    if problems:
        raise SystemExit("STRUKTURFEL:\n  " + "\n  ".join(problems))
    print("struktur ok")
    vals = lo_values(tmp, formula_cells(tmp))
    final = Book(tmp)
    for name, v in vals.items():
        final.set_cached(name, v)
    final.save(out)
    print(f"cachade värden: {sum(len(v) for v in vals.values())} celler → {out}")


def gate(path: Path) -> None:
    sys.modules["comrun"] = types.SimpleNamespace(run=lo_gate.run, probe=lo_gate.probe,
                                                  probe_dv=lo_gate.probe_dv, scan=lo_gate.scan)
    import regression as R
    print("Baslinje källa …", end=" ", flush=True)
    R.check_source_baseline(HERE / "source" / "LM371_source.xlsx")
    print("ok")
    R.run_all(path)


def main() -> int:
    if len(sys.argv) < 4:
        raise SystemExit(__doc__)
    base, out = Path(sys.argv[1]), Path(sys.argv[2])
    build(base, out, sys.argv[3:])
    gate(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
