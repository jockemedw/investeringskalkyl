"""XML-applier för LM 371 — samma ops som apply.ps1 (Excel COM), men direkt i xlsx-paketets XML.

Används när Excel saknas (Linux/moln). Till skillnad från openpyxl skrivs bara de delar som ops:arna rör;
allt annat i paketet (x14-valideringar, kommentarer, VML, customXml, skrivarinställningar) kopieras byte för
byte. calcChain tas bort och fullCalcOnLoad sätts, så Excel räknar om hela boken vid öppning.

Stöd: set, fmt_from, merge, unmerge, clear, clear_contents, align, wrap, dv_list, dv_custom, print_area
(probe/scan ignoreras). Övriga ops (replace, insert_rows, datatable, fill, page_setup …) kräver Excel.

Delade formler: skrivs en cell som är master i en delad formelgrupp materialiseras gruppens övriga celler
först (openpyxl Translator), annars skulle de tappa sin formel.
"""
from __future__ import annotations
import copy
import re
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import (column_index_from_string, coordinate_from_string, get_column_letter,
                                 range_boundaries)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
H_ALIGN = {-4108: "center", -4131: "left", -4152: "right"}
# Ordning av barn till <worksheet> (ECMA-376) — nya element placeras rätt
WS_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
            "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
            "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations",
            "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
            "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing",
            "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]


def q(tag: str) -> str:
    return f"{{{NS}}}{tag}"


def rc(ref: str) -> tuple[int, int]:
    col, row = coordinate_from_string(ref.replace("$", ""))
    return row, column_index_from_string(col)


def cells_in(ref: str):
    c1, r1, c2, r2 = range_boundaries(ref.replace("$", ""))
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            yield r, c


def addr(r: int, c: int) -> str:
    return f"{get_column_letter(c)}{r}"


class Sheet:
    def __init__(self, name: str, part: str, root):
        self.name, self.part, self.root = name, part, root
        self.data = root.find(q("sheetData"))
        self.rows = {int(row.get("r")): row for row in self.data.findall(q("row"))}
        self._shared_done: set[str] = set()

    # ── celler ──
    def find(self, r: int, c: int):
        row = self.rows.get(r)
        if row is None:
            return None
        ref = addr(r, c)
        for cell in row.findall(q("c")):
            if cell.get("r") == ref:
                return cell
        return None

    def cell(self, r: int, c: int):
        found = self.find(r, c)
        if found is not None:
            return found
        row = self.rows.get(r)
        if row is None:
            row = etree.Element(q("row"), r=str(r))
            later = [k for k in self.rows if k > r]
            if later:
                self.rows[min(later)].addprevious(row)
            else:
                self.data.append(row)
            self.rows[r] = row
        row.attrib.pop("spans", None)                       # valfri optimering; ogiltig när celler tillkommer
        new = etree.Element(q("c"), r=addr(r, c))
        for existing in row.findall(q("c")):
            if rc(existing.get("r"))[1] > c:
                existing.addprevious(new)
                break
        else:
            row.append(new)
        return new

    def child(self, tag: str, create: bool = True):
        el = self.root.find(q(tag))
        if el is None and create:
            el = etree.Element(q(tag))
            idx = WS_ORDER.index(tag)
            for later in WS_ORDER[idx + 1:]:
                nxt = self.root.find(q(later))
                if nxt is not None:
                    nxt.addprevious(el)
                    break
            else:
                self.root.append(el)
        return el

    # ── delade formler ──
    def unshare_if_master(self, cells: set[tuple[int, int]]) -> None:
        """Materialisera delade formelgrupper vars master skrivs över/rensas."""
        masters = {}
        for row in self.rows.values():
            for cell in row.findall(q("c")):
                f = cell.find(q("f"))
                if f is not None and f.get("t") == "shared" and f.get("ref") and f.text:
                    masters[f.get("si")] = cell
        for si, master in masters.items():
            if si in self._shared_done or rc(master.get("r")) not in cells:
                continue
            mf = master.find(q("f"))
            text, origin = "=" + mf.text, master.get("r")
            for (r, c) in cells_in(mf.get("ref")):
                cell = self.find(r, c)
                if cell is None:
                    continue
                f = cell.find(q("f"))
                if f is None or f.get("t") != "shared" or f.get("si") != si:
                    continue
                f.text = Translator(text, origin=origin).translate_formula(cell.get("r"))[1:]
                for a in ("t", "si", "ref"):
                    f.attrib.pop(a, None)
            self._shared_done.add(si)

    @staticmethod
    def strip_content(cell) -> None:
        for tag in ("f", "v", "is"):
            for el in cell.findall(q(tag)):
                cell.remove(el)
        cell.attrib.pop("t", None)

    def update_dimension(self) -> None:
        dim = self.root.find(q("dimension"))
        if dim is None or not self.rows:
            return
        r1, c1, r2, c2 = 10**9, 10**9, 0, 0
        for r, row in self.rows.items():
            for cell in row.findall(q("c")):
                _, c = rc(cell.get("r"))
                r1, r2, c1, c2 = min(r1, r), max(r2, r), min(c1, c), max(c2, c)
        if r2:
            dim.set("ref", f"{addr(r1, c1)}:{addr(r2, c2)}")


class Book:
    def __init__(self, path: Path):
        with zipfile.ZipFile(path) as z:
            self.parts = {i.filename: z.read(i.filename) for i in z.infolist()}
            self.order = [i.filename for i in z.infolist()]
        self.wb = etree.fromstring(self.parts["xl/workbook.xml"])
        rels = etree.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
        target = {r.get("Id"): r.get("Target") for r in rels}
        self.sheet_names, self.sheet_part = [], {}
        for s in self.wb.find(q("sheets")):
            name = s.get("name")
            self.sheet_names.append(name)
            self.sheet_part[name] = "xl/" + target[s.get(f"{{{RNS}}}id")].lstrip("/").replace("xl/", "")
        self._sheets: dict[str, Sheet] = {}
        self._styles = None
        self._xf_cache: dict[tuple, int] = {}

    def sheet(self, name: str) -> Sheet:
        if name not in self._sheets:
            part = self.sheet_part[name]
            self._sheets[name] = Sheet(name, part, etree.fromstring(self.parts[part]))
        return self._sheets[name]

    # ── stilar ──
    @property
    def styles(self):
        if self._styles is None:
            self._styles = etree.fromstring(self.parts["xl/styles.xml"])
        return self._styles

    def xf_with(self, s: int, key: tuple, mutate) -> int:
        ck = (s, key)
        if ck not in self._xf_cache:
            xfs = self.styles.find(q("cellXfs"))
            xf = copy.deepcopy(xfs[s])
            mutate(xf)
            xfs.append(xf)
            xfs.set("count", str(len(xfs)))
            self._xf_cache[ck] = len(xfs) - 1
        return self._xf_cache[ck]

    @staticmethod
    def _align(**kw):
        def mutate(xf):
            al = xf.find(q("alignment"))
            if al is None:
                al = etree.SubElement(xf, q("alignment"))
                ext = xf.find(q("protection"))
                if ext is not None:                             # alignment före protection
                    ext.addprevious(al)
            for k, v in kw.items():
                al.set(k, v)
            xf.set("applyAlignment", "1")
        return mutate

    # ── ops ──
    def apply(self, ops: list[dict]) -> None:
        for o in ops:
            getattr(self, "op_" + o["op"])(o)

    def op_probe(self, o): pass
    op_scan = op_probe_dv = op_probe

    def op_set(self, o):
        sh = self.sheet(o["sheet"])
        r, c = rc(o["ref"])
        sh.unshare_if_master({(r, c)})
        cell = sh.cell(r, c)
        sh.strip_content(cell)
        if o.get("formula") is not None:
            f = etree.SubElement(cell, q("f"))
            f.text = o["formula"][1:] if o["formula"].startswith("=") else o["formula"]
        elif isinstance(o.get("value"), str):
            if o["value"] == "":
                return
            cell.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(cell, q("is")), q("t"))
            t.text = o["value"]
            if t.text != t.text.strip():
                t.set(XML_SPACE, "preserve")
        else:
            v = float(o["value"])
            etree.SubElement(cell, q("v")).text = str(int(v)) if v.is_integer() and abs(v) < 1e15 else repr(v)

    def op_clear_contents(self, o):
        sh = self.sheet(o["sheet"])
        targets = set(cells_in(o["ref"]))
        sh.unshare_if_master(targets)
        for r, c in targets:
            cell = sh.find(r, c)
            if cell is not None:
                sh.strip_content(cell)

    def op_clear(self, o):
        self.op_clear_contents(o)
        sh = self.sheet(o["sheet"])
        for r, c in cells_in(o["ref"]):
            cell = sh.find(r, c)
            if cell is not None:
                cell.attrib.pop("s", None)

    def op_fmt_from(self, o):
        """Kopiera cellformat (stilindex är bokgemensamt; src_sheet för källa på annat blad)."""
        sh = self.sheet(o["sheet"])
        ssh = self.sheet(o.get("src_sheet", o["sheet"]))
        sc1, sr1, sc2, sr2 = range_boundaries(o["src"].replace("$", ""))
        h, w = sr2 - sr1 + 1, sc2 - sc1 + 1
        dc1, dr1, _, _ = range_boundaries(o["dst"].replace("$", ""))
        for r, c in cells_in(o["dst"]):
            src = ssh.find(sr1 + (r - dr1) % h, sc1 + (c - dc1) % w)
            s = src.get("s") if src is not None else None
            cell = sh.cell(r, c)
            if s is None:
                cell.attrib.pop("s", None)
            else:
                cell.set("s", s)

    def _restyle(self, o, key, mutate):
        sh = self.sheet(o["sheet"])
        for r, c in cells_in(o["ref"]):
            cell = sh.cell(r, c)
            cell.set("s", str(self.xf_with(int(cell.get("s", "0")), key, mutate)))

    def op_align(self, o):
        kw = {}
        if "h" in o:
            kw["horizontal"] = H_ALIGN[int(o["h"])]
        if "v" in o:
            kw["vertical"] = o["v"]                             # top | center | bottom
        self._restyle(o, ("align",) + tuple(sorted(kw.items())), self._align(**kw))

    def op_wrap(self, o):
        self._restyle(o, ("wrap",), self._align(wrapText="1"))

    def _font_flag(self, o, tag: str):
        """Kursiv/fet: klona cellens font med/utan <i/>/<b/> (on=False tar bort) och xf som pekar på den."""
        fonts = self.styles.find(q("fonts"))
        on = o.get("on", True)

        def mutate(xf):
            font = copy.deepcopy(fonts[int(xf.get("fontId", "0"))])
            for el in font.findall(q(tag)):
                font.remove(el)
            if on:
                font.insert(0 if tag == "b" else (1 if font.find(q("b")) is not None else 0), etree.Element(q(tag)))
            fonts.append(font)
            fonts.set("count", str(len(fonts)))
            xf.set("fontId", str(len(fonts) - 1))
            xf.set("applyFont", "1")
        self._restyle(o, ("font", tag, on), mutate)

    def op_italic(self, o):
        self._font_flag(o, "i")

    def op_bold(self, o):
        self._font_flag(o, "b")

    def op_merge(self, o):
        sh = self.sheet(o["sheet"])
        mc = sh.child("mergeCells")
        ref = o["ref"].replace("$", "")
        new = set(cells_in(ref))
        for m in mc.findall(q("mergeCell")):
            if m.get("ref") == ref:
                return
            if new & set(cells_in(m.get("ref"))):
                raise ValueError(f"merge {ref} överlappar {m.get('ref')}")
        etree.SubElement(mc, q("mergeCell"), ref=ref)
        mc.set("count", str(len(mc)))

    def op_unmerge(self, o):
        sh = self.sheet(o["sheet"])
        mc = sh.child("mergeCells", create=False)
        if mc is None:
            return
        for m in mc.findall(q("mergeCell")):
            if m.get("ref") == o["ref"].replace("$", ""):
                mc.remove(m)
        mc.set("count", str(len(mc)))

    def _dv(self, o, attrs: dict, formula: str):
        sh = self.sheet(o["sheet"])
        dvs = sh.child("dataValidations")
        ref = o["ref"].replace("$", "")
        target = set(cells_in(ref))
        for dv in dvs.findall(q("dataValidation")):
            covered = set().union(*(set(cells_in(x)) for x in dv.get("sqref").split()))
            if covered & target:
                if covered <= target:
                    dvs.remove(dv)
                else:
                    raise ValueError(f"validering {dv.get('sqref')} överlappar {ref} delvis")
        dv = etree.SubElement(dvs, q("dataValidation"), **attrs, sqref=ref)
        etree.SubElement(dv, q("formula1")).text = formula[1:] if formula.startswith("=") else formula
        dvs.set("count", str(len(dvs)))

    def op_dv_list(self, o):
        attrs = {"type": "list", "allowBlank": "1", "showInputMessage": "1", "showErrorMessage": "1"}
        if o.get("prompt") is not None:
            attrs |= {"promptTitle": o.get("prompt_title", ""), "prompt": o["prompt"]}
        self._dv(o, attrs, o["source"])

    def op_dv_custom(self, o):
        attrs = {"type": "custom", "allowBlank": "1", "showInputMessage": "1", "showErrorMessage": "1",
                 "errorTitle": o["err_title"], "error": o["err_msg"]}
        if o.get("prompt") is not None:
            attrs |= {"promptTitle": o.get("prompt_title", ""), "prompt": o["prompt"]}
        self._dv(o, attrs, o["formula"])

    def op_print_area(self, o):
        idx = str(self.sheet_names.index(o["sheet"]))
        area = o["area"]
        quoted = "'" + o["sheet"].replace("'", "''") + "'"
        for dn in self.wb.find(q("definedNames")):
            if dn.get("name") == "_xlnm.Print_Area" and dn.get("localSheetId") == idx:
                dn.text = f"{quoted}!{area}"
                return
        raise ValueError(f"utskriftsområde saknas för {o['sheet']}")

    # ── designops (v3.2) ──
    def _xf_from(self, o, attr: str, apply_attr: str):
        """Kopiera ett stilattribut (fillId/fontId/borderId) från källcellens xf, behåll resten."""
        sh = self.sheet(o.get("src_sheet", o["sheet"]))
        sr, sc = rc(o["src"])
        src = sh.find(sr, sc)
        val = self.styles.find(q("cellXfs"))[int(src.get("s", "0")) if src is not None else 0].get(attr, "0")

        def mutate(xf):
            xf.set(attr, val)
            xf.set(apply_attr, "1")
        self._restyle(o, (attr, val), mutate)

    def op_fill_from(self, o):
        self._xf_from(o, "fillId", "applyFill")

    def op_font_from(self, o):
        self._xf_from(o, "fontId", "applyFont")

    def op_border_from(self, o):
        self._xf_from(o, "borderId", "applyBorder")

    def op_fontcolor(self, o):
        """color: 'white' | 'black' | 'grey' | 'RRGGBB'."""
        fonts = self.styles.find(q("fonts"))
        spec = o["color"]

        def mutate(xf):
            font = copy.deepcopy(fonts[int(xf.get("fontId", "0"))])
            for el in font.findall(q("color")):
                font.remove(el)
            col = etree.Element(q("color"))
            if spec == "white":
                col.set("theme", "0")
            elif spec == "black":
                col.set("theme", "1")
            elif spec == "grey":
                col.set("rgb", "FF7F7F7F")
            else:
                col.set("rgb", "FF" + spec.upper())
            # color placeras före name/family/scheme enligt CT_Font (b,i,strike,…,sz,color,name,family,charset,scheme)
            anchor = next((font.find(q(t)) for t in ("name", "family", "charset", "scheme") if font.find(q(t)) is not None), None)
            if anchor is not None:
                anchor.addprevious(col)
            else:
                font.append(col)
            fonts.append(font)
            fonts.set("count", str(len(fonts)))
            xf.set("fontId", str(len(fonts) - 1))
            xf.set("applyFont", "1")
        self._restyle(o, ("fontcolor", spec), mutate)

    def op_numfmt(self, o):
        """fmt = formatkod i XML-syntax (en-US: ',' tusental, '.' decimal) — Excel visar enligt svensk locale."""
        st = self.styles
        nf = st.find(q("numFmts"))
        if nf is None:
            nf = etree.Element(q("numFmts"))
            st.insert(0, nf)
        code = o["fmt"]
        fid = next((int(n.get("numFmtId")) for n in nf if n.get("formatCode") == code), None)
        if fid is None:
            fid = max([163] + [int(n.get("numFmtId")) for n in nf]) + 1
            etree.SubElement(nf, q("numFmt"), numFmtId=str(fid), formatCode=code)
            nf.set("count", str(len(nf)))

        def mutate(xf):
            xf.set("numFmtId", str(fid))
            xf.set("applyNumberFormat", "1")
        self._restyle(o, ("numfmt", code), mutate)

    def _cols(self, sh):
        """Expandera <cols> till {kolumn: attrib} och returnera (cols-element, dict)."""
        cols = sh.child("cols")
        spec = {}
        for col in cols.findall(q("col")):
            for c in range(int(col.get("min")), int(col.get("max")) + 1):
                spec[c] = {k: v for k, v in col.attrib.items() if k not in ("min", "max")}
        return cols, spec

    @staticmethod
    def _write_cols(cols, spec):
        for col in list(cols):
            cols.remove(col)
        run = None
        for c in sorted(spec):
            a = spec[c]
            if run and run[1] == c - 1 and run[2] == a:
                run[1] = c
            else:
                if run:
                    etree.SubElement(cols, q("col"), min=str(run[0]), max=str(run[1]), **run[2])
                run = [c, c, a]
        if run:
            etree.SubElement(cols, q("col"), min=str(run[0]), max=str(run[1]), **run[2])

    def _colrange(self, ref):
        c1, _, c2, _ = range_boundaries(ref.replace("$", "") if any(ch.isdigit() for ch in ref) else f"{ref.split(':')[0]}1:{ref.split(':')[-1]}1")
        return range(c1, c2 + 1)

    def op_colwidth(self, o):
        sh = self.sheet(o["sheet"])
        cols, spec = self._cols(sh)
        for c in self._colrange(o["ref"]):
            a = spec.setdefault(c, {})
            a["width"] = repr(float(o["w"]))
            a["customWidth"] = "1"
        self._write_cols(cols, spec)

    def _hide_cols(self, o, hidden: bool):
        sh = self.sheet(o["sheet"])
        cols, spec = self._cols(sh)
        for c in self._colrange(o["ref"]):
            a = spec.setdefault(c, {"width": "9.140625"})
            if hidden:
                a["hidden"] = "1"
            else:
                a.pop("hidden", None)
        self._write_cols(cols, spec)

    def op_hide_cols(self, o):
        self._hide_cols(o, True)

    def op_unhide_cols(self, o):
        self._hide_cols(o, False)

    def _rows_of(self, o):
        ref = str(o.get("ref", o.get("row")))
        a, _, b = ref.partition(":")
        return range(int(re.sub(r"\D", "", a)), int(re.sub(r"\D", "", b or a)) + 1)

    def _row_el(self, sh, r):
        if r not in sh.rows:
            sh.cell(r, 1)                                       # skapar raden (cellen A är tom och ofarlig)
            row = sh.rows[r]
            for c in row.findall(q("c")):
                if c.get("r") == addr(r, 1) and not len(c) and "s" not in c.attrib:
                    row.remove(c)
        return sh.rows[r]

    def op_rowheight(self, o):
        sh = self.sheet(o["sheet"])
        for r in self._rows_of(o):
            row = self._row_el(sh, r)
            row.set("ht", repr(float(o["h"])))
            row.set("customHeight", "1")

    def op_hide_rows(self, o):
        sh = self.sheet(o["sheet"])
        for r in self._rows_of(o):
            self._row_el(sh, r).set("hidden", "1")

    def op_unhide_rows(self, o):
        sh = self.sheet(o["sheet"])
        for r in self._rows_of(o):
            self._row_el(sh, r).attrib.pop("hidden", None)

    def op_dv_prompt(self, o):
        """Inmatningsmeddelande (tooltip). Finns validering på exakt samma område uppdateras dess prompt."""
        sh = self.sheet(o["sheet"])
        dvs = sh.child("dataValidations")
        ref = o["ref"].replace("$", "")
        for dv in dvs.findall(q("dataValidation")):
            if dv.get("sqref") == ref:
                dv.set("showInputMessage", "1")
                dv.set("promptTitle", o.get("prompt_title", ""))
                dv.set("prompt", o["prompt"])
                return
        target = set(cells_in(ref))
        for dv in dvs.findall(q("dataValidation")):
            if target & set().union(*(set(cells_in(x)) for x in dv.get("sqref").split())):
                raise ValueError(f"dv_prompt {ref} överlappar validering {dv.get('sqref')}")
        etree.SubElement(dvs, q("dataValidation"), allowBlank="1", showInputMessage="1",
                         promptTitle=o.get("prompt_title", ""), prompt=o["prompt"], sqref=ref)
        dvs.set("count", str(len(dvs)))

    def op_freeze(self, o):
        """Lås rader ovanför och kolumner till vänster om ref (som Excels Frys fönsterrutor)."""
        sh = self.sheet(o["sheet"])
        sv = sh.root.find(q("sheetViews")).find(q("sheetView"))
        for el in sv.findall(q("pane")) + sv.findall(q("selection")):
            sv.remove(el)
        r, c = rc(o["ref"])
        xs, ys = c - 1, r - 1
        if xs == 0 and ys == 0:
            return
        pane_name = "bottomRight" if xs and ys else ("bottomLeft" if ys else "topRight")
        attrs = {"topLeftCell": o["ref"].replace("$", ""), "activePane": pane_name, "state": "frozen"}
        if xs:
            attrs["xSplit"] = str(xs)
        if ys:
            attrs["ySplit"] = str(ys)
        sv.insert(0, etree.Element(q("pane"), **attrs))
        etree.SubElement(sv, q("selection"), pane=pane_name, activeCell=o["ref"].replace("$", ""),
                         sqref=o["ref"].replace("$", ""))

    def op_page_setup(self, o):
        """Delmängd: orient (1 stående/2 liggande), paper (8 A3/9 A4), fit_wide/fit_tall, area, foot_c."""
        sh = self.sheet(o["sheet"])
        ps = sh.child("pageSetup")
        if "orient" in o:
            ps.set("orientation", "portrait" if int(o["orient"]) == 1 else "landscape")
        if "paper" in o:
            ps.set("paperSize", str(o["paper"]))
        if "fit_wide" in o or "fit_tall" in o:
            ps.set("fitToWidth", str(o.get("fit_wide", 1)))
            ps.set("fitToHeight", str(o.get("fit_tall", 0)))
            pr = sh.child("sheetPr")
            psp = pr.find(q("pageSetUpPr"))
            if psp is None:
                psp = etree.SubElement(pr, q("pageSetUpPr"))     # sist i CT_SheetPr (tabColor, outlinePr, pageSetUpPr)
            psp.set("fitToPage", "1")
        if "area" in o:
            self.op_print_area({"sheet": o["sheet"], "area": o["area"]})

    def op_footer_replace(self, o):
        """Ersätt text i sidhuvud/sidfot (alla varianter) på bladet."""
        sh = self.sheet(o["sheet"])
        hf = sh.root.find(q("headerFooter"))
        if hf is None:
            return
        for el in hf:
            if el.text:
                el.text = el.text.replace(o["find"], o["repl"])

    def op_row_break(self, o):
        """Manuell sidbrytning efter rad `after` (Excel: före rad after+1)."""
        sh = self.sheet(o["sheet"])
        rb = sh.child("rowBreaks")
        ids = {b.get("id") for b in rb}
        if str(o["after"]) not in ids:
            etree.SubElement(rb, q("brk"), id=str(o["after"]), max="16383", man="1")
        brks = sorted(rb, key=lambda b: int(b.get("id")))
        for b in list(rb):
            rb.remove(b)
        for b in brks:
            rb.append(b)
        rb.set("count", str(len(rb)))
        rb.set("manualBreakCount", str(sum(1 for b in rb if b.get("man") == "1")))

    def op_print_titles(self, o):
        """Utskriftsrubriker: cols '$A:$B' och/eller rows '$4:$4'."""
        idx = str(self.sheet_names.index(o["sheet"]))
        quoted = "'" + o["sheet"].replace("'", "''") + "'"
        parts = [f"{quoted}!{x}" for x in (o.get("cols"), o.get("rows")) if x]
        dns = self.wb.find(q("definedNames"))
        for dn in dns:
            if dn.get("name") == "_xlnm.Print_Titles" and dn.get("localSheetId") == idx:
                dn.text = ",".join(parts)
                return
        el = etree.SubElement(dns, q("definedName"), name="_xlnm.Print_Titles", localSheetId=idx)
        el.text = ",".join(parts)
        items = sorted(dns, key=lambda d: (d.get("name").lower(), int(d.get("localSheetId", "-1"))))
        for d in list(dns):
            dns.remove(d)
        for d in items:
            dns.append(d)

    def op_page_order(self, o):
        self.sheet(o["sheet"]).child("pageSetup").set("pageOrder", o["order"])   # overThenDown | downThenOver

    def op_x14_dv(self, o):
        """x14-validering (listor med källa på annat blad). match = befintlig sqref; ny sqref/prompt/disablePrompts."""
        sh = self.sheet(o["sheet"])
        X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        XM = "http://schemas.microsoft.com/office/excel/2006/main"
        coll = sh.root.find(f".//{{{X14}}}dataValidations")
        if coll is None:
            raise ValueError(f"inga x14-valideringar på {o['sheet']}")
        if o.get("enable_prompts"):
            coll.attrib.pop("disablePrompts", None)
        if "match" not in o:
            return
        for dv in coll.findall(f"{{{X14}}}dataValidation"):
            sq = dv.find(f"{{{XM}}}sqref")
            if set(sq.text.split()) == set(o["match"].split()):
                if "sqref" in o:
                    sq.text = o["sqref"]
                if "prompt" in o:
                    dv.set("showInputMessage", "1")
                    dv.set("promptTitle", o.get("prompt_title", ""))
                    dv.set("prompt", o["prompt"])
                if "error" in o:
                    dv.set("showErrorMessage", "1")
                    dv.set("error", o["error"])
                return
        raise ValueError(f"x14-validering {o['match']} saknas på {o['sheet']}")

    def __getattr__(self, name):
        if name.startswith("op_"):
            raise NotImplementedError(f"{name[3:]} kräver Excel COM (apply.ps1)")
        raise AttributeError(name)

    # ── cachade värden ──
    def set_cached(self, sheet: str, values: dict[str, object]) -> None:
        """Skriv cachade resultat för formelceller (str → t="str", bool → t="b", tal → <v>)."""
        sh = self.sheet(sheet)
        for ref, v in values.items():
            r, c = rc(ref)
            cell = sh.find(r, c)
            if cell is None:
                continue
            for el in cell.findall(q("v")):
                cell.remove(el)
            cell.attrib.pop("t", None)
            if v is None:
                continue
            ve = etree.SubElement(cell, q("v"))
            if isinstance(v, bool):
                cell.set("t", "b")
                ve.text = "1" if v else "0"
            elif isinstance(v, str):
                if v.startswith("#"):
                    cell.set("t", "e")
                else:
                    cell.set("t", "str")
                ve.text = v
            else:
                ve.text = repr(float(v))

    # ── spara ──
    def save(self, dest: Path) -> None:
        for sh in self._sheets.values():
            sh.update_dimension()
            self.parts[sh.part] = etree.tostring(sh.root, xml_declaration=True, encoding="UTF-8", standalone=True)
        if self._styles is not None:
            self.parts["xl/styles.xml"] = etree.tostring(self._styles, xml_declaration=True, encoding="UTF-8",
                                                         standalone=True)
        # Excel räknar om hela boken vid öppning; calcChain (gamla formelceller) tas bort
        calc = self.wb.find(q("calcPr"))
        calc.set("fullCalcOnLoad", "1")
        self.parts["xl/workbook.xml"] = etree.tostring(self.wb, xml_declaration=True, encoding="UTF-8",
                                                       standalone=True)
        if "xl/calcChain.xml" in self.parts:
            del self.parts["xl/calcChain.xml"]
            rels = etree.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
            for r in list(rels):
                if r.get("Target").endswith("calcChain.xml"):
                    rels.remove(r)
            self.parts["xl/_rels/workbook.xml.rels"] = etree.tostring(rels, xml_declaration=True,
                                                                      encoding="UTF-8", standalone=True)
            ct = etree.fromstring(self.parts["[Content_Types].xml"])
            for o in list(ct):
                if o.get("PartName") == "/xl/calcChain.xml":
                    ct.remove(o)
            self.parts["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8",
                                                               standalone=True)
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
            for name in self.order:
                if name in self.parts:
                    z.writestr(name, self.parts[name])


def validate(path: Path) -> list[str]:
    """Strukturkontroll av de skrivna bladen: ordning, dubbletter, delade formler, merge-överlapp."""
    problems = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
                continue
            root = etree.fromstring(z.read(name))
            last_r, masters, users = 0, set(), set()
            for row in root.find(q("sheetData")):
                r = int(row.get("r"))
                if r <= last_r:
                    problems.append(f"{name}: rad {r} i fel ordning")
                last_r, last_c = r, 0
                for cell in row:
                    rr, cc = rc(cell.get("r"))
                    if rr != r or cc <= last_c:
                        problems.append(f"{name}: cell {cell.get('r')} i fel ordning/rad")
                    last_c = cc
                    f = cell.find(q("f"))
                    if f is not None and f.get("t") == "shared":
                        (masters if f.get("ref") else users).add(f.get("si"))
            if users - masters:
                problems.append(f"{name}: delade formler utan master: {sorted(users - masters)[:5]}")
            mc = root.find(q("mergeCells"))
            seen = {}
            for m in (mc if mc is not None else []):
                for p in cells_in(m.get("ref")):
                    if p in seen:
                        problems.append(f"{name}: merge {m.get('ref')} överlappar {seen[p]}")
                        break
                    seen[p] = m.get("ref")
    return problems
