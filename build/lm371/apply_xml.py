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
        sh = self.sheet(o["sheet"])
        sc1, sr1, sc2, sr2 = range_boundaries(o["src"].replace("$", ""))
        h, w = sr2 - sr1 + 1, sc2 - sc1 + 1
        dc1, dr1, _, _ = range_boundaries(o["dst"].replace("$", ""))
        for r, c in cells_in(o["dst"]):
            src = sh.find(sr1 + (r - dr1) % h, sc1 + (c - dc1) % w)
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
        h = H_ALIGN[int(o["h"])]
        self._restyle(o, ("h", h), self._align(horizontal=h))

    def op_wrap(self, o):
        self._restyle(o, ("wrap",), self._align(wrapText="1"))

    def _font_flag(self, o, tag: str):
        """Kursiv/fet: klona cellens font med <i/>/<b/> och xf som pekar på den."""
        fonts = self.styles.find(q("fonts"))

        def mutate(xf):
            font = copy.deepcopy(fonts[int(xf.get("fontId", "0"))])
            if font.find(q(tag)) is None:
                font.insert(0, etree.Element(q(tag)))
            fonts.append(font)
            fonts.set("count", str(len(fonts)))
            xf.set("fontId", str(len(fonts) - 1))
            xf.set("applyFont", "1")
        self._restyle(o, ("font", tag), mutate)

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
