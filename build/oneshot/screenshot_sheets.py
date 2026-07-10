"""Skärmdumpar av varje fliks faktiska öppningsvy — sanningen för design-polish.

Öppnar xlsx i Excel COM med SYNLIGT maximerat fönster (ReadOnly — krockar inte
med en redan öppen instans), 100 % zoom, och tar en skärmdump per flik.

VIKTIGT RENDERINGSFYND: Excel målar INTE om rutnätet vid programmatisk scroll
(ScrollRow/Goto/Select uppdaterar modellen men pixlarna förblir öppningsvyn;
zoom-ändringar målar dock om). Scrollade vyer löses därför via XML-patch av
pane@topLeftCell i en TEMPORÄR KOPIA av filen — Excel ritar då rätt vy direkt
vid öppning. Originalfilen röres aldrig.

COM via PowerShell (mönster: tools/recalc.py) — ingen pywin32.

Kör:  python build/oneshot/screenshot_sheets.py [xlsx] --out DIR [--sheets A,B]
"""
from __future__ import annotations
import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
DEFAULT_XLSX = ROOT / "build" / "oneshot" / "Investeringskalkyl_v2.xlsx"

PS_COMMON = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public struct RECT { public int Left, Top, Right, Bottom; }
public class Win {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
}
'@
[Win]::SetProcessDPIAware() | Out-Null

# PrintWindow (flagga 2 = PW_RENDERFULLCONTENT) fångar Excel-fönstrets pixlar
# även när det ligger bakom andra fönster — stjäl inte fokus från användaren
# och kan inte förorenas av andra appar (till skillnad från CopyFromScreen).
function Snap($path) {
    Start-Sleep -Milliseconds 500
    $r = New-Object RECT
    [Win]::GetWindowRect([IntPtr]$excel.Hwnd, [ref]$r) | Out-Null
    $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    [Win]::PrintWindow([IntPtr]$excel.Hwnd, $hdc, 2) | Out-Null
    $g.ReleaseHdc($hdc)
    $g.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}
"""

# Pass 1: öppningsvyer + mät synliga rader / sista innehållsrad per flik
PS_PASS1 = PS_COMMON + r"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open('__XLSX__', 0, $true)
    $excel.WindowState = -4137
    $only = @(__SHEETS__)
    $i = 0
    foreach ($ws in $wb.Worksheets) {
        $i++
        if ($ws.Visible -ne -1) { continue }
        if ($only.Count -gt 0 -and $only -notcontains $ws.Name) { continue }
        $ws.Activate()
        $win = $excel.ActiveWindow
        $win.Zoom = 100
        $safe = $ws.Name -replace '[^\w]', '_'
        Snap "__OUT__\$($i.ToString('00'))_$($safe)_r001.png"
        $vis = $win.VisibleRange.Rows.Count
        $last = 1
        $found = $ws.Cells.Find('*', $ws.Cells.Item(1,1), -4123, 2, 1, 2)
        if ($found) { $last = $found.Row }
        # Till fil, inte stdout — subprocess-pipen tappar intermittent output
        Add-Content -Path '__MEASURE__' -Value "MEASURE|$($ws.Name)|$i|$last|$vis" -Encoding UTF8
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
"""

# Pass 2: öppna patchad kopia och snappa angivna flikar (öppningsvyn = rätt scroll)
PS_PASS2 = PS_COMMON + r"""
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open('__XLSX__', 0, $true)
    $excel.WindowState = -4137
    __SNAPS__
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
"""


def _run_ps(script: str, timeout: int = 300) -> str:
    # encoding + errors: OEM-kodade å/ä/ö i PS-output kraschar annars
    # subprocess-lästråden → stdout None trots exit 0 (rotorsaken bakom
    # de "intermittenta" tomma svaren)
    res = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
        timeout=timeout, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
    )
    if res.returncode != 0:
        raise RuntimeError(f"powershell fail:\n{res.stdout}\n{res.stderr}")
    return res.stdout or ""


def _sheet_files(src: Path) -> dict[int, str]:
    """Tab-position (1-baserad) → worksheets/sheetN.xml. sheetN.xml följer
    skapandeordningen, inte flikordningen — måste lösas via workbook.xml.rels."""
    with zipfile.ZipFile(src) as z:
        wbxml = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_to_file: dict[str, str] = {}
    for rel in re.findall(r"<Relationship\b[^>]*>", rels):
        rid = re.search(r'Id="(rId\d+)"', rel)
        tgt = re.search(r'Target="(?:/xl/)?worksheets/([^"]+)"', rel)
        if rid and tgt:
            rid_to_file[rid.group(1)] = tgt.group(1)
    order = re.findall(r'<sheet[^>]*r:id="(rId\d+)"', wbxml)
    return {i + 1: rid_to_file[rid].split("/")[-1]
            for i, rid in enumerate(order) if rid in rid_to_file}


def _patch_toplef(src: Path, dst: Path, scroll: dict[int, int]) -> None:
    """Kopiera xlsx med pane/sheetView topLeftCell satt till rad R per
    tab-position (1-baserad)."""
    tabmap = _sheet_files(src)
    names = {tabmap[i]: r for i, r in scroll.items() if i in tabmap}
    with zipfile.ZipFile(src, "r") as zin, \
         zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.namelist():
            data = zin.read(item)
            base = item.split("/")[-1]
            if base in names:
                r = names[base]
                text = data.decode("utf-8")
                # Vid ren kolumnfrys (xSplit utan ySplit) styrs radscrollen av
                # sheetView@topLeftCell — pane@topLeftCell räcker inte. Sätt båda.
                text = re.sub(r'(<sheetView[^>]*?)\s+topLeftCell="[A-Z]+\d+"',
                              r"\1", text, count=1)
                text = re.sub(r"(<sheetView )", rf'\g<1>topLeftCell="A{r}" ',
                              text, count=1)
                if "<pane " in text:
                    text = re.sub(r'(<pane[^>]*?topLeftCell=")[A-Z]+\d+(")',
                                  rf"\g<1>B{r}\g<2>", text, count=1)
                    if f'topLeftCell="B{r}"' not in text:  # pane utan topLeftCell
                        text = text.replace("<pane ", f'<pane topLeftCell="B{r}" ', 1)
                data = text.encode("utf-8")
            zout.writestr(item, data)


def screenshot(xlsx: Path, out_dir: Path, sheets: list[str] | None = None,
               timeout: int = 600) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx = xlsx.resolve()
    ps_sheets = ", ".join(f"'{s}'" for s in (sheets or []))

    measure_file = out_dir / "_measure.txt"
    measure_file.unlink(missing_ok=True)
    _run_ps(PS_PASS1
            .replace("__XLSX__", str(xlsx))
            .replace("__OUT__", str(out_dir))
            .replace("__MEASURE__", str(measure_file))
            .replace("__SHEETS__", ps_sheets), timeout)
    out1 = measure_file.read_text(encoding="utf-8-sig") if measure_file.exists() else ""
    measure_file.unlink(missing_ok=True)

    # MEASURE|namn|index|lastRow|visRows
    pages: dict[int, dict] = {}  # sida k -> {sheet_index: (namn, rad)}
    for line in out1.splitlines():
        if not line.startswith("MEASURE|"):
            continue
        _, name, idx, last, vis = line.split("|")
        idx, last, vis = int(idx), int(last), int(vis)
        if vis < 4:
            continue
        r, k = 1 + vis, 1
        while r <= last:
            pages.setdefault(k, {})[idx] = (name, r)
            r += vis
            k += 1

    tmpdir = Path(tempfile.mkdtemp(prefix="sheetshots_"))
    try:
        for k, entries in sorted(pages.items()):
            copy = tmpdir / f"page{k}.xlsx"
            _patch_toplef(xlsx, copy, {i: r for i, (_, r) in entries.items()})
            snaps = []
            for i, (name, r) in sorted(entries.items()):
                safe = re.sub(r"[^\w]", "_", name)
                snaps.append(
                    f"$ws = $wb.Worksheets.Item('{name}'); $ws.Activate(); "
                    f"$excel.ActiveWindow.Zoom = 100; "
                    f"Snap '{out_dir}\\{i:02d}_{safe}_r{r:03d}.png'"
                )
            _run_ps(PS_PASS2
                    .replace("__XLSX__", str(copy))
                    .replace("__SNAPS__", "\n    ".join(snaps)), timeout)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    return sorted(out_dir.glob("*.png"))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", nargs="?", default=str(DEFAULT_XLSX))
    ap.add_argument("--out", required=True)
    ap.add_argument("--sheets", default="")
    a = ap.parse_args()
    shots = screenshot(Path(a.xlsx), Path(a.out),
                       [s for s in a.sheets.split(",") if s])
    for p in shots:
        print(p)
    sys.exit(0 if shots else 1)
