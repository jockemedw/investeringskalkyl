"""Skärmdumpar av varje fliks faktiska öppningsvy — sanningen för design-polish.

Öppnar xlsx i Excel COM med SYNLIGT maximerat fönster (ReadOnly — krockar inte
med en redan öppen instans), sätter 100 % zoom, scrollar till öppningsläget och
tar en skärmdump per flik. Flikar med innehåll under folden får extra scrollade
vyer. COM via PowerShell (mönster: tools/recalc.py) — ingen pywin32.

Kör:  python build/oneshot/screenshot_sheets.py [xlsx] --out DIR [--sheets A,B]
"""
from __future__ import annotations
import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
DEFAULT_XLSX = ROOT / "build" / "oneshot" / "Investeringskalkyl_v2.xlsx"

PS_TEMPLATE = r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public class DPI { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h); }'
[DPI]::SetProcessDPIAware() | Out-Null

function Snap($path) {
    Start-Sleep -Milliseconds 500
    $b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($b.Location, [System.Drawing.Point]::Empty, $b.Size)
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Output "SNAP $path"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open('__XLSX__', 0, $true)
    $excel.WindowState = -4137  # xlMaximized
    [DPI]::SetForegroundWindow($excel.Hwnd) | Out-Null
    $only = @(__SHEETS__)
    $i = 0
    foreach ($ws in $wb.Worksheets) {
        $i++
        if ($ws.Visible -ne -1) { continue }
        if ($only.Count -gt 0 -and $only -notcontains $ws.Name) { continue }
        $ws.Activate()
        $win = $excel.ActiveWindow
        $win.Zoom = 100
        $win.ScrollRow = 1
        $win.ScrollColumn = 1
        $safe = $ws.Name -replace '[^\w]', '_'
        Snap "__OUT__\$($i.ToString('00'))_$($safe)_r001.png"
        # Scrollade vyer genom hela använda området
        $lastRow = $ws.UsedRange.Rows($ws.UsedRange.Rows.Count).Row
        $visRows = $win.VisibleRange.Rows.Count
        $r = 1 + $visRows
        while ($r -le $lastRow -and $visRows -gt 3) {
            $win.ScrollRow = $r
            Snap "__OUT__\$($i.ToString('00'))_$($safe)_r$($r.ToString('000')).png"
            $r += $visRows
        }
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""


def screenshot(xlsx: Path, out_dir: Path, sheets: list[str] | None = None,
               timeout: int = 300) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    ps_sheets = ", ".join(f"'{s}'" for s in (sheets or []))
    script = (PS_TEMPLATE
              .replace("__XLSX__", str(xlsx.resolve()))
              .replace("__OUT__", str(out_dir.resolve()))
              .replace("__SHEETS__", ps_sheets))
    res = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
        timeout=timeout, capture_output=True, text=True,
    )
    if res.returncode != 0:
        raise RuntimeError(f"screenshot fail:\n{res.stdout}\n{res.stderr}")
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
