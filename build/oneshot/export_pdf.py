"""Trogen utskriftsexport via Excel COM ExportAsFixedFormat (ingen page-setup-manipulation).

Per flik: sidantal (PageSetup.Pages.Count) + egen PDF under .print/.
Hela boken: build/oneshot/print_preview.pdf.

NIGHTRUN-läxa: render_local/xlsx-review ljuger på flikar med fitToWidth/Height —
detta är den trogna verifieringen.
"""
from __future__ import annotations
import io
import subprocess
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parent.parent.parent
XLSX = (Path(sys.argv[1]) if len(sys.argv) > 1
        else ROOT / "build" / "oneshot" / "Investeringskalkyl_v2.xlsx").resolve()
OUTDIR = XLSX.parent / ".print"
OUTDIR.mkdir(exist_ok=True)
BOOK_PDF = XLSX.parent / "print_preview.pdf"

PS = f"""
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $wb = $excel.Workbooks.Open('{XLSX}')
    foreach ($ws in $wb.Worksheets) {{
        if ($ws.Visible -eq -1) {{
            $pages = $ws.PageSetup.Pages.Count
            $safe = $ws.Name -replace '[^a-zA-Z0-9]', '_'
            $pdf = '{OUTDIR}' + '\\' + $safe + '.pdf'
            $ws.ExportAsFixedFormat(0, $pdf)
            Write-Output ("SHEET|" + $ws.Name + "|" + $pages)
        }}
    }}
    $wb.ExportAsFixedFormat(0, '{BOOK_PDF}')
    $wb.Close($false)
}} finally {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
"""

res = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", PS],
                     timeout=300, capture_output=True, text=True)
if res.returncode != 0:
    print(res.stdout)
    print(res.stderr)
    sys.exit(1)
total = 0
for line in res.stdout.splitlines():
    if line.startswith("SHEET|"):
        _, name, pages = line.split("|")
        total += int(pages)
        flag = " ⚠" if int(pages) > 2 else ""
        print(f"  {name:22s} {pages:>3} sidor{flag}")
print(f"  TOTALT                {total:>3} sidor -> {BOOK_PDF.name}")
