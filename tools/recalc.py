"""Tvinga formelutvärdering så cached values uppdateras i xlsx.

openpyxl bevarar formeltext men nollar cached values vid spara.
För regressionstest mot Resultat!D14 etc. krävs recalc av en av:
  1. LibreOffice headless (--convert-to xlsx)
  2. Excel COM via PowerShell (Windows + Excel installerat)

Probar i ordning. Använd recalc(path) för automatiskt val, eller
recalc_libreoffice() / recalc_excel() för explicit.
"""
from __future__ import annotations
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

LIBREOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]
EXCEL_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
]


def find_soffice() -> str | None:
    for env in ("SOFFICE", "LIBREOFFICE"):
        if (p := os.environ.get(env)) and Path(p).exists():
            return p
    if (p := shutil.which("soffice")):
        return p
    if (p := shutil.which("libreoffice")):
        return p
    for cand in LIBREOFFICE_CANDIDATES:
        if Path(cand).exists():
            return cand
    return None


def find_excel() -> str | None:
    for cand in EXCEL_CANDIDATES:
        if Path(cand).exists():
            return cand
    return None


def recalc_libreoffice(xlsx: Path | str, timeout: int = 60) -> Path:
    """Recalc in-place via LibreOffice headless."""
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice ej hittad")
    xlsx = Path(xlsx).resolve()
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        res = subprocess.run(
            [soffice, "--headless", "--calc", "--convert-to", "xlsx",
             "--outdir", str(td_path), str(xlsx)],
            timeout=timeout, capture_output=True, text=True,
        )
        if res.returncode != 0:
            raise RuntimeError(f"soffice fail:\n{res.stderr}")
        produced = td_path / xlsx.name
        if not produced.exists():
            raise RuntimeError(f"soffice producerade ingen fil i {td_path}")
        shutil.copy2(produced, xlsx)
    return xlsx


def recalc_excel(xlsx: Path | str, timeout: int = 60) -> Path:
    """Recalc in-place via Excel COM (PowerShell-anrop, ingen pywin32 krävs)."""
    if not find_excel():
        raise RuntimeError("Excel ej hittad")
    xlsx = Path(xlsx).resolve()
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $wb = $excel.Workbooks.Open('{xlsx}')
    $excel.CalculateFullRebuild()
    $wb.Save()
    $wb.Close($false)
}} finally {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
"""
    res = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", ps_script],
        timeout=timeout, capture_output=True, text=True,
    )
    if res.returncode != 0:
        raise RuntimeError(f"Excel COM fail:\n{res.stdout}\n{res.stderr}")
    return xlsx


def recalc(xlsx: Path | str, timeout: int = 60) -> Path:
    """Auto: prova LibreOffice → Excel. Returnerar samma path (in-place recalc)."""
    last_err: Exception | None = None
    for fn, label in [(recalc_libreoffice, "LibreOffice"), (recalc_excel, "Excel")]:
        try:
            return fn(xlsx, timeout=timeout)
        except RuntimeError as e:
            last_err = e
            continue
    raise RuntimeError(
        "Varken LibreOffice eller Excel kunde användas för recalc.\n"
        f"Senaste fel: {last_err}\n"
        "Lös genom att installera LibreOffice eller köra på en maskin med Excel."
    )


if __name__ == "__main__":
    import sys
    so = find_soffice()
    ex = find_excel()
    print(f"LibreOffice: {so or '(ej hittad)'}")
    print(f"Excel:       {ex or '(ej hittad)'}")
    if not (so or ex):
        sys.exit(1)
