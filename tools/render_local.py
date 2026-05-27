"""Render an .xlsx file as PNG images via Excel COM + PyMuPDF.

Usage: python render.py <path-to-xlsx>

Output: JSON to stdout
{
  "xlsx": "<absolute path>",
  "sheets": [{"name": "...", "png": "...", "hidden": false}, ...]
}

Side effects: PNG files saved under <cwd>/.xlsx-review/<basename>/
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path


def slugify(name: str) -> str:
    """Sanitize sheet name for use in a filename."""
    s = re.sub(r"[^\w\-]+", "_", name, flags=re.UNICODE)
    return s.strip("_") or "sheet"


def render(xlsx_path: Path, out_dir: Path) -> dict:
    try:
        import win32com.client  # type: ignore
    except ImportError:
        raise RuntimeError("pywin32 not installed. Run: pip install pywin32")

    try:
        import fitz  # type: ignore  # PyMuPDF
    except ImportError:
        raise RuntimeError("PyMuPDF not installed. Run: pip install PyMuPDF")

    out_dir.mkdir(parents=True, exist_ok=True)

    # Excel can't open the same file twice. Copy to a temp location to be safe.
    tmp_root = Path(tempfile.mkdtemp(prefix="xlsx_review_"))
    work_xlsx = tmp_root / xlsx_path.name
    shutil.copy2(xlsx_path, work_xlsx)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    result_sheets = []
    try:
        wb = excel.Workbooks.Open(str(work_xlsx), ReadOnly=True, UpdateLinks=0)
        try:
            for idx, ws in enumerate(wb.Worksheets, start=1):
                name = ws.Name
                hidden = ws.Visible != -1  # xlSheetVisible = -1

                if hidden:
                    result_sheets.append({"name": name, "png": None, "hidden": True})
                    continue

                used = ws.UsedRange
                if used.Rows.Count == 1 and used.Columns.Count == 1 and not used.Value:
                    # Empty sheet — render anyway so reviewer sees emptiness.
                    pass

                # Set print area to used range if none set.
                try:
                    if not ws.PageSetup.PrintArea:
                        ws.PageSetup.PrintArea = used.Address
                except Exception:
                    pass

                # Page setup: respektera filens val om de är uttryckligt satta.
                ps = ws.PageSetup
                try:
                    file_fit_w = ps.FitToPagesWide
                    file_fit_h = ps.FitToPagesTall
                except Exception:
                    file_fit_w, file_fit_h = 1, False

                if file_fit_w == 1 and not file_fit_h:
                    ps.Zoom = False
                    ps.FitToPagesWide = 1
                    ps.FitToPagesTall = False
                    try:
                        width = used.Columns.Count
                        height = used.Rows.Count
                        ps.Orientation = 2 if width > height * 0.5 else 1
                    except Exception:
                        ps.Orientation = 2

                pdf_path = tmp_root / f"sheet-{idx:02d}.pdf"
                # xlTypePDF = 0
                ws.ExportAsFixedFormat(Type=0, Filename=str(pdf_path),
                                       Quality=0, IncludeDocProperties=False,
                                       IgnorePrintAreas=False, OpenAfterPublish=False)

                # Rasterize each PDF page; combine vertically if multiple pages.
                png_name = f"sheet-{idx:02d}-{slugify(name)}.png"
                png_path = out_dir / png_name
                _pdf_to_png(pdf_path, png_path, dpi=150)

                result_sheets.append({
                    "name": name,
                    "png": str(png_path.resolve()),
                    "hidden": False,
                })
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        # Try to remove temp dir; ignore errors (Excel sometimes holds locks briefly).
        try:
            shutil.rmtree(tmp_root, ignore_errors=True)
        except Exception:
            pass

    return {
        "xlsx": str(xlsx_path.resolve()),
        "sheets": result_sheets,
    }


def _pdf_to_png(pdf_path: Path, png_path: Path, dpi: int = 150) -> None:
    """Rasterize PDF to PNG. If multiple pages, stack vertically."""
    import fitz  # type: ignore
    from io import BytesIO

    doc = fitz.open(str(pdf_path))
    try:
        mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pages = []
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            pages.append(pix)

        if not pages:
            return

        if len(pages) == 1:
            pages[0].save(str(png_path))
            return

        # Stack vertically: requires PIL since fitz can't compose pixmaps natively.
        try:
            from PIL import Image
        except ImportError:
            # Fall back to first page only with a warning marker in filename.
            pages[0].save(str(png_path.with_name(png_path.stem + "_p1of" + str(len(pages)) + ".png")))
            return

        images = []
        for pix in pages:
            img = Image.open(BytesIO(pix.tobytes("png")))
            images.append(img)

        max_w = max(img.width for img in images)
        total_h = sum(img.height for img in images)
        combined = Image.new("RGB", (max_w, total_h), "white")
        y = 0
        for img in images:
            combined.paste(img, (0, y))
            y += img.height
        combined.save(str(png_path))
    finally:
        doc.close()


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print("Usage: python render.py <path-to-xlsx>", file=sys.stderr)
        return 2

    xlsx_path = Path(argv[1]).resolve()
    if not xlsx_path.exists():
        print(f"Error: file not found: {xlsx_path}", file=sys.stderr)
        return 1
    if xlsx_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        print(f"Error: not an xlsx/xlsm file: {xlsx_path}", file=sys.stderr)
        return 1

    out_dir = Path.cwd() / ".xlsx-review" / xlsx_path.stem

    try:
        result = render(xlsx_path, out_dir)
    except RuntimeError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
