"""Kolla iter9.xlsx för features utanför cellspec: validering, CF, names, kommentarer."""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import openpyxl

wb = openpyxl.load_workbook("build/iter9.xlsx")
print("defined names:", dict(wb.defined_names))
for ws in wb.worksheets:
    dv = ws.data_validations.dataValidation
    cf = ws.conditional_formatting
    cfs = list(cf)
    comments = [c for row in ws.iter_rows() for c in row if c.comment]
    extras = []
    if dv:
        extras.append(f"validations={[(d.type, d.formula1, str(d.sqref)) for d in dv]}")
    if cfs:
        extras.append(f"condfmt={[(str(r.sqref), [type(x).__name__ for x in r.rules]) for r in cfs]}")
    if comments:
        extras.append(f"comments={[(c.coordinate, c.comment.text[:60]) for c in comments]}")
    if extras:
        print(f"--- {ws.title}")
        for e in extras:
            print("   ", e)
