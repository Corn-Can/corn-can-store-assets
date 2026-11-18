"""Recalculate a workbook in LibreOffice and report formula errors plus selected cells."""
import subprocess, sys, os, shutil, tempfile
from openpyxl import load_workbook

src = sys.argv[1]
show = sys.argv[2:]  # e.g. Sheet!A1:B5
tmp = tempfile.mkdtemp()
subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                "--outdir", tmp, src], check=True, capture_output=True, timeout=180)
out = os.path.join(tmp, os.path.basename(src))
wb = load_workbook(out, data_only=True)
errs = []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#") and c.value.rstrip("!?/0").upper() in (
                    "#REF", "#VALUE", "#NAME", "#DIV", "#N/A", "#NUM", "#NULL") or (isinstance(c.value, str) and c.value[:4] in ("#REF", "#VAL", "#NAM", "#DIV", "#N/A", "#NUM", "Err:")):
                errs.append(f"{ws.title}!{c.coordinate}={c.value}")
print("ERRORS:", len(errs), errs[:30])
for spec in show:
    sh, rng = spec.split("!")
    ws = wb[sh]
    for row in ws[rng]:
        print(sh, [c.value if not hasattr(c.value, "date") else str(c.value)[:10] for c in row])
