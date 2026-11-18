from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DARK = "2F3E46"
CORN = "F2C14E"
LIGHT = "FFF6DD"
GREY = "F3F4F6"
RED_FILL = PatternFill("solid", fgColor="F8D7DA")
AMBER_FILL = PatternFill("solid", fgColor="FFE8B3")
GREEN_FILL = PatternFill("solid", fgColor="D9F2E3")
BLUE_FILL = PatternFill("solid", fgColor="DCEBFA")
INPUT_FILL = PatternFill("solid", fgColor="FFFBEA")
HEAD_FILL = PatternFill("solid", fgColor=DARK)
CORN_FILL = PatternFill("solid", fgColor=CORN)
GREY_FILL = PatternFill("solid", fgColor=GREY)
thin = Side(style="thin", color="D1D5DB")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(cjk=False, **kw):
    name = "Microsoft JhengHei" if cjk else "Calibri"
    return Font(name=name, **kw)


def title(ws, text, sub=None, cjk=False, width_cols=8):
    ws["A1"] = text
    ws["A1"].font = font(cjk, size=18, bold=True, color=DARK)
    ws.row_dimensions[1].height = 30
    if sub:
        ws["A2"] = sub
        ws["A2"].font = font(cjk, size=10, italic=True, color="6B7280")


def header_row(ws, row, headers, cjk=False, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = font(cjk, bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def landscape(ws, fit_width=True):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    if fit_width:
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4


def howto(ws, title_text, lines, cjk=False):
    ws["A1"] = title_text
    ws["A1"].font = font(cjk, size=18, bold=True, color=DARK)
    ws.column_dimensions["A"].width = 110
    r = 3
    for line in lines:
        c = ws.cell(row=r, column=1, value=line)
        if line.startswith("## "):
            c.value = line[3:]
            c.font = font(cjk, size=12, bold=True, color=DARK)
            c.fill = CORN_FILL
        else:
            c.font = font(cjk, size=11)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    landscape(ws)
