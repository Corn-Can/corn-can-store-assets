import sys, datetime as dt
sys.path.insert(0, "/home/user/corn-can-store-assets/tools")
from common import *
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule, CellIsRule
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font

OUT = sys.argv[1]
D = dt.date


def grid(ws, r1, r2, c1, c2, cjk=False, input_cols=(), fmt=None):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.font = font(cjk, size=10)
            cell.alignment = Alignment(vertical="center")
            if c in input_cols:
                cell.fill = INPUT_FILL
            if fmt and c in fmt:
                cell.number_format = fmt[c]


def dv_list(ws, formula, rng):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


def kpi(ws, r, c, label, formula, fmtc="0", cjk=False):
    a = ws.cell(row=r, column=c, value=label)
    a.font = font(cjk, size=10, bold=True, color="FFFFFF")
    a.fill = HEAD_FILL
    a.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    b = ws.cell(row=r + 1, column=c, value=formula)
    b.font = font(cjk, size=16, bold=True, color=DARK)
    b.fill = PatternFill("solid", fgColor=LIGHT)
    b.alignment = Alignment(horizontal="center", vertical="center")
    b.number_format = fmtc
    a.border = b.border = BORDER
    ws.row_dimensions[r].height = 30
    ws.row_dimensions[r + 1].height = 32


def setup_rows(ws, r0, rows, cjk=False):
    for i, (k, v, f) in enumerate(rows):
        a = ws.cell(row=r0 + i, column=1, value=k); a.font = font(cjk, bold=True); a.border = BORDER
        b = ws.cell(row=r0 + i, column=2, value=v); b.fill = INPUT_FILL; b.border = BORDER; b.number_format = f
        b.font = font(cjk, size=12, bold=True)


def list_col(ws, row, col, head, items, n, cjk=False, inp=True):
    header_row(ws, row, [head], cjk=cjk, start_col=col)
    for i in range(n):
        c = ws.cell(row=row + 1 + i, column=col, value=items[i] if i < len(items) else None)
        c.border = BORDER; c.font = font(cjk, size=10)
        if inp:
            c.fill = INPUT_FILL


# =====================================================================
# 1) Gantt Chart Project Planner
# =====================================================================
def gantt():
    wb = Workbook()
    hw = wb.active; hw.title = "How to Use"
    st = wb.create_sheet("Setup")
    pp = wb.create_sheet("Project Plan")
    db = wb.create_sheet("Dashboard")
    N = 100
    R1, R2 = 5, 4 + N
    NT = 56  # timeline columns
    TC0 = 14  # column N

    howto(hw, "Gantt Chart Project Planner – How to Use", [
        "## What's inside (4 sheets)",
        "How to Use – this page.",
        "Setup – project name, timeline start date, days per timeline column (1 = daily, 7 = weekly), and the Owner and Phase lists used by the drop-downs.",
        "Project Plan – your task list (up to 100 tasks) with an automatic Gantt chart on the right. Bars are drawn with conditional formatting: yellow = planned, green = completed part (based on % done), red column = today.",
        "Dashboard – tasks done, duration-weighted project progress, overdue tasks, tasks due in the next 7 days, project start/end dates, plus progress by phase and workload by owner with a chart.",
        "## Quick start",
        "1. On Setup, type your project name and the first date you want the timeline to show. Pick 1 day per column for short projects or 7 for long ones.",
        "2. On Project Plan, replace the example tasks (or delete them). Light-yellow cells are inputs; white cells are formulas – please don't type over them.",
        "3. For each task enter a Start date and a Duration in calendar days. End date = Start + Duration − 1.",
        "4. Dependencies (optional): type the ID of another task in 'Depends on'. The task then starts the day after that task ends, and its own Start date cell is ignored. Change one task and everything after it moves.",
        "5. Update '% done' and Status as you go. A task is flagged OVERDUE when its end date has passed and Status is not Done, and DUE SOON when it ends within 7 days.",
        "## Tips",
        "Work days: the 'Work days' column counts Monday–Friday days between start and end, so you can compare with calendar duration.",
        "Milestones: give a task a duration of 1 day and the phase 'Milestone'.",
        "To print: the Project Plan print area covers the task list and the timeline on one landscape page width (File → Print).",
        "Don't make a task depend on itself or create a loop (A depends on B, B depends on A) – Excel will report a circular reference.",
        "## Works with",
        "Microsoft Excel 2010 or later (Windows/Mac), Google Sheets (File → Import) and LibreOffice Calc. No macros, no add-ins.",
        "Made by Corn_Can. For personal and business use; please don't resell the file.",
    ])

    title(st, "Setup", "Type your project details in the yellow cells.")
    setup_rows(st, 4, [("Project name", "Online Store Launch", "@"),
                       ("Project manager", "Alex", "@"),
                       ("Timeline start date", D(2027, 1, 4), "yyyy-mm-dd"),
                       ("Days per timeline column", 1, "0")])
    dv = DataValidation(type="list", formula1='"1,2,3,7,14"', allow_blank=False)
    st.add_data_validation(dv); dv.add("B7")
    list_col(st, 10, 1, "Owners", ["Alex", "Sam", "Jordan", "Priya", "Chris"], 12)
    list_col(st, 10, 3, "Phases", ["Planning", "Design", "Build", "Testing", "Launch", "Milestone"], 12)
    list_col(st, 10, 5, "Statuses", ["Not started", "In progress", "Done", "On hold"], 4, inp=False)
    widths(st, {"A": 26, "B": 22, "C": 18, "D": 3, "E": 16})
    landscape(st)

    # Project Plan
    title(pp, None)
    pp["A1"] = "=Setup!B4&\" – Project Plan\""
    pp["A1"].font = font(size=18, bold=True, color=DARK)
    pp["A2"] = '="Manager: "&Setup!B5&"   ·   Today: "&TEXT(TODAY(),"yyyy-mm-dd")&"   ·   Yellow bar = planned, green = done, red column = today"'
    pp["A2"].font = font(size=10, italic=True, color="6B7280")
    hdr = ["ID", "Phase", "Task", "Owner", "Depends on (ID)", "Start date", "Duration (days)", "Start",
           "End", "Work days", "% done", "Status", "Flag"]
    header_row(pp, 4, hdr)
    ex = [
        (1, "Planning", "Define goals & budget", "Alex", None, D(2027, 1, 4), 5, 1.0, "Done"),
        (2, "Planning", "Choose platform & suppliers", "Sam", 1, None, 7, 0.6, "In progress"),
        (3, "Design", "Logo, brand colors, photos", "Priya", 1, None, 10, 0.3, "In progress"),
        (4, "Build", "Set up store & product pages", "Jordan", 2, None, 12, 0, "Not started"),
        (5, "Build", "Payments, shipping & policies", "Sam", 4, None, 5, 0, "Not started"),
        (6, "Testing", "Test orders & fix issues", "Chris", 5, None, 6, 0, "Not started"),
        (7, "Launch", "Email list & social posts", "Priya", 3, None, 14, 0, "Not started"),
        (8, "Milestone", "Store goes live", "Alex", 6, None, 1, 0, "Not started"),
    ]
    for i, e in enumerate(ex):
        r = R1 + i
        for col, v in zip((1, 2, 3, 4, 5, 6, 7, 11, 12), e):
            pp.cell(row=r, column=col, value=v)
    for r in range(R1, R2 + 1):
        pp.cell(row=r, column=8, value=f'=IF(C{r}="","",IF(E{r}<>"",IFERROR(INDEX($I${R1}:$I${R2},MATCH(E{r},$A${R1}:$A${R2},0))+1,F{r}),F{r}))')
        pp.cell(row=r, column=9, value=f'=IF(OR(C{r}="",NOT(ISNUMBER(H{r})),NOT(ISNUMBER(G{r}))),"",H{r}+G{r}-1)')
        pp.cell(row=r, column=10, value=f'=IF(ISNUMBER(I{r}),NETWORKDAYS(H{r},I{r}),"")')
        pp.cell(row=r, column=13, value=f'=IF(OR(C{r}="",NOT(ISNUMBER(I{r}))),"",IF(L{r}="Done","✓",IF(I{r}<TODAY(),"OVERDUE",IF(I{r}-TODAY()<=7,"DUE SOON",""))))')
    grid(pp, R1, R2, 1, 13, input_cols=(1, 2, 3, 4, 5, 6, 7, 11, 12),
         fmt={6: "yyyy-mm-dd", 8: "mmm d", 9: "mmm d", 11: "0%"})
    for r in range(R1, R2 + 1):
        pp.cell(row=r, column=13).font = font(size=9, bold=True)
    dv_list(pp, "=Setup!$C$11:$C$22", f"B{R1}:B{R2}")
    dv_list(pp, "=Setup!$A$11:$A$22", f"D{R1}:D{R2}")
    dv_list(pp, '"Not started,In progress,Done,On hold"', f"L{R1}:L{R2}")
    dvp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                         showErrorMessage=True, error="Enter a percentage from 0% to 100%.")
    pp.add_data_validation(dvp); dvp.add(f"K{R1}:K{R2}")
    pp.conditional_formatting.add(f"M{R1}:M{R2}", FormulaRule(formula=[f'M{R1}="OVERDUE"'], fill=RED_FILL, font=Font(color="9B1C1C", bold=True)))
    pp.conditional_formatting.add(f"M{R1}:M{R2}", FormulaRule(formula=[f'M{R1}="DUE SOON"'], fill=AMBER_FILL))
    pp.conditional_formatting.add(f"M{R1}:M{R2}", FormulaRule(formula=[f'M{R1}="✓"'], fill=GREEN_FILL))
    pp.conditional_formatting.add(f"K{R1}:K{R2}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="9AD1A8"))

    # timeline
    first, last = CL(TC0), CL(TC0 + NT - 1)
    for k in range(NT):
        c = TC0 + k
        col = CL(c)
        cell = pp.cell(row=4, column=c, value="=Setup!$B$6" if k == 0 else f"={CL(c-1)}4+Setup!$B$7")
        cell.number_format = "d"
        cell.font = font(size=8, bold=True, color="FFFFFF"); cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", text_rotation=0)
        m = pp.cell(row=3, column=c, value=f'=IF(OR({col}4=Setup!$B$6,DAY({col}4)<=Setup!$B$7),TEXT({col}4,"mmm"),"")')
        m.font = font(size=8, bold=True, color=DARK)
        pp.column_dimensions[col].width = 3.2
        for r in range(R1, R2 + 1):
            pp.cell(row=r, column=c).border = Border(left=Side(style="hair", color="E5E7EB"), right=Side(style="hair", color="E5E7EB"),
                                                     top=Side(style="hair", color="E5E7EB"), bottom=Side(style="hair", color="E5E7EB"))
    rng = f"{first}{R1}:{last}{R2}"
    f0 = first
    inbar = f'AND(ISNUMBER($H{R1}),ISNUMBER($I{R1}),{f0}$4<=$I{R1},{f0}$4+Setup!$B$7-1>=$H{R1})'
    pp.conditional_formatting.add(rng, FormulaRule(formula=[f'AND({inbar},{f0}$4<$H{R1}+$G{R1}*N($K{R1}))'], fill=PatternFill("solid", fgColor="4CAF7A"), stopIfTrue=True))
    pp.conditional_formatting.add(rng, FormulaRule(formula=[f'AND({inbar},$M{R1}="OVERDUE")'], fill=PatternFill("solid", fgColor="E57373"), stopIfTrue=True))
    pp.conditional_formatting.add(rng, FormulaRule(formula=[inbar], fill=PatternFill("solid", fgColor="F2C14E"), stopIfTrue=True))
    pp.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(TODAY()>={f0}$4,TODAY()<{f0}$4+Setup!$B$7)'], fill=PatternFill("solid", fgColor="FBD5D5")))
    pp.conditional_formatting.add(f"{first}4:{last}4", FormulaRule(formula=[f'AND(TODAY()>={f0}$4,TODAY()<{f0}$4+Setup!$B$7)'], fill=PatternFill("solid", fgColor="D64545")))
    pp.conditional_formatting.add(f"{first}4:{last}4", FormulaRule(formula=[f'WEEKDAY({f0}$4,2)>5'], fill=PatternFill("solid", fgColor="6B7C85")))
    widths(pp, {"A": 5, "B": 11, "C": 30, "D": 10, "E": 9, "F": 11, "G": 9, "H": 8, "I": 8, "J": 7, "K": 8, "L": 11, "M": 10})
    pp.freeze_panes = "D5"
    landscape(pp); pp.print_area = f"A1:{last}30"

    # Dashboard
    P = "'Project Plan'!"
    C_, G_, H_, I_, K_, L_, M_, B_, D_ = (f"{P}${c}${R1}:${c}${R2}" for c in "CGHIKLMBD")
    title(db, "Project Dashboard")
    db["A2"] = '=Setup!B4&"   ·   Manager: "&Setup!B5&"   ·   Today: "&TEXT(TODAY(),"yyyy-mm-dd")'
    db["A2"].font = font(size=11, italic=True, color="6B7280")
    kpi(db, 4, 1, "Tasks", f'=COUNTIF({C_},"?*")')
    kpi(db, 4, 2, "Done", f'=COUNTIFS({C_},"?*",{L_},"Done")')
    kpi(db, 4, 3, "Project progress", f'=IFERROR(SUMPRODUCT({G_},{K_})/SUMIFS({G_},{C_},"?*"),0)', "0%")
    kpi(db, 4, 4, "Overdue", f'=COUNTIF({M_},"OVERDUE")')
    kpi(db, 4, 5, "Due in next 7 days", f'=COUNTIF({M_},"DUE SOON")')
    kpi(db, 7, 1, "Project start", f'=IF(COUNT({H_})=0,"",MIN({H_}))', "yyyy-mm-dd")
    kpi(db, 7, 2, "Project end", f'=IF(COUNT({I_})=0,"",MAX({I_}))', "yyyy-mm-dd")
    kpi(db, 7, 3, "Total length (days)", '=IF(OR(A8="",B8=""),"",B8-A8+1)')
    kpi(db, 7, 4, "Days until end", '=IF(B8="","",IF(B8-TODAY()<0,"Ended",B8-TODAY()))')
    kpi(db, 7, 5, "On hold", f'=COUNTIF({L_},"On hold")')
    db.conditional_formatting.add("C5", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F2C14E"))
    db.conditional_formatting.add("D5", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
    db.conditional_formatting.add("E5", CellIsRule(operator="greaterThan", formula=["0"], fill=AMBER_FILL))

    header_row(db, 10, ["Phase", "Tasks", "Done", "Progress", "Open tasks"])
    for i in range(12):
        r = 11 + i
        db.cell(row=r, column=1, value=f'=IF(Setup!C{11+i}="","",Setup!C{11+i})')
        db.cell(row=r, column=2, value=f'=IF(A{r}="","",COUNTIFS({B_},A{r},{C_},"?*"))')
        db.cell(row=r, column=3, value=f'=IF(A{r}="","",COUNTIFS({B_},A{r},{C_},"?*",{L_},"Done"))')
        db.cell(row=r, column=4, value=f'=IF(OR(A{r}="",N(B{r})=0),"",IFERROR(SUMPRODUCT(--({B_}=A{r}),{G_},{K_})/SUMIFS({G_},{B_},A{r},{C_},"?*"),0))')
        db.cell(row=r, column=5, value=f'=IF(A{r}="","",B{r}-C{r})')
    grid(db, 11, 22, 1, 5, fmt={4: "0%"})
    db.conditional_formatting.add("D11:D22", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="9AD1A8"))

    header_row(db, 10, ["Owner", "Tasks", "Open", "Overdue", "Days assigned"], start_col=7)
    for i in range(12):
        r = 11 + i
        db.cell(row=r, column=7, value=f'=IF(Setup!A{11+i}="","",Setup!A{11+i})')
        db.cell(row=r, column=8, value=f'=IF(G{r}="","",COUNTIFS({D_},G{r},{C_},"?*"))')
        db.cell(row=r, column=9, value=f'=IF(G{r}="","",COUNTIFS({D_},G{r},{C_},"?*",{L_},"<>Done"))')
        db.cell(row=r, column=10, value=f'=IF(G{r}="","",COUNTIFS({D_},G{r},{M_},"OVERDUE"))')
        db.cell(row=r, column=11, value=f'=IF(G{r}="","",SUMIFS({G_},{D_},G{r},{C_},"?*"))')
    grid(db, 11, 22, 7, 11)
    db.conditional_formatting.add("J11:J22", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))

    ch = BarChart(); ch.type = "bar"; ch.title = "Progress by phase"; ch.style = 10
    ch.add_data(Reference(db, min_col=4, min_row=10, max_row=16), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=1, min_row=11, max_row=16))
    ch.legend = None; ch.height = 7.5; ch.width = 13
    ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1; ch.y_axis.numFmt = "0%"
    ch.series[0].graphicalProperties.solidFill = "F2C14E"
    db.add_chart(ch, "A25")
    ch2 = BarChart(); ch2.type = "col"; ch2.title = "Days assigned by owner"; ch2.style = 10
    ch2.add_data(Reference(db, min_col=11, min_row=10, max_row=15), titles_from_data=True)
    ch2.set_categories(Reference(db, min_col=7, min_row=11, max_row=15))
    ch2.legend = None; ch2.height = 7.5; ch2.width = 13
    ch2.series[0].graphicalProperties.solidFill = "2F3E46"
    db.add_chart(ch2, "G25")
    widths(db, {"A": 16, "B": 14, "C": 16, "D": 14, "E": 16, "F": 3, "G": 14, "H": 9, "I": 9, "J": 10, "K": 13})
    landscape(db); db.page_setup.fitToHeight = 1; db.print_area = "A1:K41"
    wb.move_sheet("Dashboard", offset=-2)
    wb.active = 2
    wb.save(f"{OUT}/gantt-chart-project-planner.xlsx")


# =====================================================================
# 2) Inventory & Stock Tracker
# =====================================================================
def inventory():
    wb = Workbook()
    hw = wb.active; hw.title = "How to Use"
    st = wb.create_sheet("Setup")
    pr = wb.create_sheet("Products")
    mv = wb.create_sheet("Stock Moves")
    sc = wb.create_sheet("Stock Count")
    db = wb.create_sheet("Dashboard")
    NP, NM = 200, 1000
    P1, P2 = 5, 4 + NP
    M1, M2 = 5, 4 + NM

    howto(hw, "Inventory & Stock Tracker – How to Use", [
        "## What's inside (6 sheets)",
        "How to Use – this page.",
        "Setup – business name, sales year for the dashboard, and the Category and Supplier lists used by the drop-downs.",
        "Products – one row per SKU: cost, selling price, opening stock and reorder point. Margin, units in/out, current stock, stock value and a stock status (OK / LOW – REORDER / OUT OF STOCK) are calculated.",
        "Stock Moves – log every change: Stock in (purchases), Sale, Customer return or Adjustment. Pick the SKU from the drop-down; name and line totals fill in.",
        "Stock Count – for a physical count (e.g. year-end): system quantity is pulled in, you type what you counted, and the sheet shows differences in units and value.",
        "Dashboard – SKUs, units on hand, inventory value at cost and at retail, items to reorder, plus monthly sales, cost of goods and gross profit for the year with a chart, and top sellers.",
        "## Quick start",
        "1. On Setup, type your business name and the year you want the dashboard to report.",
        "2. On Products, replace the example items. Light-yellow cells are inputs; white cells are formulas – please don't type over them. SKUs must be unique.",
        "3. Opening stock = what you have on the day you start the file. After that, record every change on Stock Moves.",
        "4. Quantities on Stock Moves are always positive for Stock in, Sale and Customer return. For Adjustment, type a negative number to remove units (damaged, lost) and a positive number to add.",
        "5. Unit price on a Sale row is optional – leave it empty to use the selling price from Products, or type the actual price (e.g. a discount).",
        "## Physical count",
        "On Stock Count, enter the SKUs you counted and the counted quantity. If there's a difference, add an Adjustment row on Stock Moves so the system matches the shelf.",
        "## Notes",
        "Cost of goods sold uses the unit cost currently on Products (a simple standard-cost method). It's a tracking tool, not accounting or tax advice.",
        "## Works with",
        "Microsoft Excel 2010 or later (Windows/Mac), Google Sheets (File → Import) and LibreOffice Calc. No macros.",
        "Made by Corn_Can. For personal and business use; please don't resell the file.",
    ])

    title(st, "Setup", "Type your details in the yellow cells.")
    setup_rows(st, 4, [("Business name", "Maple & Thread Studio", "@"), ("Sales year (dashboard)", 2027, "0")])
    list_col(st, 8, 1, "Categories", ["Candles", "Soap", "Accessories", "Packaging", "Gift sets"], 15)
    list_col(st, 8, 3, "Suppliers", ["WaxWorks Co.", "Pure Oils Ltd.", "BoxPro", "Local maker"], 15)
    list_col(st, 8, 5, "Move types", ["Stock in", "Sale", "Customer return", "Adjustment"], 4, inp=False)
    widths(st, {"A": 26, "B": 24, "C": 18, "D": 3, "E": 18})
    landscape(st)

    # Products
    title(pr, "Products", "One row per SKU. Yellow = input, white = calculated.")
    hdr = ["SKU", "Product", "Category", "Supplier", "Unit cost", "Selling price", "Margin %", "Opening stock",
           "Reorder point", "Reorder qty", "Stock in", "Sold", "Returns", "Adjust.", "Current stock",
           "Stock value (cost)", "Status"]
    header_row(pr, 4, hdr)
    ex = [("CDL-001", "Vanilla soy candle 8oz", "Candles", "WaxWorks Co.", 6.20, 18, 40, 15, 30),
          ("CDL-002", "Cedar & sage candle 8oz", "Candles", "WaxWorks Co.", 6.50, 19, 25, 15, 30),
          ("SOP-001", "Oatmeal honey soap bar", "Soap", "Pure Oils Ltd.", 2.10, 7.5, 60, 25, 50),
          ("SOP-002", "Lavender soap bar", "Soap", "Pure Oils Ltd.", 2.10, 7.5, 18, 25, 50),
          ("ACC-001", "Wooden candle wick trimmer", "Accessories", "Local maker", 4.00, 12, 10, 5, 10),
          ("GFT-001", "Gift box: candle + 2 soaps", "Gift sets", "BoxPro", 12.50, 34, 8, 5, 10)]
    for i, e in enumerate(ex):
        for j, v in zip((1, 2, 3, 4, 5, 6, 8, 9, 10), e):
            pr.cell(row=P1 + i, column=j, value=v)
    MS = "'Stock Moves'!"
    for r in range(P1, P2 + 1):
        def s(t):
            return f'SUMIFS({MS}$E${M1}:$E${M2},{MS}$B${M1}:$B${M2},$A{r},{MS}$D${M1}:$D${M2},"{t}")'
        pr.cell(row=r, column=7, value=f'=IF(AND(ISNUMBER(E{r}),N(F{r})>0),(F{r}-E{r})/F{r},"")')
        pr.cell(row=r, column=11, value=f'=IF(A{r}="","",{s("Stock in")})')
        pr.cell(row=r, column=12, value=f'=IF(A{r}="","",{s("Sale")})')
        pr.cell(row=r, column=13, value=f'=IF(A{r}="","",{s("Customer return")})')
        pr.cell(row=r, column=14, value=f'=IF(A{r}="","",{s("Adjustment")})')
        pr.cell(row=r, column=15, value=f'=IF(A{r}="","",N(H{r})+K{r}-L{r}+M{r}+N{r})')
        pr.cell(row=r, column=16, value=f'=IF(A{r}="","",O{r}*N(E{r}))')
        pr.cell(row=r, column=18, value=f'=IF(A{r}="","",L{r}+(1000-ROW())/100000)')
        pr.cell(row=r, column=17, value=f'=IF(A{r}="","",IF(O{r}<=0,"OUT OF STOCK",IF(O{r}<=N(I{r}),"LOW – REORDER "&N(J{r}),"OK")))')
    pr["R4"] = "sort key"; pr["R4"].font = font(size=8, color="9CA3AF")
    for r in range(P1, P2 + 1):
        pr.cell(row=r, column=18).font = font(size=8, color="D1D5DB")
    grid(pr, P1, P2, 1, 17, input_cols=(1, 2, 3, 4, 5, 6, 8, 9, 10),
         fmt={5: "#,##0.00", 6: "#,##0.00", 7: "0%", 16: "#,##0.00"})
    dv_list(pr, "=Setup!$A$9:$A$23", f"C{P1}:C{P2}")
    dv_list(pr, "=Setup!$C$9:$C$23", f"D{P1}:D{P2}")
    pr.conditional_formatting.add(f"Q{P1}:Q{P2}", FormulaRule(formula=[f'Q{P1}="OUT OF STOCK"'], fill=RED_FILL, font=Font(color="9B1C1C", bold=True)))
    pr.conditional_formatting.add(f"Q{P1}:Q{P2}", FormulaRule(formula=[f'LEFT(Q{P1},3)="LOW"'], fill=AMBER_FILL, font=Font(bold=True)))
    pr.conditional_formatting.add(f"Q{P1}:Q{P2}", FormulaRule(formula=[f'Q{P1}="OK"'], fill=GREEN_FILL))
    pr.conditional_formatting.add(f"A{P1}:A{P2}", FormulaRule(formula=[f'AND(A{P1}<>"",COUNTIF($A${P1}:$A${P2},A{P1})>1)'], fill=RED_FILL))
    widths(pr, {"A": 10, "B": 28, "C": 13, "D": 15, "E": 9, "F": 9, "G": 8, "H": 9, "I": 9, "J": 9,
                "K": 8, "L": 8, "M": 8, "N": 8, "O": 9, "P": 11, "Q": 18})
    pr.freeze_panes = "C5"; pr.auto_filter.ref = f"A4:Q{P2}"
    landscape(pr); pr.print_area = "A1:Q35"

    # Stock Moves
    title(mv, "Stock Moves", "Log every change. Quantity is positive except for Adjustments that remove stock.")
    header_row(mv, 4, ["Date", "SKU", "Product", "Type", "Quantity", "Unit price (sales, optional)", "Sale amount",
                       "Cost of goods", "Note"])
    PR = "Products!"
    mex = [(D(2027, 1, 3), "CDL-001", "Sale", 3, None, "Online order"),
           (D(2027, 1, 5), "SOP-002", "Sale", 6, 6.5, "Market stall, bundle price"),
           (D(2027, 1, 8), "CDL-002", "Stock in", 24, None, "PO #1042"),
           (D(2027, 1, 12), "GFT-001", "Sale", 4, None, ""),
           (D(2027, 1, 15), "CDL-001", "Customer return", 1, None, "Unopened"),
           (D(2027, 2, 2), "SOP-001", "Adjustment", -2, None, "Damaged in shipping")]
    for i, (d, sku, t, q, p, n) in enumerate(mex):
        r = M1 + i
        for col, v in zip((1, 2, 4, 5, 6, 9), (d, sku, t, q, p, n)):
            mv.cell(row=r, column=col, value=v)
    for r in range(M1, M2 + 1):
        mv.cell(row=r, column=3, value=f'=IF(B{r}="","",IFERROR(INDEX({PR}$B${P1}:$B${P2},MATCH(B{r},{PR}$A${P1}:$A${P2},0)),"SKU not found"))')
        mv.cell(row=r, column=7, value=f'=IF(D{r}<>"Sale","",N(E{r})*IF(ISNUMBER(F{r}),F{r},IFERROR(INDEX({PR}$F${P1}:$F${P2},MATCH(B{r},{PR}$A${P1}:$A${P2},0)),0)))')
        mv.cell(row=r, column=8, value=f'=IF(D{r}<>"Sale","",N(E{r})*IFERROR(INDEX({PR}$E${P1}:$E${P2},MATCH(B{r},{PR}$A${P1}:$A${P2},0)),0))')
    grid(mv, M1, M2, 1, 9, input_cols=(1, 2, 4, 5, 6, 9), fmt={1: "yyyy-mm-dd", 6: "#,##0.00", 7: "#,##0.00", 8: "#,##0.00"})
    dv_list(mv, f"=Products!$A${P1}:$A${P2}", f"B{M1}:B{M2}")
    dv_list(mv, '"Stock in,Sale,Customer return,Adjustment"', f"D{M1}:D{M2}")
    mv.conditional_formatting.add(f"C{M1}:C{M2}", FormulaRule(formula=[f'C{M1}="SKU not found"'], fill=RED_FILL))
    mv.conditional_formatting.add(f"A{M1}:I{M2}", FormulaRule(formula=[f'$D{M1}="Stock in"'], fill=BLUE_FILL))
    widths(mv, {"A": 12, "B": 11, "C": 28, "D": 16, "E": 9, "F": 13, "G": 11, "H": 11, "I": 28})
    mv.freeze_panes = "A5"; mv.auto_filter.ref = f"A4:I{M2}"
    landscape(mv); mv.print_area = "A1:I40"

    # Stock Count
    title(sc, "Stock Count", None)
    sc["A2"] = "Count date:"; sc["B2"] = D(2027, 12, 31); sc["B2"].fill = INPUT_FILL; sc["B2"].number_format = "yyyy-mm-dd"
    sc["A2"].font = font(bold=True)
    header_row(sc, 4, ["SKU", "Product", "System qty", "Counted qty", "Difference", "Unit cost", "Value difference", "Counted by / note"])
    for i, (sku, q) in enumerate([("CDL-001", 38), ("CDL-002", 49), ("SOP-001", 58), ("SOP-002", 12), ("ACC-001", 10), ("GFT-001", 3)]):
        sc.cell(row=5 + i, column=1, value=sku); sc.cell(row=5 + i, column=4, value=q)
    for r in range(5, 205):
        m = f'MATCH(A{r},{PR}$A${P1}:$A${P2},0)'
        sc.cell(row=r, column=2, value=f'=IF(A{r}="","",IFERROR(INDEX({PR}$B${P1}:$B${P2},{m}),"SKU not found"))')
        sc.cell(row=r, column=3, value=f'=IF(A{r}="","",IFERROR(INDEX({PR}$O${P1}:$O${P2},{m}),""))')
        sc.cell(row=r, column=5, value=f'=IF(OR(A{r}="",NOT(ISNUMBER(D{r}))),"",D{r}-N(C{r}))')
        sc.cell(row=r, column=6, value=f'=IF(A{r}="","",IFERROR(INDEX({PR}$E${P1}:$E${P2},{m}),""))')
        sc.cell(row=r, column=7, value=f'=IF(E{r}="","",E{r}*N(F{r}))')
    grid(sc, 5, 204, 1, 8, input_cols=(1, 4, 8), fmt={6: "#,##0.00", 7: "#,##0.00"})
    dv_list(sc, f"=Products!$A${P1}:$A${P2}", "A5:A204")
    sc.conditional_formatting.add("E5:G204", CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
    sc.conditional_formatting.add("E5:G204", CellIsRule(operator="greaterThan", formula=["0"], fill=BLUE_FILL))
    sc["J4"] = "Lines counted"; sc["K4"] = "=COUNT(D5:D204)"
    sc["J5"] = "Lines with difference"; sc["K5"] = '=COUNTIF(E5:E204,"<>0")-COUNTBLANK(E5:E204)'
    sc["J6"] = "Net unit difference"; sc["K6"] = "=SUM(E5:E204)"
    sc["J7"] = "Net value difference"; sc["K7"] = "=SUM(G5:G204)"
    grid(sc, 4, 7, 10, 11, fmt={11: "#,##0.00"})
    sc["K4"].number_format = sc["K5"].number_format = sc["K6"].number_format = "0"
    for c in ("J4", "J5", "J6", "J7"):
        sc[c].font = font(bold=True); sc[c].fill = PatternFill("solid", fgColor=LIGHT)
    widths(sc, {"A": 11, "B": 28, "C": 11, "D": 12, "E": 11, "F": 10, "G": 14, "H": 24, "I": 3, "J": 22, "K": 12})
    sc.freeze_panes = "A5"
    landscape(sc); sc.print_area = "A1:K40"

    # Dashboard
    title(db, "Inventory Dashboard")
    db["A2"] = '=Setup!B4&"   ·   Sales year: "&Setup!B5&"   ·   Today: "&TEXT(TODAY(),"yyyy-mm-dd")'
    db["A2"].font = font(size=11, italic=True, color="6B7280")
    A_, E_, F_, O_, P_, Q_ = (f"{PR}${c}${P1}:${c}${P2}" for c in "AEFOPQ")
    kpi(db, 4, 1, "SKUs", f'=COUNTIF({A_},"?*")')
    kpi(db, 4, 2, "Units on hand", f"=SUM({O_})", "#,##0")
    kpi(db, 4, 3, "Stock value (cost)", f"=SUM({P_})", "#,##0.00")
    kpi(db, 4, 4, "Stock value (retail)", f"=SUMPRODUCT({O_},{F_})", "#,##0.00")
    kpi(db, 4, 5, "Low – reorder", f'=COUNTIF({Q_},"LOW*")')
    kpi(db, 4, 6, "Out of stock", f'=COUNTIF({Q_},"OUT OF STOCK")')
    db.conditional_formatting.add("E5", CellIsRule(operator="greaterThan", formula=["0"], fill=AMBER_FILL))
    db.conditional_formatting.add("F5", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))

    Y = "Setup!$B$5"
    MA, MD, ME, MG, MH = (f"{MS}${c}${M1}:${c}${M2}" for c in "ADEGH")
    header_row(db, 8, ["Month", "Units sold", "Sales", "Cost of goods", "Gross profit", "Margin"])
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for i, mname in enumerate(months):
        r = 9 + i
        crit = f'{MD},"Sale",{MA},">="&DATE({Y},{i+1},1),{MA},"<"&DATE({Y},{i+2},1)' if i < 11 else \
            f'{MD},"Sale",{MA},">="&DATE({Y},12,1),{MA},"<"&DATE({Y}+1,1,1)'
        db.cell(row=r, column=1, value=mname)
        db.cell(row=r, column=2, value=f"=SUMIFS({ME},{crit})")
        db.cell(row=r, column=3, value=f"=SUMIFS({MG},{crit})")
        db.cell(row=r, column=4, value=f"=SUMIFS({MH},{crit})")
        db.cell(row=r, column=5, value=f"=C{r}-D{r}")
        db.cell(row=r, column=6, value=f'=IF(C{r}=0,"",E{r}/C{r})')
    db["A21"] = "Total"
    for c in "BCDE":
        db[f"{c}21"] = f"=SUM({c}9:{c}20)"
    db["F21"] = '=IF(C21=0,"",E21/C21)'
    grid(db, 9, 21, 1, 6, fmt={2: "#,##0", 3: "#,##0.00", 4: "#,##0.00", 5: "#,##0.00", 6: "0%"})
    for c in range(1, 7):
        db.cell(row=21, column=c).font = font(bold=True)

    # Top sellers (all-time units sold) - helper with tie-breaker
    header_row(db, 8, ["Rank", "SKU", "Product", "Units sold", "Status"], start_col=8)
    L_ = f"{PR}$L${P1}:$L${P2}"
    R_ = f"{PR}$R${P1}:$R${P2}"
    for i in range(10):
        r = 9 + i
        db.cell(row=r, column=8, value=i + 1)
        m = f'MATCH(LARGE({R_},{i+1}),{R_},0)'
        db.cell(row=r, column=9, value=f'=IF(COUNT({R_})<{i+1},"",IF(INDEX({L_},{m})<1,"",INDEX({A_},{m})))')
        db.cell(row=r, column=10, value=f'=IF(I{r}="","",INDEX({PR}$B${P1}:$B${P2},{m}))')
        db.cell(row=r, column=11, value=f'=IF(I{r}="","",INDEX({L_},{m}))')
        db.cell(row=r, column=12, value=f'=IF(I{r}="","",INDEX({Q_},{m}))')
    grid(db, 9, 18, 8, 12)
    db.conditional_formatting.add("K9:K18", DataBarRule(start_type="num", start_value=0, end_type="max", color="F2C14E"))

    ch = BarChart(); ch.type = "col"; ch.title = "Sales and gross profit by month"; ch.style = 10
    ch.add_data(Reference(db, min_col=3, min_row=8, max_row=20), titles_from_data=True)
    ch.add_data(Reference(db, min_col=5, min_row=8, max_row=20), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=1, min_row=9, max_row=20))
    ch.height = 7.5; ch.width = 17
    ch.series[0].graphicalProperties.solidFill = "F2C14E"
    ch.series[1].graphicalProperties.solidFill = "2F3E46"
    db.add_chart(ch, "A23")
    widths(db, {"A": 14, "B": 14, "C": 16, "D": 16, "E": 16, "F": 14, "G": 3, "H": 6, "I": 10, "J": 28, "K": 10, "L": 18, "N": 8})
    landscape(db); db.page_setup.fitToHeight = 1; db.print_area = "A1:L39"
    wb.move_sheet("Dashboard", offset=-4)
    wb.active = 1
    wb.save(f"{OUT}/inventory-stock-tracker.xlsx")


# =====================================================================
# 3) 家庭保單整理表
# =====================================================================
def insurance():
    Z = True
    wb = Workbook()
    hw = wb.active; hw.title = "使用說明"
    st = wb.create_sheet("設定")
    pl = wb.create_sheet("保單清單")
    cal = wb.create_sheet("繳費月曆")
    ov = wb.create_sheet("保障總覽")
    cl = wb.create_sheet("理賠紀錄")
    N = 60
    R1, R2 = 5, 4 + N

    howto(hw, "家庭保單整理表 – 使用說明", [
        "## 內容（6 個工作表）",
        "使用說明：本頁。",
        "設定：家庭成員、險種分類、繳費方式清單（下拉選單會自動使用）。",
        "保單清單：一張保單一列，記錄保險公司、險種、要保人／被保險人／受益人、保額、保費、繳費年期等。自動算出「年繳化保費」「繳費期滿日」「下次繳費日」與倒數天數。",
        "繳費月曆：自動列出每張保單 1～12 月各要繳多少，月底不怕忘記、年底一眼看到全年保費支出。",
        "保障總覽：依家人×險種彙整有效保單的保額，以及每位要保人每年要付的保費。",
        "理賠紀錄：記錄申請理賠的日期、事由、申請與核付金額、進度。",
        "## 開始使用",
        "1. 到「設定」填入家人名字（最多 8 位）與要用的險種分類。",
        "2. 到「保單清單」刪掉範例（選取列→右鍵→刪除），或直接覆蓋。淺黃色格子是輸入欄，白色格子是公式，請不要覆蓋。",
        "3. 生效日、繳費方式、每期保費、繳費年期填好後，下次繳費日會自動計算；30 天內到期顯示黃色，7 天內顯示紅色。",
        "　一年期、每年自動續保的保單，「繳費年期」可留空，系統會一直推算下次繳費日。",
        "4. 保額單位為「萬元」；醫療險日額、實支實付限額等可填在「保額說明」欄，方便查看。",
        "5. 保單狀態選「有效」才會計入繳費月曆與保障總覽。",
        "## 小提醒",
        "不知道自己有哪些保單？可向各保險公司或業務員索取保單明細，再逐一填入。",
        "下次繳費日以「生效日＋繳費週期」推算，實際扣款日以保險公司通知為準。",
        "本範本僅協助整理保單資料，不提供任何保險或理財建議；保障是否足夠請洽專業人員。",
        "檔案含個人資料，請妥善保存，不要上傳到公開的地方。",
        "## 可用軟體",
        "Microsoft Excel 2010 以上（Windows／Mac）、Google 試算表（檔案→匯入）、LibreOffice Calc。無巨集。",
        "Corn_Can 製作，僅供個人使用，請勿轉售。",
    ], cjk=Z)

    title(st, "設定", "黃色格子可修改，下拉選單會自動更新。", cjk=Z)
    list_col(st, 4, 1, "家庭成員", ["爸爸", "媽媽", "小明", "小美"], 8, cjk=Z)
    list_col(st, 4, 3, "險種分類", ["壽險", "醫療險", "重大傷病險", "癌症險", "意外險", "失能險", "長照險", "儲蓄險", "投資型", "其他"], 12, cjk=Z)
    header_row(st, 4, ["繳費方式", "每期月數"], cjk=Z, start_col=5)
    for i, (k, v) in enumerate([("年繳", 12), ("半年繳", 6), ("季繳", 3), ("月繳", 1), ("躉繳", 0)]):
        st.cell(row=5 + i, column=5, value=k).border = BORDER
        st.cell(row=5 + i, column=6, value=v).border = BORDER
        st.cell(row=5 + i, column=5).font = st.cell(row=5 + i, column=6).font = font(Z, size=10)
    st["E11"] = "躉繳＝一次繳清，不會出現在繳費月曆。"; st["E11"].font = font(Z, size=9, color="6B7280")
    widths(st, {"A": 16, "B": 3, "C": 16, "D": 3, "E": 12, "F": 10})
    landscape(st)

    title(pl, "保單清單", "一張保單一列。黃色＝輸入，白色＝自動計算。", cjk=Z)
    hdr = ["保險公司", "商品名稱", "保單號碼", "險種", "要保人", "被保險人", "受益人", "狀態", "生效日",
           "繳費方式", "每期保費", "繳費年期(年)", "年繳化保費", "繳費期滿日", "下次繳費日", "距繳費(天)",
           "主要保額(萬)", "保額說明", "保障至", "業務員／電話"]
    header_row(pl, 4, hdr, cjk=Z)
    ex = [("安心人壽", "定期壽險", "A12345678", "壽險", "爸爸", "爸爸", "媽媽", "有效", D(2019, 3, 15), "年繳", 12800, 20, 500, "身故／完全失能", "2039-03-14"),
          ("安心人壽", "實支實付醫療險", "A12345679", "醫療險", "爸爸", "爸爸", "媽媽", "有效", D(2019, 3, 15), "年繳", 6200, 20, None, "住院日額 2,000、雜費限額 15 萬", "續保至 75 歲"),
          ("和平產險", "意外傷害險", "P2023-0088", "意外險", "媽媽", "媽媽", "法定繼承人", "有效", D(2023, 8, 1), "月繳", 450, None, 300, "意外身故 300 萬、實支 3 萬", "一年期自動續保"),
          ("康寧人壽", "重大傷病一次金", "K55660021", "重大傷病險", "媽媽", "小明", "媽媽", "有效", D(2021, 12, 10), "半年繳", 3150, 20, 100, "確診一次給付", "終身"),
          ("康寧人壽", "兒童醫療險", "K55660022", "醫療險", "媽媽", "小美", "媽媽", "有效", D(2022, 1, 5), "季繳", 1980, 20, None, "住院日額 1,000", "續保至 75 歲"),
          ("安心人壽", "六年期儲蓄險", "A99001234", "儲蓄險", "爸爸", "爸爸", "爸爸", "已滿期", D(2016, 6, 1), "年繳", 50000, 6, 30, "已滿期領回", "2022-05-31")]
    for i, e in enumerate(ex):
        r = R1 + i
        for col, v in zip((1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 17, 18, 19), e):
            pl.cell(row=r, column=col, value=v)
        pl.cell(row=r, column=20, value=["王先生 0912-000-000", "王先生 0912-000-000", "客服 0800-000-000", "林小姐 0933-000-000", "林小姐 0933-000-000", ""][i])
    for r in range(R1, R2 + 1):
        m = f'IFERROR(VLOOKUP(J{r},設定!$E$5:$F$9,2,FALSE),0)'
        pl.cell(row=r, column=13, value=f'=IF(OR(A{r}="",N(K{r})=0),"",IF({m}=0,0,K{r}*12/{m}))')
        pl.cell(row=r, column=14, value=f'=IF(OR(NOT(ISNUMBER(I{r})),N(L{r})=0),"",EDATE(I{r},L{r}*12)-1)')
        nxt = (f'IF(I{r}>TODAY(),I{r},IF(EDATE(I{r},INT(DATEDIF(I{r},TODAY(),"m")/{m})*{m})>=TODAY(),'
               f'EDATE(I{r},INT(DATEDIF(I{r},TODAY(),"m")/{m})*{m}),EDATE(I{r},(INT(DATEDIF(I{r},TODAY(),"m")/{m})+1)*{m})))')
        pl.cell(row=r, column=15, value=f'=IF(OR(A{r}="",H{r}<>"有效",NOT(ISNUMBER(I{r})),{m}=0),"",IF(AND(ISNUMBER(N{r}),{nxt}>N{r}),"已繳清",{nxt}))')
        pl.cell(row=r, column=16, value=f'=IF(ISNUMBER(O{r}),O{r}-TODAY(),"")')
    grid(pl, R1, R2, 1, 20, cjk=Z, input_cols=(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 17, 18, 19, 20),
         fmt={9: "yyyy-mm-dd", 11: "#,##0", 13: "#,##0", 14: "yyyy-mm-dd", 15: "yyyy-mm-dd", 16: "0", 17: "#,##0"})
    dv_list(pl, "=設定!$C$5:$C$16", f"D{R1}:D{R2}")
    dv_list(pl, "=設定!$A$5:$A$12", f"E{R1}:F{R2}")
    dv_list(pl, '"有效,停效,已滿期,已解約"', f"H{R1}:H{R2}")
    dv_list(pl, "=設定!$E$5:$E$9", f"J{R1}:J{R2}")
    pl.conditional_formatting.add(f"O{R1}:P{R2}", FormulaRule(formula=[f'AND(ISNUMBER($P{R1}),$P{R1}<=7)'], fill=RED_FILL))
    pl.conditional_formatting.add(f"O{R1}:P{R2}", FormulaRule(formula=[f'AND(ISNUMBER($P{R1}),$P{R1}<=30)'], fill=AMBER_FILL))
    pl.conditional_formatting.add(f"A{R1}:T{R2}", FormulaRule(formula=[f'AND($A{R1}<>"",$H{R1}<>"有效")'], font=Font(color="9CA3AF")))
    widths(pl, {"A": 11, "B": 16, "C": 12, "D": 11, "E": 8, "F": 8, "G": 10, "H": 8, "I": 11, "J": 8,
                "K": 9, "L": 8, "M": 10, "N": 11, "O": 11, "P": 8, "Q": 9, "R": 24, "S": 13, "T": 18})
    pl.freeze_panes = "C5"; pl.auto_filter.ref = f"A4:T{R2}"
    landscape(pl); pl.print_area = "A1:T30"

    # 繳費月曆
    title(cal, "繳費月曆", None, cjk=Z)
    cal["A2"] = "年份："; cal["B2"] = 2027; cal["B2"].fill = INPUT_FILL; cal["B2"].font = font(Z, bold=True, size=12)
    cal["A2"].font = font(Z, bold=True)
    cal["C2"] = "只計入狀態為「有效」、仍在繳費期間內的保單（躉繳不列）。"; cal["C2"].font = font(Z, size=9, color="6B7280")
    header_row(cal, 4, ["商品名稱", "被保險人", "繳費方式"] + [f"{m}月" for m in range(1, 13)] + ["全年合計"], cjk=Z)
    L = "保單清單!"
    for i in range(N):
        r = 5 + i; s = R1 + i
        cal.cell(row=r, column=1, value=f'=IF({L}A{s}="","",{L}B{s})')
        cal.cell(row=r, column=2, value=f'=IF({L}A{s}="","",{L}F{s})')
        cal.cell(row=r, column=3, value=f'=IF({L}A{s}="","",{L}J{s})')
        mm = f'IFERROR(VLOOKUP({L}J{s},設定!$E$5:$F$9,2,FALSE),0)'
        for k in range(12):
            d = f"DATE($B$2,{k+1},DAY({L}I{s}))"
            d = f"MIN(DATE($B$2,{k+1},DAY({L}I{s})),EOMONTH(DATE($B$2,{k+1},1),0))"
            cond = (f'AND(MOD(({k+1}-MONTH({L}I{s}))+($B$2-YEAR({L}I{s}))*12,{mm})=0,'
                    f'{d}>={L}I{s},OR(NOT(ISNUMBER({L}N{s})),{d}<={L}N{s}))')
            cal.cell(row=r, column=4 + k, value=f'=IF(OR({L}A{s}="",{L}H{s}<>"有效",NOT(ISNUMBER({L}I{s})),{mm}=0),0,IF({cond},N({L}K{s}),0))')
        cal.cell(row=r, column=16, value=f"=SUM(D{r}:O{r})")
    tr = 5 + N
    cal.cell(row=tr, column=1, value="每月合計")
    for c in range(4, 17):
        cal.cell(row=tr, column=c, value=f"=SUM({CL(c)}5:{CL(c)}{tr-1})")
    grid(cal, 5, tr, 1, 16, cjk=Z, fmt={c: '#,##0;-#,##0;""' for c in range(4, 17)})
    for c in range(1, 17):
        cal.cell(row=tr, column=c).font = font(Z, bold=True); cal.cell(row=tr, column=c).fill = CORN_FILL
    cal.conditional_formatting.add(f"D5:O{tr-1}", CellIsRule(operator="greaterThan", formula=["0"], fill=PatternFill("solid", fgColor=LIGHT)))
    widths(cal, {"A": 18, "B": 9, "C": 8, **{CL(c): 8 for c in range(4, 16)}, "P": 11})
    cal.freeze_panes = "D5"
    landscape(cal); cal.page_setup.fitToHeight = 1; cal.print_area = f"A1:P{tr}"

    # 保障總覽
    title(ov, "保障總覽", "有效保單的主要保額（萬元），依被保險人與險種加總。", cjk=Z)
    F_, D_, H_, Q_, E_, M_ = (f"{L}${c}${R1}:${c}${R2}" for c in "FDHQEM")
    header_row(ov, 4, ["被保險人"] + [None] * 10 + ["有效保單數"], cjk=Z)
    for k in range(10):
        c = ov.cell(row=4, column=2 + k, value=f'=IF(設定!C{5+k}="","",設定!C{5+k})')
        c.font = font(Z, bold=True, color="FFFFFF"); c.fill = HEAD_FILL; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i in range(8):
        r = 5 + i
        ov.cell(row=r, column=1, value=f'=IF(設定!A{5+i}="","",設定!A{5+i})')
        for k in range(10):
            col = CL(2 + k)
            ov.cell(row=r, column=2 + k, value=f'=IF(OR($A{r}="",{col}$4=""),"",SUMIFS({Q_},{F_},$A{r},{D_},{col}$4,{H_},"有效"))')
        ov.cell(row=r, column=12, value=f'=IF($A{r}="","",COUNTIFS({F_},$A{r},{H_},"有效"))')
    grid(ov, 5, 12, 1, 12, cjk=Z, fmt={c: '#,##0;-#,##0;"-"' for c in range(2, 12)})
    ov.conditional_formatting.add("B5:K12", CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL))

    header_row(ov, 15, ["要保人（付款人）", "有效保單數", "年繳化保費", "每月平均"], cjk=Z)
    for i in range(8):
        r = 16 + i
        ov.cell(row=r, column=1, value=f'=IF(設定!A{5+i}="","",設定!A{5+i})')
        ov.cell(row=r, column=2, value=f'=IF(A{r}="","",COUNTIFS({E_},A{r},{H_},"有效"))')
        ov.cell(row=r, column=3, value=f'=IF(A{r}="","",SUMIFS({M_},{E_},A{r},{H_},"有效"))')
        ov.cell(row=r, column=4, value=f'=IF(A{r}="","",ROUND(C{r}/12,0))')
    ov["A24"] = "合計"; ov["B24"] = "=SUM(B16:B23)"; ov["C24"] = "=SUM(C16:C23)"; ov["D24"] = "=SUM(D16:D23)"
    grid(ov, 16, 24, 1, 4, cjk=Z, fmt={3: "#,##0", 4: "#,##0"})
    for c in range(1, 5):
        ov.cell(row=24, column=c).font = font(Z, bold=True)
    ov["F15"] = "註：醫療險多為日額或限額，主要保額可能空白，請看保單清單的「保額說明」。"
    ov["F15"].font = font(Z, size=9, color="6B7280")
    ch = BarChart(); ch.type = "col"; ch.title = "每位要保人年繳化保費"; ch.style = 10
    ch.add_data(Reference(ov, min_col=3, min_row=15, max_row=19), titles_from_data=True)
    ch.set_categories(Reference(ov, min_col=1, min_row=16, max_row=19))
    ch.legend = None; ch.height = 7; ch.width = 13
    ch.series[0].graphicalProperties.solidFill = "F2C14E"
    ov.add_chart(ch, "F17")
    widths(ov, {"A": 16, **{CL(c): 10 for c in range(2, 12)}, "L": 11})
    landscape(ov); ov.page_setup.fitToHeight = 1; ov.print_area = "A1:L32"

    # 理賠紀錄
    title(cl, "理賠紀錄", "每次申請理賠記一列，追蹤進度與金額。", cjk=Z)
    header_row(cl, 4, ["事故／就醫日", "被保險人", "保單（商品名稱）", "事由", "送件日", "申請金額", "核付金額", "進度", "等待天數", "備註"], cjk=Z)
    cex = [(D(2026, 9, 2), "小美", "兒童醫療險", "腸胃炎住院 3 天", D(2026, 9, 10), 6800, 6000, "已核付", "部分自費項目不給付"),
           (D(2026, 10, 18), "爸爸", "實支實付醫療險", "門診手術", D(2026, 10, 25), 12500, None, "審核中", "補件：診斷證明正本")]
    for i, e in enumerate(cex):
        for col, v in zip((1, 2, 3, 4, 5, 6, 7, 8, 10), e):
            cl.cell(row=5 + i, column=col, value=v)
    for r in range(5, 55):
        cl.cell(row=r, column=9, value=f'=IF(OR(NOT(ISNUMBER(E{r})),H{r}="已核付",H{r}="不給付"),"",TODAY()-E{r})')
    grid(cl, 5, 54, 1, 10, cjk=Z, input_cols=(1, 2, 3, 4, 5, 6, 7, 8, 10), fmt={1: "yyyy-mm-dd", 5: "yyyy-mm-dd", 6: "#,##0", 7: "#,##0"})
    dv_list(cl, "=設定!$A$5:$A$12", "B5:B54")
    dv_list(cl, '"準備文件,審核中,補件中,已核付,不給付"', "H5:H54")
    cl.conditional_formatting.add("A5:J54", FormulaRule(formula=['$H5="已核付"'], fill=GREEN_FILL))
    cl.conditional_formatting.add("A5:J54", FormulaRule(formula=['OR($H5="審核中",$H5="補件中")'], fill=AMBER_FILL))
    cl["L4"] = "申請總額"; cl["M4"] = "=SUM(F5:F54)"
    cl["L5"] = "已核付總額"; cl["M5"] = "=SUM(G5:G54)"
    cl["L6"] = "處理中件數"; cl["M6"] = '=COUNTIF(H5:H54,"審核中")+COUNTIF(H5:H54,"補件中")+COUNTIF(H5:H54,"準備文件")'
    grid(cl, 4, 6, 12, 13, cjk=Z, fmt={13: "#,##0"})
    for c in ("L4", "L5", "L6"):
        cl[c].font = font(Z, bold=True); cl[c].fill = PatternFill("solid", fgColor=LIGHT)
    widths(cl, {"A": 12, "B": 9, "C": 16, "D": 20, "E": 12, "F": 10, "G": 10, "H": 9, "I": 9, "J": 22, "K": 3, "L": 12, "M": 10})
    cl.freeze_panes = "A5"
    landscape(cl); cl.print_area = "A1:M25"
    wb.active = 2
    wb.save(f"{OUT}/family-insurance-organizer-tw.xlsx")


gantt(); inventory(); insurance()
print("ok")
