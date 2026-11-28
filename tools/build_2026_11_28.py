import sys, datetime as dt
sys.path.insert(0, "/home/user/corn-can-store-assets/tools")
from common import *
from openpyxl import Workbook
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment

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
    b.fill = CORN_FILL if False else PatternFill("solid", fgColor=LIGHT)
    b.alignment = Alignment(horizontal="center", vertical="center")
    b.number_format = fmtc
    b.border = BORDER
    ws.row_dimensions[r + 1].height = 30


from openpyxl.styles import PatternFill

MONEY = '"$"#,##0.00;[Red]-"$"#,##0.00'
MONEY0 = '"$"#,##0;[Red]-"$"#,##0'
DATE = "yyyy-mm-dd"


# ---------------------------------------------------------------- rental
def rental():
    wb = Workbook()
    hw = wb.active; hw.title = "How to Use"
    howto(hw, "Rental Property Income & Expense Tracker – How to Use", [
        "## What's inside (7 sheets)",
        "How to Use – this page.",
        "Settings – tax year, your properties (up to 10) and your expense / income categories (editable).",
        "Leases – one row per tenant or lease (up to 40): property, unit, tenant, lease dates, monthly rent, deposit. Flags leases ending within 60 days or already expired.",
        "Rent Roll – automatic 12-month grid: rent received per lease per month for the tax year, compared with the rent due. Unpaid months show in red, partial payments in amber.",
        "Income – log every payment received (rent, late fees, pet rent, other) – up to 500 rows.",
        "Expenses – log every expense (up to 800 rows) with property, category, vendor and amount.",
        "Summary – per-property income, expenses and net, an expense-by-category table (categories modeled on common rental expense groupings such as those on IRS Schedule E), a monthly cash-flow table and chart.",
        "## Quick start",
        "1. Settings: type the tax year in B4 (the sample uses 2026 – change it to 2027 when you start the new year) and list your property names in A8:A17. Category lists can be renamed.",
        "2. Leases: add each tenant with the property (drop-down), lease start/end and monthly rent.",
        "3. Income: every time rent arrives, add a row – date, property, tenant (type the name exactly as on Leases), type 'Rent', the month it is for, and the amount.",
        "4. Expenses: add each bill or receipt with property and category.",
        "5. Check Rent Roll for missing rent, and Summary for totals by property and category.",
        "## Tips",
        "Light-yellow cells are inputs; white cells are formulas. Delete the sample rows (gray example names) before using for real.",
        "'Rent For (month)' on the Income sheet decides which month a payment counts toward on Rent Roll (useful when January rent arrives on Dec 30). Enter any date inside that month.",
        "Summary and the monthly table use the payment / expense date and the tax year in Settings.",
        "Mortgage principal is not an expense – only interest. Record capital improvements separately if your tax preparer depreciates them.",
        "## Requirements",
        "Microsoft Excel 2019 / 2021 / Microsoft 365, Excel for Mac, or LibreOffice Calc. Also opens in Google Sheets (upload, then File > Save as Google Sheets).",
        "## Disclaimer",
        "This is a record-keeping tool, not tax, legal or accounting advice. Category names are for organizing records; confirm how to report items with a qualified tax professional.",
    ])

    st = wb.create_sheet("Settings")
    title(st, "Settings", "Yellow cells are yours to edit.")
    st["A4"] = "Tax year"; st["B4"] = 2026
    st["A5"] = "Lease warning (days before end)"; st["B5"] = 60
    for a in ("A4", "A5"):
        st[a].font = font(bold=True)
    for a in ("B4", "B5"):
        st[a].fill = INPUT_FILL; st[a].border = BORDER
    header_row(st, 7, ["Properties"])
    props = ["Maple St Duplex", "Oak Ave Condo", "Pine Rd House"]
    for i in range(10):
        c = st.cell(row=8 + i, column=1, value=props[i] if i < len(props) else None)
        c.fill = INPUT_FILL; c.border = BORDER
    header_row(st, 7, ["Expense categories"], start_col=3)
    cats = ["Advertising", "Auto & travel", "Cleaning & maintenance", "Commissions", "Insurance",
            "Legal & professional fees", "Management fees", "Mortgage interest", "Other interest", "Repairs",
            "Supplies", "Property taxes", "Utilities", "HOA fees", "Other"]
    for i in range(20):
        c = st.cell(row=8 + i, column=3, value=cats[i] if i < len(cats) else None)
        c.fill = INPUT_FILL; c.border = BORDER
    header_row(st, 7, ["Income types"], start_col=5)
    inc = ["Rent", "Late fee", "Pet rent", "Parking / storage", "Other income"]
    for i in range(8):
        c = st.cell(row=8 + i, column=5, value=inc[i] if i < len(inc) else None)
        c.fill = INPUT_FILL; c.border = BORDER
    header_row(st, 7, ["Payment methods"], start_col=7)
    for i, v in enumerate(["Bank transfer", "Check", "Cash", "Card", "Online app", "Other"]):
        c = st.cell(row=8 + i, column=7, value=v); c.fill = INPUT_FILL; c.border = BORDER
    widths(st, {"A": 32, "B": 12, "C": 28, "D": 3, "E": 20, "F": 3, "G": 18})
    PROPS = "Settings!$A$8:$A$17"; CATS = "Settings!$C$8:$C$27"; INC = "Settings!$E$8:$E$15"; METH = "Settings!$G$8:$G$13"

    # Leases
    le = wb.create_sheet("Leases")
    title(le, "Leases & Tenants", "One row per lease. Status updates automatically from today's date.")
    H = ["Property", "Unit", "Tenant", "Phone / Email", "Lease Start", "Lease End", "Monthly Rent", "Deposit Held",
         "Rent Due Day", "Days to End", "Status", "Notes"]
    header_row(le, 4, H)
    L1, L2 = 5, 44
    sample = [("Maple St Duplex", "A", "J. Rivera", "", D(2026, 9, 1), D(2027, 8, 31), 1450, 1450, 1),
              ("Maple St Duplex", "B", "K. Chen", "", D(2026, 9, 1), D(2027, 8, 31), 1400, 1400, 1),
              ("Oak Ave Condo", "12", "M. Patel", "", D(2026, 9, 1), D(2027, 1, 15), 1750, 1750, 1),
              ("Pine Rd House", "", "The Nguyen family", "", D(2026, 9, 15), D(2027, 3, 14), 2100, 2100, 15)]
    for i, s in enumerate(sample):
        for j, v in enumerate(s):
            le.cell(row=L1 + i, column=1 + j, value=v)
    for r in range(L1, L2 + 1):
        le.cell(row=r, column=10, value=f'=IF(F{r}="","",F{r}-TODAY())')
        le.cell(row=r, column=11, value=f'=IF(C{r}="","",IF(F{r}="","Month-to-month",IF(J{r}<0,"Ended",IF(J{r}<=Settings!$B$5,"Ending soon","Active"))))')
    grid(le, L1, L2, 1, 12, input_cols=(1, 2, 3, 4, 5, 6, 7, 8, 9, 12),
         fmt={5: DATE, 6: DATE, 7: MONEY0, 8: MONEY0, 10: "0"})
    dv_list(le, f"={PROPS}", f"A{L1}:A{L2}")
    le.conditional_formatting.add(f"K{L1}:K{L2}", FormulaRule(formula=[f'K{L1}="Ended"'], fill=RED_FILL))
    le.conditional_formatting.add(f"K{L1}:K{L2}", FormulaRule(formula=[f'K{L1}="Ending soon"'], fill=AMBER_FILL))
    le.conditional_formatting.add(f"K{L1}:K{L2}", FormulaRule(formula=[f'K{L1}="Active"'], fill=GREEN_FILL))
    widths(le, {"A": 20, "B": 7, "C": 22, "D": 22, "E": 12, "F": 12, "G": 12, "H": 12, "I": 9, "J": 9, "K": 15, "L": 26})
    le.freeze_panes = "A5"; landscape(le)

    # Income
    ic = wb.create_sheet("Income")
    title(ic, "Income Log", "Every payment received. 'Rent For' = any date in the month the payment covers.")
    header_row(ic, 4, ["Date Received", "Property", "Tenant", "Type", "Rent For (month)", "Amount", "Method", "Notes", "Month (auto)"])
    I1, I2 = 5, 504
    y = 2026
    rows = []
    for m in range(9, 12):
        rows += [(D(y, m, 1), "Maple St Duplex", "J. Rivera", "Rent", D(y, m, 1), 1450, "Bank transfer"),
                 (D(y, m, 2), "Maple St Duplex", "K. Chen", "Rent", D(y, m, 1), 1400, "Online app"),
                 (D(y, m, 3), "Oak Ave Condo", "M. Patel", "Rent", D(y, m, 1), 1750, "Check")]
    rows += [(D(y, 9, 15), "Pine Rd House", "The Nguyen family", "Rent", D(y, 9, 1), 2100, "Bank transfer"),
             (D(y, 10, 15), "Pine Rd House", "The Nguyen family", "Rent", D(y, 10, 1), 2100, "Bank transfer"),
             (D(y, 11, 18), "Pine Rd House", "The Nguyen family", "Rent", D(y, 11, 1), 1000, "Bank transfer"),
             (D(y, 11, 20), "Pine Rd House", "The Nguyen family", "Late fee", None, 50, "Bank transfer")]
    rows = [r for r in rows if not (r[2] == "M. Patel" and r[4] and r[4].month == 11)]  # March unpaid example
    for i, s in enumerate(rows):
        for j, v in enumerate(s):
            ic.cell(row=I1 + i, column=1 + j, value=v)
    for r in range(I1, I2 + 1):
        ic.cell(row=r, column=9, value=f'=IF(A{r}="","",DATE(YEAR(A{r}),MONTH(A{r}),1))')
    grid(ic, I1, I2, 1, 9, input_cols=(1, 2, 3, 4, 5, 6, 7, 8), fmt={1: DATE, 5: "mmm yyyy", 6: MONEY, 9: "mmm yyyy"})
    dv_list(ic, f"={PROPS}", f"B{I1}:B{I2}")
    dv_list(ic, f"={INC}", f"D{I1}:D{I2}")
    dv_list(ic, f"={METH}", f"G{I1}:G{I2}")
    widths(ic, {"A": 13, "B": 20, "C": 22, "D": 16, "E": 14, "F": 12, "G": 15, "H": 26, "I": 11})
    ic.freeze_panes = "A5"; landscape(ic)

    # Expenses
    ex = wb.create_sheet("Expenses")
    title(ex, "Expense Log", "Every bill or receipt. Choose the property and category from the drop-downs.")
    header_row(ex, 4, ["Date", "Property", "Category", "Vendor / Payee", "Description", "Amount", "Method", "Receipt Saved?"])
    E1, E2 = 5, 804
    exs = [(D(y, 9, 5), "Maple St Duplex", "Mortgage interest", "First Bank", "September interest portion", 612.40, "Bank transfer", "Yes"),
           (D(y, 9, 12), "Oak Ave Condo", "HOA fees", "Oak Ave HOA", "Monthly dues", 285, "Bank transfer", "Yes"),
           (D(y, 9, 20), "Pine Rd House", "Repairs", "QuickFix Plumbing", "Kitchen sink leak", 240, "Card", "Yes"),
           (D(y, 10, 3), "Maple St Duplex", "Insurance", "HomeSure", "Annual landlord policy", 1180, "Card", "Yes"),
           (D(y, 10, 14), "Pine Rd House", "Cleaning & maintenance", "GreenCut Lawn", "Yard service", 120, "Check", "No"),
           (D(y, 11, 1), "Oak Ave Condo", "Utilities", "City Water", "Water & sewer", 68.25, "Online app", "Yes")]
    for i, s in enumerate(exs):
        for j, v in enumerate(s):
            ex.cell(row=E1 + i, column=1 + j, value=v)
    grid(ex, E1, E2, 1, 8, input_cols=(1, 2, 3, 4, 5, 6, 7, 8), fmt={1: DATE, 6: MONEY})
    dv_list(ex, f"={PROPS}", f"B{E1}:B{E2}")
    dv_list(ex, f"={CATS}", f"C{E1}:C{E2}")
    dv_list(ex, f"={METH}", f"G{E1}:G{E2}")
    dv_list(ex, '"Yes,No"', f"H{E1}:H{E2}")
    widths(ex, {"A": 12, "B": 20, "C": 24, "D": 22, "E": 30, "F": 12, "G": 15, "H": 10})
    ex.freeze_panes = "A5"; landscape(ex)

    # Rent roll
    rr = wb.create_sheet("Rent Roll")
    title(rr, "Rent Roll", "Rent received per lease per month (tax year from Settings). Red = nothing received for a month the lease was active; amber = partial.")
    rr["A3"] = '="Tax year: "&Settings!B4'; rr["A3"].font = font(bold=True)
    header_row(rr, 4, ["Property", "Unit", "Tenant", "Monthly Rent"] + [f"=TEXT(DATE(Settings!$B$4,{m},1),\"mmm\")" for m in range(1, 13)] + ["Year Total", "Balance Due"])
    for m in range(1, 13):
        rr.cell(row=4, column=4 + m).number_format = "@"
    IC = f"Income!$C${I1}:$C${I2}"; IE = f"Income!$E${I1}:$E${I2}"; IF_ = f"Income!$F${I1}:$F${I2}"; ID = f"Income!$D${I1}:$D${I2}"
    for i in range(40):
        r = 5 + i; lr = L1 + i
        rr.cell(row=r, column=1, value=f'=IF(Leases!C{lr}="","",Leases!A{lr})')
        rr.cell(row=r, column=2, value=f'=IF(Leases!C{lr}="","",Leases!B{lr}&"")')
        rr.cell(row=r, column=3, value=f'=IF(Leases!C{lr}="","",Leases!C{lr})')
        rr.cell(row=r, column=4, value=f'=IF(Leases!C{lr}="","",Leases!G{lr})')
        for m in range(1, 13):
            col = 4 + m
            ms = f"DATE(Settings!$B$4,{m},1)"; me = f"EOMONTH({ms},0)"
            rr.cell(row=r, column=col, value=f'=IF($C{r}="","",SUMIFS({IF_},{IC},$C{r},{ID},"Rent",{IE},">="&{ms},{IE},"<="&{me}))')
        rr.cell(row=r, column=17, value=f'=IF(C{r}="","",SUM(E{r}:P{r}))')
        # due months = months overlapping lease within tax year, up to current month
        due = "+".join(
            f'(AND(DATE(Settings!$B$4,{m},1)<=MIN(TODAY(),IF(Leases!$F{lr}="",DATE(9999,1,1),Leases!$F{lr})),EOMONTH(DATE(Settings!$B$4,{m},1),0)>=Leases!$E{lr}))*1'
            for m in range(1, 13))
        rr.cell(row=r, column=18, value=f'=IF(C{r}="","",MAX(0,({due})*D{r}-Q{r}))')
    grid(rr, 5, 44, 1, 18, fmt={**{c: MONEY0 for c in range(4, 19)}})
    for m in range(1, 13):
        cl = CL(4 + m); lr0 = L1
        active = (f'AND($C5<>"",DATE(Settings!$B$4,{m},1)<=TODAY(),DATE(Settings!$B$4,{m},1)<=IF(INDEX(Leases!$F${L1}:$F${L2},ROW()-4)="",DATE(9999,1,1),INDEX(Leases!$F${L1}:$F${L2},ROW()-4)),'
                  f'EOMONTH(DATE(Settings!$B$4,{m},1),0)>=INDEX(Leases!$E${L1}:$E${L2},ROW()-4))')
        rr.conditional_formatting.add(f"{cl}5:{cl}44", FormulaRule(formula=[f'AND({active},N({cl}5)=0)'], fill=RED_FILL))
        rr.conditional_formatting.add(f"{cl}5:{cl}44", FormulaRule(formula=[f'AND({active},N({cl}5)>0,N({cl}5)<$D5)'], fill=AMBER_FILL))
        rr.conditional_formatting.add(f"{cl}5:{cl}44", FormulaRule(formula=[f'AND($C5<>"",N({cl}5)>=$D5,N($D5)>0)'], fill=GREEN_FILL))
    rr.conditional_formatting.add("R5:R44", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
    rr.cell(row=45, column=3, value="Total").font = font(bold=True)
    for c in range(4, 19):
        cc = rr.cell(row=45, column=c, value=f"=SUM({CL(c)}5:{CL(c)}44)")
        cc.font = font(bold=True); cc.number_format = MONEY0; cc.border = BORDER; cc.fill = GREY_FILL
    rr["A47"] = "Balance Due = rent due for lease months up to today (in the tax year) minus rent received. It assumes a full month's rent is due for every month the lease overlaps."
    rr["A47"].font = font(size=9, italic=True, color="6B7280")
    widths(rr, {"A": 18, "B": 6, "C": 20, "D": 11, **{CL(c): 8.5 for c in range(5, 17)}, "Q": 11, "R": 11})
    rr.freeze_panes = "E5"; landscape(rr)

    # Summary
    sm = wb.create_sheet("Summary")
    title(sm, "Summary", None)
    sm["A2"] = '="Tax year "&Settings!B4&" – all figures by payment / expense date"'
    sm["A2"].font = font(size=10, italic=True, color="6B7280")
    YS = "DATE(Settings!$B$4,1,1)"; YE = "DATE(Settings!$B$4,12,31)"
    IA = f"Income!$A${I1}:$A${I2}"; IB = f"Income!$B${I1}:$B${I2}"
    EA = f"Expenses!$A${E1}:$E${E2}".replace(":$E$", ":$A$"); EB = f"Expenses!$B${E1}:$B${E2}"; EC = f"Expenses!$C${E1}:$C${E2}"; EF = f"Expenses!$F${E1}:$F${E2}"
    kpi(sm, 4, 1, "Total income", f'=SUMIFS({IF_},{IA},">="&{YS},{IA},"<="&{YE})', MONEY0)
    kpi(sm, 4, 2, "Total expenses", f'=SUMIFS({EF},{EA},">="&{YS},{EA},"<="&{YE})', MONEY0)
    kpi(sm, 4, 3, "Net income", "=A5-B5", MONEY0)
    kpi(sm, 4, 4, "Rent balance due", "='Rent Roll'!R45", MONEY0)
    kpi(sm, 4, 5, "Leases ending soon", f'=COUNTIF(Leases!$K${L1}:$K${L2},"Ending soon")', "0")
    kpi(sm, 4, 6, "Expenses w/o receipt", f'=COUNTIFS(Expenses!$H${E1}:$H${E2},"No",{EA},">="&{YS},{EA},"<="&{YE})', "0")
    sm.conditional_formatting.add("C5", CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
    sm.conditional_formatting.add("D5", CellIsRule(operator="greaterThan", formula=["0"], fill=AMBER_FILL))
    # by property
    header_row(sm, 8, ["Property", "Rent", "Other Income", "Total Income", "Expenses", "Net"])
    for i in range(10):
        r = 9 + i
        sm.cell(row=r, column=1, value=f'=IF(Settings!A{8+i}="","",Settings!A{8+i})')
        rent = f'SUMIFS({IF_},{IB},$A{r},{ID},"Rent",{IA},">="&{YS},{IA},"<="&{YE})'
        allinc = f'SUMIFS({IF_},{IB},$A{r},{IA},">="&{YS},{IA},"<="&{YE})'
        sm.cell(row=r, column=2, value=f'=IF(A{r}="","",{rent})')
        sm.cell(row=r, column=3, value=f'=IF(A{r}="","",{allinc}-B{r})')
        sm.cell(row=r, column=4, value=f'=IF(A{r}="","",B{r}+C{r})')
        sm.cell(row=r, column=5, value=f'=IF(A{r}="","",SUMIFS({EF},{EB},$A{r},{EA},">="&{YS},{EA},"<="&{YE}))')
        sm.cell(row=r, column=6, value=f'=IF(A{r}="","",D{r}-E{r})')
    grid(sm, 9, 18, 1, 6, fmt={c: MONEY0 for c in range(2, 7)})
    sm.conditional_formatting.add("F9:F18", CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
    # category x property
    header_row(sm, 21, ["Expense category"] + [f'=IF(Settings!A{8+i}="","",Settings!A{8+i})' for i in range(10)] + ["Total"])
    for i in range(20):
        r = 22 + i
        sm.cell(row=r, column=1, value=f'=IF(Settings!C{8+i}="","",Settings!C{8+i})')
        for p in range(10):
            col = 2 + p; pc = f"{CL(col)}$21"
            sm.cell(row=r, column=col, value=f'=IF(OR($A{r}="",{pc}=""),"",SUMIFS({EF},{EC},$A{r},{EB},{pc},{EA},">="&{YS},{EA},"<="&{YE}))')
        sm.cell(row=r, column=12, value=f'=IF($A{r}="","",SUMIFS({EF},{EC},$A{r},{EA},">="&{YS},{EA},"<="&{YE}))')
    grid(sm, 22, 41, 1, 12, fmt={c: MONEY0 for c in range(2, 13)})
    sm.cell(row=42, column=1, value="Total").font = font(bold=True)
    for c in range(2, 13):
        cc = sm.cell(row=42, column=c, value=f"=SUM({CL(c)}22:{CL(c)}41)"); cc.number_format = MONEY0; cc.font = font(bold=True); cc.fill = GREY_FILL; cc.border = BORDER
    # monthly
    header_row(sm, 45, ["Month", "Income", "Expenses", "Net"])
    for m in range(1, 13):
        r = 45 + m
        sm.cell(row=r, column=1, value=f"=DATE(Settings!$B$4,{m},1)").number_format = "mmm"
        ms = f"DATE(Settings!$B$4,{m},1)"; me = f"EOMONTH({ms},0)"
        sm.cell(row=r, column=2, value=f'=SUMIFS({IF_},{IA},">="&{ms},{IA},"<="&{me})')
        sm.cell(row=r, column=3, value=f'=SUMIFS({EF},{EA},">="&{ms},{EA},"<="&{me})')
        sm.cell(row=r, column=4, value=f"=B{r}-C{r}")
    grid(sm, 46, 57, 1, 4, fmt={1: "mmm", 2: MONEY0, 3: MONEY0, 4: MONEY0})
    ch = BarChart(); ch.type = "col"; ch.title = "Monthly income vs expenses"; ch.style = 10
    ch.add_data(Reference(sm, min_col=2, max_col=3, min_row=45, max_row=57), titles_from_data=True)
    ch.set_categories(Reference(sm, min_col=1, min_row=46, max_row=57))
    ch.height = 7.5; ch.width = 16
    ch.series[0].graphicalProperties.solidFill = "F2C14E"; ch.series[1].graphicalProperties.solidFill = "2F3E46"
    sm.add_chart(ch, "F45")
    widths(sm, {"A": 24, **{CL(c): 15 for c in range(2, 13)}})
    landscape(sm)
    wb.move_sheet("Summary", offset=-5)
    wb.move_sheet("Rent Roll", offset=-2)
    wb.active = 1
    wb.save(f"{OUT}/rental-property-income-expense-tracker.xlsx")


# ---------------------------------------------------------------- car log (zh)
def carlog():
    Z = True
    wb = Workbook()
    hw = wb.active; hw.title = "使用說明"
    howto(hw, "汽機車保養＆油耗紀錄表 – 使用說明", [
        "## 內容（共 7 個工作表）",
        "使用說明 – 本頁。",
        "車輛資料 – 最多 6 台車（汽車、機車都可以），填車名、車牌、購入日期與購入時里程。",
        "加油紀錄 – 每次加油填日期、車輛、里程表讀數、公升數與金額，自動算出這次行駛公里數、油耗（km/L）與每公里油錢。最多 600 筆。",
        "保養維修 – 每次保養、換零件、維修填日期、里程、項目、店家與金額。最多 400 筆。",
        "保養提醒 – 每台車每個保養項目設定「每幾公里」或「每幾個月」，自動抓最後一次保養紀錄，算出下次保養里程／日期，並標示「該保養了」或「快到了」。",
        "稅費保險到期 – 牌照稅、燃料費、強制險、任意險、定期檢驗、停車費等到期日與金額，自動倒數天數並標示即將到期。",
        "用車成本總覽 – 選年度，自動統計每台車的油錢、保養維修、稅費保險、行駛里程、平均油耗與每公里成本，並有每月花費表與圖表。",
        "## 快速開始",
        "1. 車輛資料：把範例車輛改成你的車（車名要和其他工作表選的一樣）。",
        "2. 加油紀錄：每次加油記一筆。建議每次都加滿，油耗會最準（油耗＝這次與上次加油之間的里程 ÷ 這次加的公升數）。",
        "3. 保養維修：記下每次保養。「項目」請從下拉選單選，保養提醒才能自動對應。",
        "4. 保養提醒：每台車要追蹤的項目各一列，填間隔（公里或月數，兩個都填時，以先到者為準）。",
        "5. 稅費保險到期：填上繳費或到期日期；繳完後把日期改成下一期。",
        "## 小提醒",
        "淺黃色格子為輸入欄，白色格子為公式，請勿覆蓋。範例資料可以直接刪除或覆寫。",
        "保養間隔、稅費繳納期間等請以你的車輛使用手冊、原廠建議與監理機關、保險公司的通知為準；本表的範例數字僅供示範。",
        "「目前里程」取該車在加油紀錄與保養紀錄中最大的里程數，記得常記錄才會準。",
        "## 開啟方式",
        "Microsoft Excel 2019／2021／Microsoft 365、Mac 版 Excel、LibreOffice Calc；也可上傳到 Google 雲端硬碟用 Google 試算表開啟。",
    ], cjk=Z)

    cv = wb.create_sheet("車輛資料")
    title(cv, "車輛資料", "最多 6 台。車名請勿重複。", cjk=Z)
    header_row(cv, 4, ["車名", "類型", "車牌", "廠牌／型號", "購入日期", "購入時里程", "油品", "備註"], cjk=Z)
    cars = [("家用車", "汽車", "ABC-1234", "（自填）", D(2022, 3, 10), 0, "95 無鉛", ""),
            ("通勤機車", "機車", "XYZ-567", "（自填）", D(2023, 7, 1), 0, "95 無鉛", "")]
    for i, s in enumerate(cars):
        for j, v in enumerate(s):
            cv.cell(row=5 + i, column=1 + j, value=v)
    grid(cv, 5, 10, 1, 8, cjk=Z, input_cols=range(1, 9), fmt={5: DATE, 6: "#,##0"})
    dv_list(cv, '"汽車,機車,其他"', "B5:B10")
    widths(cv, {"A": 16, "B": 8, "C": 12, "D": 18, "E": 12, "F": 12, "G": 12, "H": 24})
    landscape(cv)
    CARS = "車輛資料!$A$5:$A$10"

    # fuel
    fu = wb.create_sheet("加油紀錄")
    title(fu, "加油紀錄", "建議每次加滿；油耗＝本次與上次加油間的行駛里程 ÷ 本次公升數。", cjk=Z)
    header_row(fu, 4, ["日期", "車輛", "里程表（km）", "公升", "每公升單價", "金額", "加油站", "本次行駛 km", "油耗 km/L", "每公里油錢"], cjk=Z)
    F1, F2 = 5, 604
    fuel = [(D(2026, 10, 3), "家用車", 41250, 38.5, 30.2), (D(2026, 10, 17), "家用車", 41702, 36.9, 30.2),
            (D(2026, 11, 2), "家用車", 42180, 39.4, 30.5), (D(2026, 10, 6), "通勤機車", 12030, 4.2, 30.2),
            (D(2026, 10, 20), "通勤機車", 12205, 4.0, 30.2), (D(2026, 11, 5), "通勤機車", 12390, 4.3, 30.5)]
    for i, s in enumerate(fuel):
        for j, v in enumerate(s):
            fu.cell(row=F1 + i, column=1 + j, value=v)
    for r in range(F1, F2 + 1):
        fu.cell(row=r, column=6, value=f'=IF(OR(D{r}="",E{r}=""),"",ROUND(D{r}*E{r},0))')
        prev = f'_xlfn.MAXIFS($C${F1}:$C${F2},$B${F1}:$B${F2},B{r},$C${F1}:$C${F2},"<"&C{r})'
        fu.cell(row=r, column=8, value=f'=IF(OR(B{r}="",C{r}=""),"",IF({prev}=0,0,C{r}-{prev}))')
        fu.cell(row=r, column=9, value=f'=IF(OR(H{r}="",N(H{r})=0,N(D{r})=0),"",ROUND(H{r}/D{r},1))')
        fu.cell(row=r, column=10, value=f'=IF(OR(I{r}="",N(F{r})=0),"",ROUND(F{r}/H{r},2))')
    grid(fu, F1, F2, 1, 10, cjk=Z, input_cols=(1, 2, 3, 4, 5, 7), fmt={1: DATE, 3: "#,##0", 4: "0.00", 5: "0.0", 6: "#,##0", 8: "#,##0", 9: "0.0", 10: "0.00"})
    dv_list(fu, f"={CARS}", f"B{F1}:B{F2}")
    widths(fu, {"A": 12, "B": 14, "C": 13, "D": 8, "E": 11, "F": 10, "G": 14, "H": 12, "I": 11, "J": 11})
    fu.freeze_panes = "A5"; landscape(fu)
    fu["A3"] = "說明：每台車第一筆加油紀錄沒有「上一次」可比，所以行駛 km 顯示 0、不計油耗。"
    fu["A3"].font = font(Z, size=9, italic=True, color="6B7280")

    # maintenance
    ITEMS = ["定期保養", "機油", "機油芯", "齒輪油", "變速箱油", "煞車油", "冷卻水", "空氣濾清器", "冷氣濾網", "火星塞",
             "煞車皮", "輪胎", "電瓶", "雨刷", "皮帶", "大燈／燈泡", "維修", "美容／洗車", "其他"]
    mt = wb.create_sheet("保養維修")
    title(mt, "保養維修紀錄", "「項目」請用下拉選單，保養提醒會自動抓最後一次紀錄。", cjk=Z)
    header_row(mt, 4, ["日期", "車輛", "里程表（km）", "項目", "店家", "金額", "備註"], cjk=Z)
    M1, M2 = 5, 404
    mts = [(D(2026, 9, 12), "家用車", 38900, "定期保養", "（自填店家）", 3200, "含機油、機油芯"),
           (D(2026, 9, 12), "家用車", 38900, "機油", "（自填店家）", 0, "含在定保內"),
           (D(2026, 5, 2), "家用車", 33800, "輪胎", "（自填店家）", 12800, "換四輪"),
           (D(2026, 9, 20), "通勤機車", 11300, "機油", "（自填店家）", 350, ""),
           (D(2026, 6, 8), "通勤機車", 9200, "齒輪油", "（自填店家）", 150, ""),
           (D(2025, 12, 25), "家用車", 32100, "雨刷", "（自填店家）", 600, "")]
    for i, s in enumerate(mts):
        for j, v in enumerate(s):
            mt.cell(row=M1 + i, column=1 + j, value=v)
    grid(mt, M1, M2, 1, 7, cjk=Z, input_cols=range(1, 8), fmt={1: DATE, 3: "#,##0", 6: "#,##0"})
    dv_list(mt, f"={CARS}", f"B{M1}:B{M2}")
    mt["J4"] = "項目清單"; mt["J4"].font = font(Z, bold=True)
    for i, it in enumerate(ITEMS):
        mt.cell(row=5 + i, column=10, value=it).font = font(Z, size=9, color="6B7280")
    dv_list(mt, f"=$J$5:$J${4+len(ITEMS)}", f"D{M1}:D{M2}")
    widths(mt, {"A": 12, "B": 14, "C": 13, "D": 14, "E": 18, "F": 10, "G": 26, "H": 3, "I": 3, "J": 14})
    mt.freeze_panes = "A5"; landscape(mt)
    ITEMR = f"保養維修!$J$5:$J${4+len(ITEMS)}"

    # reminders
    rm = wb.create_sheet("保養提醒")
    title(rm, "保養提醒", "間隔可只填公里或只填月數；兩者都填時，先到者為準。", cjk=Z)
    rm["A3"] = "「快到了」門檻：剩餘公里 ≤"; rm["E3"] = 500; rm["F3"] = "或剩餘天數 ≤"; rm["G3"] = 30
    for a in ("E3", "G3"):
        rm[a].fill = INPUT_FILL; rm[a].border = BORDER
    rm["A3"].font = font(Z, size=10); rm["F3"].font = font(Z, size=10)
    header_row(rm, 4, ["車輛", "項目", "每幾公里", "每幾個月", "上次日期", "上次里程", "目前里程", "下次里程", "剩餘公里", "下次日期", "剩餘天數", "狀態"], cjk=Z)
    R1, R2 = 5, 34
    rms = [("家用車", "機油", 5000, 6), ("家用車", "定期保養", 10000, 12), ("家用車", "輪胎", 40000, 48),
           ("家用車", "雨刷", None, 12), ("通勤機車", "機油", 1000, 3), ("通勤機車", "齒輪油", 3000, 6)]
    for i, s in enumerate(rms):
        for j, v in enumerate(s):
            rm.cell(row=R1 + i, column=1 + j, value=v)
    MB = f"保養維修!$B${M1}:$B${M2}"; MD = f"保養維修!$D${M1}:$D${M2}"; MA = f"保養維修!$A${M1}:$A${M2}"; MC = f"保養維修!$C${M1}:$C${M2}"
    FB = f"加油紀錄!$B${F1}:$B${F2}"; FC = f"加油紀錄!$C${F1}:$C${F2}"
    for r in range(R1, R2 + 1):
        rm.cell(row=r, column=5, value=f'=IF(OR(A{r}="",B{r}=""),"",IF(COUNTIFS({MB},A{r},{MD},B{r})=0,"尚無紀錄",_xlfn.MAXIFS({MA},{MB},A{r},{MD},B{r})))')
        rm.cell(row=r, column=6, value=f'=IF(OR(A{r}="",B{r}=""),"",IF(COUNTIFS({MB},A{r},{MD},B{r})=0,"",_xlfn.MAXIFS({MC},{MB},A{r},{MD},B{r})))')
        rm.cell(row=r, column=7, value=f'=IF(A{r}="","",MAX(_xlfn.MAXIFS({FC},{FB},A{r}),_xlfn.MAXIFS({MC},{MB},A{r})))')
        rm.cell(row=r, column=8, value=f'=IF(OR(F{r}="",N(C{r})=0),"",F{r}+C{r})')
        rm.cell(row=r, column=9, value=f'=IF(H{r}="","",H{r}-G{r})')
        rm.cell(row=r, column=10, value=f'=IF(OR(NOT(ISNUMBER(E{r})),N(D{r})=0),"",EDATE(E{r},D{r}))')
        rm.cell(row=r, column=11, value=f'=IF(J{r}="","",J{r}-TODAY())')
        rm.cell(row=r, column=12, value=(f'=IF(OR(A{r}="",B{r}=""),"",IF(E{r}="尚無紀錄","請先記錄一次",'
                                          f'IF(OR(AND(I{r}<>"",N(I{r})<=0),AND(K{r}<>"",N(K{r})<=0)),"該保養了",'
                                          f'IF(OR(AND(I{r}<>"",N(I{r})<=$E$3),AND(K{r}<>"",N(K{r})<=$G$3)),"快到了","正常"))))'))
    grid(rm, R1, R2, 1, 12, cjk=Z, input_cols=(1, 2, 3, 4), fmt={5: DATE, 6: "#,##0", 7: "#,##0", 8: "#,##0", 9: "#,##0;[Red]-#,##0", 10: DATE, 11: "0;[Red]-0"})
    dv_list(rm, f"={CARS}", f"A{R1}:A{R2}")
    dv_list(rm, f"={ITEMR}", f"B{R1}:B{R2}")
    rm.conditional_formatting.add(f"L{R1}:L{R2}", FormulaRule(formula=[f'L{R1}="該保養了"'], fill=RED_FILL))
    rm.conditional_formatting.add(f"L{R1}:L{R2}", FormulaRule(formula=[f'L{R1}="快到了"'], fill=AMBER_FILL))
    rm.conditional_formatting.add(f"L{R1}:L{R2}", FormulaRule(formula=[f'L{R1}="正常"'], fill=GREEN_FILL))
    widths(rm, {"A": 14, "B": 13, "C": 10, "D": 10, "E": 12, "F": 11, "G": 11, "H": 11, "I": 10, "J": 12, "K": 10, "L": 13})
    rm.freeze_panes = "A5"; landscape(rm)

    # fees
    fe = wb.create_sheet("稅費保險到期")
    title(fe, "稅費保險到期", "填到期日（或繳費截止日）。繳完後把日期改成下一期即可。", cjk=Z)
    header_row(fe, 4, ["車輛", "項目", "到期／截止日", "金額", "已繳日期", "剩餘天數", "狀態", "備註"], cjk=Z)
    T1, T2 = 5, 34
    fes = [("家用車", "汽車強制險", D(2027, 3, 9), 0, None, "金額依保單填寫"),
           ("家用車", "任意險（車體／第三人）", D(2027, 3, 9), 15800, D(2026, 3, 5), "範例金額，請依保單填寫"),
           ("家用車", "牌照稅", None, 0, None, "依繳款書填寫截止日"),
           ("家用車", "燃料費", None, 0, None, "依繳款書填寫截止日"),
           ("通勤機車", "機車強制險", D(2027, 6, 30), 0, None, ""),
           ("通勤機車", "排氣定檢", D(2027, 7, 31), 0, None, "依通知單填寫")]
    for i, s in enumerate(fes):
        for j, v in enumerate(s):
            if j == 5:
                fe.cell(row=T1 + i, column=8, value=v)
            elif j == 4:
                fe.cell(row=T1 + i, column=5, value=v)
            else:
                fe.cell(row=T1 + i, column=1 + j, value=v)
    for r in range(T1, T2 + 1):
        fe.cell(row=r, column=6, value=f'=IF(C{r}="","",C{r}-TODAY())')
        fe.cell(row=r, column=7, value=f'=IF(B{r}="","",IF(C{r}="","請填日期",IF(AND(E{r}<>"",E{r}>=C{r}-365),"已繳",IF(F{r}<0,"已逾期",IF(F{r}<=30,"30 天內到期","未到期")))))')
    grid(fe, T1, T2, 1, 8, cjk=Z, input_cols=(1, 2, 3, 4, 5, 8), fmt={3: DATE, 4: "#,##0", 5: DATE, 6: "0;[Red]-0"})
    dv_list(fe, f"={CARS}", f"A{T1}:A{T2}")
    fe.conditional_formatting.add(f"G{T1}:G{T2}", FormulaRule(formula=[f'G{T1}="已逾期"'], fill=RED_FILL))
    fe.conditional_formatting.add(f"G{T1}:G{T2}", FormulaRule(formula=[f'G{T1}="30 天內到期"'], fill=AMBER_FILL))
    fe.conditional_formatting.add(f"G{T1}:G{T2}", FormulaRule(formula=[f'G{T1}="已繳"'], fill=GREEN_FILL))
    widths(fe, {"A": 14, "B": 24, "C": 14, "D": 10, "E": 12, "F": 10, "G": 13, "H": 26})
    fe.freeze_panes = "A5"; landscape(fe)

    # overview
    ov = wb.create_sheet("用車成本總覽")
    title(ov, "用車成本總覽", None, cjk=Z)
    ov["A3"] = "統計年度"; ov["B3"] = 2026; ov["B3"].fill = INPUT_FILL; ov["B3"].border = BORDER
    ov["A3"].font = font(Z, bold=True)
    YS = "DATE($B$3,1,1)"; YE = "DATE($B$3,12,31)"
    FA = f"加油紀錄!$A${F1}:$A${F2}"; FF = f"加油紀錄!$F${F1}:$F${F2}"; FH = f"加油紀錄!$H${F1}:$H${F2}"; FD = f"加油紀錄!$D${F1}:$D${F2}"
    MF = f"保養維修!$F${M1}:$F${M2}"
    TA = f"稅費保險到期!$A${T1}:$A${T2}"; TD = f"稅費保險到期!$D${T1}:$D${T2}"; TE = f"稅費保險到期!$E${T1}:$E${T2}"
    yr = lambda rng: f'{rng},">="&{YS},{rng},"<="&{YE}'
    header_row(ov, 5, ["車輛", "油錢", "保養維修", "稅費保險（已繳）", "總花費", "行駛里程 km", "平均油耗 km/L", "每公里成本"], cjk=Z)
    for i in range(6):
        r = 6 + i
        ov.cell(row=r, column=1, value=f'=IF(車輛資料!A{5+i}="","",車輛資料!A{5+i})')
        ov.cell(row=r, column=2, value=f'=IF(A{r}="","",SUMIFS({FF},{FB},A{r},{yr(FA)}))')
        ov.cell(row=r, column=3, value=f'=IF(A{r}="","",SUMIFS({MF},{MB},A{r},{yr(MA)}))')
        ov.cell(row=r, column=4, value=f'=IF(A{r}="","",SUMIFS({TD},{TA},A{r},{yr(TE)}))')
        ov.cell(row=r, column=5, value=f'=IF(A{r}="","",B{r}+C{r}+D{r})')
        ov.cell(row=r, column=6, value=f'=IF(A{r}="","",SUMIFS({FH},{FB},A{r},{yr(FA)}))')
        ov.cell(row=r, column=7, value=f'=IF(A{r}="","",IFERROR(ROUND(F{r}/SUMIFS({FD},{FB},A{r},{FH},">0",{yr(FA)}),1),"-"))')
        ov.cell(row=r, column=8, value=f'=IF(A{r}="","",IF(N(F{r})=0,"-",ROUND(E{r}/F{r},2)))')
    grid(ov, 6, 11, 1, 8, cjk=Z, fmt={2: "#,##0", 3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "0.0", 8: "0.00"})
    ov.cell(row=12, column=1, value="合計").font = font(Z, bold=True)
    for c in range(2, 7):
        cc = ov.cell(row=12, column=c, value=f"=SUM({CL(c)}6:{CL(c)}11)"); cc.number_format = "#,##0"; cc.font = font(Z, bold=True); cc.fill = GREY_FILL; cc.border = BORDER
    ov["A14"] = '=IF(COUNTIF(保養提醒!$L$5:$L$34,"該保養了")+COUNTIF(稅費保險到期!$G$5:$G$34,"已逾期")>0,"⚠ 有 "&COUNTIF(保養提醒!$L$5:$L$34,"該保養了")&" 項保養已到期、"&COUNTIF(稅費保險到期!$G$5:$G$34,"已逾期")&" 項稅費保險已逾期，請查看對應工作表。","")'
    ov["A14"].font = font(Z, size=11, bold=True, color="9B1C1C")
    ov["A15"] = '="即將到期：保養 "&COUNTIF(保養提醒!$L$5:$L$34,"快到了")&" 項、稅費保險 "&COUNTIF(稅費保險到期!$G$5:$G$34,"30 天內到期")&" 項"'
    ov["A15"].font = font(Z, size=10, color=DARK)
    header_row(ov, 17, ["月份", "油錢", "保養維修", "合計"], cjk=Z)
    for m in range(1, 13):
        r = 17 + m
        ms = f"DATE($B$3,{m},1)"; me = f"EOMONTH({ms},0)"
        ov.cell(row=r, column=1, value=f"{m} 月")
        ov.cell(row=r, column=2, value=f'=SUMIFS({FF},{FA},">="&{ms},{FA},"<="&{me})')
        ov.cell(row=r, column=3, value=f'=SUMIFS({MF},{MA},">="&{ms},{MA},"<="&{me})')
        ov.cell(row=r, column=4, value=f"=B{r}+C{r}")
    grid(ov, 18, 29, 1, 4, cjk=Z, fmt={2: "#,##0", 3: "#,##0", 4: "#,##0"})
    ch = BarChart(); ch.type = "col"; ch.grouping = "stacked"; ch.overlap = 100; ch.title = "每月用車花費"; ch.style = 10
    ch.add_data(Reference(ov, min_col=2, max_col=3, min_row=17, max_row=29), titles_from_data=True)
    ch.set_categories(Reference(ov, min_col=1, min_row=18, max_row=29))
    ch.height = 7.5; ch.width = 14
    ch.series[0].graphicalProperties.solidFill = "F2C14E"; ch.series[1].graphicalProperties.solidFill = "2F3E46"
    ov.add_chart(ch, "F17")
    widths(ov, {"A": 14, "B": 11, "C": 11, "D": 18, "E": 11, "F": 12, "G": 15, "H": 11})
    landscape(ov); ov.page_setup.fitToHeight = 1
    wb.move_sheet("用車成本總覽", offset=-5)
    wb.active = 1
    wb.save(f"{OUT}/car-scooter-maintenance-fuel-log-tw.xlsx")


rental(); carlog()
print("ok")
