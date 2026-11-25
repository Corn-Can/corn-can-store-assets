import sys, datetime as dt
sys.path.insert(0, "/home/user/corn-can-store-assets/tools")
from common import *
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule, CellIsRule
from openpyxl.chart import BarChart, Reference, PieChart
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
    b.fill = PatternFill("solid", fgColor=LIGHT)
    b.alignment = Alignment(horizontal="center", vertical="center")
    b.number_format = fmtc
    a.border = b.border = BORDER
    ws.row_dimensions[r + 1].height = 32


# =====================================================================
# 1) Reading Tracker
# =====================================================================
def reading():
    wb = Workbook()
    hw = wb.active
    hw.title = "How to Use"
    st = wb.create_sheet("Setup")
    lg = wb.create_sheet("Book Log")
    tb = wb.create_sheet("TBR List")
    db = wb.create_sheet("Dashboard")
    N = 300

    howto(hw, "Reading Tracker & Book Log – How to Use", [
        "## What's inside (5 sheets)",
        "How to Use – this page.",
        "Setup – your reading year, yearly book goal, page goal, and the lists used by the drop-down menus (genres, formats).",
        "Book Log – one row per book: title, author, genre, format, status, start & finish dates, pages and your 1–5 rating. Days to read, pages per day and finish month are calculated for you.",
        "TBR List – your to-be-read wishlist with priority, source and a 'Started?' check that looks the title up in the Book Log.",
        "Dashboard – books finished vs goal, pages read, average rating, pace (ahead/behind), projected books by year-end, books & pages per month, and a genre breakdown with charts.",
        "## Quick start",
        "1. On Setup, type your reading year (e.g. 2027) and your goals. Everything else follows that year.",
        "2. Delete the example rows on Book Log and TBR List (select the rows → right-click → Delete), or just type over them.",
        "3. Add books on Book Log. Light-yellow cells are inputs; white cells are formulas – please don't type over them.",
        "4. A book counts toward your goal when Status = Finished and the Finish date is inside your reading year.",
        "5. Open the Dashboard any time to see your progress. Pace compares your finished books with where you'd be if you read evenly through the year.",
        "## Tips",
        "Add your own genres or formats in the lists on Setup – drop-downs update automatically (up to 20 genres / 8 formats).",
        "Use DNF (did not finish) for books you abandon – they don't count toward the goal but stay in your history.",
        "Re-reads: just add the book again with new dates.",
        "## Works with",
        "Microsoft Excel (2010 or later, Windows/Mac), Google Sheets (File → Import) and LibreOffice Calc. No macros.",
        "Made by Corn_Can. For personal use.",
    ])

    # Setup
    title(st, "Setup", "Type your year and goals in the yellow cells.")
    rows = [("Reading year", 2027, "0"), ("Book goal for the year", 30, "0"), ("Page goal for the year", 9000, "#,##0")]
    for i, (k, v, f) in enumerate(rows):
        a = st.cell(row=4 + i, column=1, value=k); a.font = font(bold=True); a.border = BORDER
        b = st.cell(row=4 + i, column=2, value=v); b.fill = INPUT_FILL; b.border = BORDER; b.number_format = f
        b.font = font(size=12, bold=True)
    header_row(st, 9, ["Genres"], start_col=1)
    header_row(st, 9, ["Formats"], start_col=3)
    header_row(st, 9, ["Statuses"], start_col=5)
    genres = ["Fiction", "Mystery & Thriller", "Romance", "Fantasy", "Sci-Fi", "Historical Fiction", "Non-fiction",
              "Biography & Memoir", "Self-help", "Business", "Science", "History", "Poetry", "Classics", "Young Adult"]
    formats = ["Paperback", "Hardcover", "E-book", "Audiobook", "Library"]
    statuses = ["TBR", "Reading", "Finished", "DNF"]
    for i in range(20):
        c = st.cell(row=10 + i, column=1, value=genres[i] if i < len(genres) else None); c.fill = INPUT_FILL; c.border = BORDER
    for i in range(8):
        c = st.cell(row=10 + i, column=3, value=formats[i] if i < len(formats) else None); c.fill = INPUT_FILL; c.border = BORDER
    for i, s in enumerate(statuses):
        c = st.cell(row=10 + i, column=5, value=s); c.border = BORDER
    widths(st, {"A": 26, "B": 14, "C": 16, "D": 3, "E": 14})
    landscape(st)

    # Book Log
    title(lg, "Book Log", "One row per book. Yellow = input, white = calculated.")
    hdr = ["Title", "Author", "Genre", "Format", "Status", "Start date", "Finish date", "Pages", "Rating (1–5)",
           "Stars", "Days to read", "Pages / day", "Finish month", "Notes / favorite quote"]
    header_row(lg, 4, hdr)
    ex = [
        ("The Quiet Orchard", "M. Hale", "Fiction", "Paperback", "Finished", D(2027, 1, 2), D(2027, 1, 11), 320, 4, "Cozy, slow-burn family story"),
        ("Small Steps, Big Focus", "J. Rivera", "Self-help", "E-book", "Finished", D(2027, 1, 12), D(2027, 1, 20), 256, 3, ""),
        ("The Last Ferry Home", "A. Novak", "Mystery & Thriller", "Audiobook", "Finished", D(2027, 1, 21), D(2027, 1, 30), 410, 5, "Couldn't stop listening"),
        ("Stars Over Kepler", "R. Osei", "Sci-Fi", "Hardcover", "Reading", D(2027, 2, 1), None, 480, None, ""),
        ("A Short History of Bread", "L. Moreau", "History", "Library", "DNF", D(2027, 1, 5), None, 300, 2, "Stopped at p. 90"),
        ("The Lighthouse Letters", "K. Byrne", "Romance", "Paperback", "TBR", None, None, 352, None, ""),
    ]
    for i, e in enumerate(ex):
        r = 5 + i
        for j, v in enumerate(e[:9]):
            lg.cell(row=r, column=j + 1, value=v)
        lg.cell(row=r, column=14, value=e[9])
    for r in range(5, 5 + N):
        lg.cell(row=r, column=10, value=f'=IF(ISNUMBER(I{r}),REPT("★",I{r})&REPT("☆",5-I{r}),"")')
        lg.cell(row=r, column=11, value=f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(G{r})),G{r}-F{r}+1,"")')
        lg.cell(row=r, column=12, value=f'=IF(AND(ISNUMBER(K{r}),ISNUMBER(H{r})),ROUND(H{r}/K{r},0),"")')
        lg.cell(row=r, column=13, value=f'=IF(AND(E{r}="Finished",ISNUMBER(G{r})),TEXT(G{r},"mmm yyyy"),"")')
    grid(lg, 5, 4 + N, 1, 14, input_cols=(1, 2, 3, 4, 5, 6, 7, 8, 9, 14),
         fmt={6: "yyyy-mm-dd", 7: "yyyy-mm-dd", 8: "#,##0", 11: "0", 12: "0"})
    for r in range(5, 5 + N):
        lg.cell(row=r, column=10).font = font(size=11, color="C98A00")
    dv_list(lg, "=Setup!$A$10:$A$29", f"C5:C{4+N}")
    dv_list(lg, "=Setup!$C$10:$C$17", f"D5:D{4+N}")
    dv_list(lg, '"TBR,Reading,Finished,DNF"', f"E5:E{4+N}")
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True,
                        showErrorMessage=True, error="Please enter a whole number from 1 to 5.")
    lg.add_data_validation(dv); dv.add(f"I5:I{4+N}")
    rng = f"A5:N{4+N}"
    lg.conditional_formatting.add(rng, FormulaRule(formula=['$E5="Finished"'], fill=GREEN_FILL))
    lg.conditional_formatting.add(rng, FormulaRule(formula=['$E5="Reading"'], fill=BLUE_FILL))
    lg.conditional_formatting.add(rng, FormulaRule(formula=['$E5="DNF"'], fill=GREY_FILL))
    lg.freeze_panes = "B5"
    widths(lg, {"A": 30, "B": 16, "C": 18, "D": 12, "E": 11, "F": 12, "G": 12, "H": 8, "I": 10, "J": 11,
                "K": 9, "L": 9, "M": 12, "N": 30})
    lg.auto_filter.ref = f"A4:N{4+N}"
    landscape(lg)
    lg.print_area = "A1:N30"

    # TBR
    title(tb, "TBR List (To Be Read)", "Your wishlist. 'Started?' checks whether the title is already on the Book Log.")
    header_row(tb, 4, ["Title", "Author", "Genre", "Priority", "Where I heard about it", "Own it?", "Started?"])
    tex = [("The Lighthouse Letters", "K. Byrne", "Romance", "High", "Friend's recommendation", "Yes"),
           ("Deep Rivers", "T. Amari", "Non-fiction", "Medium", "Podcast", "No"),
           ("The Ninth Clockmaker", "E. Saito", "Fantasy", "High", "Book club pick", "No"),
           ("Small Wins at Work", "P. Grant", "Business", "Low", "Library display", "No")]
    for i, e in enumerate(tex):
        for j, v in enumerate(e):
            tb.cell(row=5 + i, column=j + 1, value=v)
    for r in range(5, 105):
        tb.cell(row=r, column=7, value=f'=IF(A{r}="","",IF(COUNTIFS(\'Book Log\'!$A$5:$A${4+N},A{r},\'Book Log\'!$E$5:$E${4+N},"<>TBR")>0,"Yes","Not yet"))')
    grid(tb, 5, 104, 1, 7, input_cols=(1, 2, 3, 4, 5, 6))
    dv_list(tb, "=Setup!$A$10:$A$29", "C5:C104")
    dv_list(tb, '"High,Medium,Low"', "D5:D104")
    dv_list(tb, '"Yes,No"', "F5:F104")
    tb.conditional_formatting.add("A5:G104", FormulaRule(formula=['$D5="High"'], fill=AMBER_FILL))
    tb.conditional_formatting.add("A5:G104", FormulaRule(formula=['$G5="Yes"'], fill=GREEN_FILL))
    widths(tb, {"A": 30, "B": 18, "C": 18, "D": 10, "E": 26, "F": 9, "G": 10})
    tb.freeze_panes = "A5"
    landscape(tb)
    tb.print_area = "A1:G24"

    # Dashboard
    L = "'Book Log'!"
    S, G, P = f"{L}$E$5:$E${4+N}", f"{L}$G$5:$G${4+N}", f"{L}$H$5:$H${4+N}"
    Y = "Setup!$B$4"
    inyear = f'{S},"Finished",{G},">="&DATE({Y},1,1),{G},"<="&DATE({Y},12,31)'
    title(db, "Reading Dashboard", None)
    db["A2"] = '="Reading year: "&Setup!B4&"   ·   Goal: "&Setup!B5&" books / "&TEXT(Setup!B6,"#,##0")&" pages"'
    db["A2"].font = font(size=11, italic=True, color="6B7280")
    # helper cells
    db["J3"] = "Days in year"; db["K3"] = f"=DATE({Y},12,31)-DATE({Y},1,1)+1"
    db["J4"] = "Days elapsed"; db["K4"] = f"=MAX(0,MIN(K3,TODAY()-DATE({Y},1,1)+1))"
    for c in ("J3", "J4", "K3", "K4"):
        db[c].font = font(size=9, color="9CA3AF")
    kpi(db, 4, 1, "Books finished", f"=COUNTIFS({inyear})")
    kpi(db, 4, 2, "Goal progress", "=IF(Setup!B5>0,A5/Setup!B5,0)", "0%")
    kpi(db, 4, 3, "Pages read", f"=SUMIFS({P},{inyear})", "#,##0")
    kpi(db, 4, 4, "Page goal", "=IF(Setup!B6>0,C5/Setup!B6,0)", "0%")
    kpi(db, 4, 5, "Avg rating", f'=IFERROR(AVERAGEIFS({L}$I$5:$I${4+N},{inyear}),0)', "0.0")
    kpi(db, 4, 6, "Currently reading", f'=COUNTIF({S},"Reading")')
    kpi(db, 7, 1, "On-pace target today", "=ROUND(Setup!B5*K4/K3,1)", "0.0")
    kpi(db, 7, 2, "Ahead / behind", '=IF(K4=0,"Not started",IF(A5-A8>=0,"Ahead by "&TEXT(A5-A8,"0.0"),"Behind by "&TEXT(A8-A5,"0.0")))', "@")
    kpi(db, 7, 3, "Projected by year-end", "=IF(K4=0,0,ROUND(A5/K4*K3,0))")
    kpi(db, 7, 4, "Books left", "=MAX(0,Setup!B5-A5)")
    kpi(db, 7, 5, "Needed per month", '=IF(K4>=K3,0,ROUND(D8/((K3-K4)/30.4),1))', "0.0")
    kpi(db, 7, 6, "DNF (all time)", f'=COUNTIF({S},"DNF")')
    db.conditional_formatting.add("B8", FormulaRule(formula=['LEFT(B8,6)="Behind"'], fill=RED_FILL))
    db.conditional_formatting.add("B8", FormulaRule(formula=['LEFT(B8,5)="Ahead"'], fill=GREEN_FILL))
    db.conditional_formatting.add("B5", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F2C14E"))

    header_row(db, 11, ["Month", "Books finished", "Pages read", "Avg rating"])
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for i, m in enumerate(months):
        r = 12 + i
        db.cell(row=r, column=1, value=m)
        rng_m = f'{S},"Finished",{G},">="&DATE({Y},{i+1},1),{G},"<"&DATE({Y},{i+2},1)' if i < 11 else \
            f'{S},"Finished",{G},">="&DATE({Y},12,1),{G},"<="&DATE({Y},12,31)'
        db.cell(row=r, column=2, value=f"=COUNTIFS({rng_m})")
        db.cell(row=r, column=3, value=f"=SUMIFS({P},{rng_m})")
        db.cell(row=r, column=4, value=f'=IFERROR(AVERAGEIFS({L}$I$5:$I${4+N},{rng_m}),"")')
    db.cell(row=24, column=1, value="Total"); db.cell(row=24, column=2, value="=SUM(B12:B23)")
    db.cell(row=24, column=3, value="=SUM(C12:C23)"); db.cell(row=24, column=4, value="=E5")
    grid(db, 12, 24, 1, 4, fmt={3: "#,##0", 4: "0.0"})
    for c in range(1, 5):
        db.cell(row=24, column=c).font = font(bold=True)

    header_row(db, 11, ["Genre", "Books finished", "Pages read"], start_col=6)
    for i in range(20):
        r = 12 + i
        db.cell(row=r, column=6, value=f'=IF(Setup!A{10+i}="","",Setup!A{10+i})')
        db.cell(row=r, column=7, value=f'=IF(F{r}="","",COUNTIFS({inyear},{L}$C$5:$C${4+N},F{r}))')
        db.cell(row=r, column=8, value=f'=IF(F{r}="","",SUMIFS({P},{inyear},{L}$C$5:$C${4+N},F{r}))')
    grid(db, 12, 31, 6, 8, fmt={8: "#,##0"})
    db.conditional_formatting.add("G12:G31", DataBarRule(start_type="num", start_value=0, end_type="max", color="F2C14E"))

    ch = BarChart(); ch.type = "col"; ch.title = "Books finished per month"; ch.style = 10
    ch.add_data(Reference(db, min_col=2, min_row=11, max_row=23), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=1, min_row=12, max_row=23))
    ch.legend = None; ch.height = 7; ch.width = 15
    ch.series[0].graphicalProperties.solidFill = "F2C14E"
    db.add_chart(ch, "A33")
    widths(db, {"A": 18, "B": 18, "C": 18, "D": 18, "E": 18, "F": 22, "G": 15, "H": 12, "J": 12})
    landscape(db); db.page_setup.fitToHeight = 1
    db.print_area = "A1:H48"
    wb.move_sheet("Dashboard", offset=-3)
    wb.active = 1
    wb.save(f"{OUT}/reading-tracker-book-log.xlsx")


# =====================================================================
# 2) Freelance Invoice & Client Payment Tracker
# =====================================================================
def freelance():
    wb = Workbook()
    hw = wb.active; hw.title = "How to Use"
    se = wb.create_sheet("Settings")
    cl = wb.create_sheet("Clients")
    iv = wb.create_sheet("Invoices")
    db = wb.create_sheet("Dashboard")
    N = 300
    howto(hw, "Freelance Invoice & Client Payment Tracker – How to Use", [
        "## What's inside (5 sheets)",
        "How to Use – this page.",
        "Settings – your business name, report year, 'due soon' warning window, and an optional set-aside % you choose for taxes or savings.",
        "Clients – your client list with default payment terms (e.g. 14 or 30 days) and an automatic summary per client: invoiced, paid, outstanding, overdue and average days to pay.",
        "Invoices – one row per invoice. Pick the client and the due date fills in from their payment terms. Status updates automatically: Paid, Partially Paid, Overdue, Due Soon or Open.",
        "Dashboard – invoiced, collected, outstanding and overdue totals for the report year, monthly invoiced vs collected table and chart, top clients, and your chosen set-aside amount.",
        "## Quick start",
        "1. On Settings, type your business name and the report year. Set the 'due soon' window (default 7 days).",
        "2. Replace the example clients on Clients. Default terms = days a client has to pay (0 = due on receipt).",
        "3. On Invoices, add a row for each invoice you send: invoice #, client, project, issue date and amount. Leave 'Custom terms' (days) blank to use the client's default terms.",
        "4. When money arrives, type the total received so far in 'Amount paid' and the date in 'Paid date'. Status and balance update automatically.",
        "5. Check the Dashboard each week to see who owes you and what's overdue.",
        "## Notes",
        "Light-yellow cells are inputs; white cells are formulas – please don't type over them.",
        "This is a tracking tool, not an invoice generator or accounting software. It does not send invoices or reminders.",
        "Set-aside %: enter any percentage you personally want to reserve (for example, for taxes). The template does not calculate taxes and is not tax advice – check your local rules or an accountant.",
        "Amounts use a plain number format so any currency works. Monthly 'Collected' uses the Paid date of each invoice.",
        "## Works with",
        "Microsoft Excel (2010 or later, Windows/Mac), Google Sheets (File → Import) and LibreOffice Calc. No macros.",
        "Made by Corn_Can.",
    ])

    title(se, "Settings", "Yellow cells are inputs.")
    items = [("Business name", "Your Studio Name", "@"), ("Report year", "=YEAR(TODAY())", "0"),
             ("'Due soon' window (days)", 7, "0"), ("Set-aside % (your choice)", 0.25, "0%"),
             ("Default payment terms (days)", 30, "0")]
    for i, (k, v, f) in enumerate(items):
        a = se.cell(row=4 + i, column=1, value=k); a.font = font(bold=True); a.border = BORDER
        b = se.cell(row=4 + i, column=2, value=v); b.fill = INPUT_FILL; b.border = BORDER; b.number_format = f
        b.font = font(size=12, bold=True)
    se["C5"] = "Defaults to the current year – type a year (e.g. 2027) to fix it."
    se["C7"] = "Optional. Not tax advice – use the % you decide on."
    se["C8"] = "Used when a client has no default terms."
    for c in ("C5", "C7", "C8"):
        se[c].font = font(size=9, italic=True, color="6B7280")
    widths(se, {"A": 30, "B": 20, "C": 60})
    landscape(se)

    # Clients
    title(cl, "Clients", "Client list + automatic payment summary (all years).")
    header_row(cl, 4, ["Client", "Contact person", "Email", "Default terms (days)", "Invoices", "Total invoiced",
                       "Total paid", "Outstanding", "Overdue amount", "Avg days to pay"])
    cex = [("Bright Pixel Co.", "Dana Lee", "dana@example.com", 30), ("Northwind Cafe", "Sam Ortiz", "sam@example.com", 14),
           ("Greenleaf Studio", "Priya N.", "priya@example.com", 30), ("Harbor Realty", "Tom W.", "tom@example.com", 0)]
    I = "Invoices!"
    for i, e in enumerate(cex):
        for j, v in enumerate(e):
            cl.cell(row=5 + i, column=j + 1, value=v)
    for r in range(5, 55):
        cl.cell(row=r, column=5, value=f'=IF(A{r}="","",COUNTIF({I}$B$6:$B${5+N},A{r}))')
        cl.cell(row=r, column=6, value=f'=IF(A{r}="","",SUMIF({I}$B$6:$B${5+N},A{r},{I}$G$6:$G${5+N}))')
        cl.cell(row=r, column=7, value=f'=IF(A{r}="","",SUMIF({I}$B$6:$B${5+N},A{r},{I}$H$6:$H${5+N}))')
        cl.cell(row=r, column=8, value=f'=IF(A{r}="","",F{r}-G{r})')
        cl.cell(row=r, column=9, value=f'=IF(A{r}="","",SUMIFS({I}$I$6:$I${5+N},{I}$B$6:$B${5+N},A{r},{I}$K$6:$K${5+N},"Overdue"))')
        cl.cell(row=r, column=10, value=f'=IF(A{r}="","",IFERROR(AVERAGEIFS({I}$M$6:$M${5+N},{I}$B$6:$B${5+N},A{r},{I}$K$6:$K${5+N},"Paid"),"–"))')
    grid(cl, 5, 54, 1, 10, input_cols=(1, 2, 3, 4), fmt={6: "#,##0.00", 7: "#,##0.00", 8: "#,##0.00", 9: "#,##0.00", 10: "0"})
    cl.conditional_formatting.add("I5:I54", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
    widths(cl, {"A": 22, "B": 16, "C": 22, "D": 12, "E": 9, "F": 14, "G": 14, "H": 14, "I": 14, "J": 11})
    cl.freeze_panes = "B5"
    landscape(cl); cl.print_area = "A1:J20"

    # Invoices
    title(iv, "Invoices", None)
    iv["A2"] = '="Today: "&TEXT(TODAY(),"yyyy-mm-dd")&"   ·   Due-soon window: "&Settings!B6&" days"'
    iv["A2"].font = font(size=10, italic=True, color="6B7280")
    header_row(iv, 5, ["Invoice #", "Client", "Project / description", "Issue date", "Custom terms", "Due date",
                       "Amount", "Amount paid", "Balance", "Paid date", "Status", "Days overdue", "Days to pay", "Notes"])
    today = dt.date.today()
    T = lambda d: today + dt.timedelta(days=d)
    vex = [("INV-001", "Bright Pixel Co.", "Logo refresh", T(-70), None, 1200, 1200, T(-45), ""),
           ("INV-002", "Northwind Cafe", "Menu design", T(-40), None, 450, 450, T(-30), ""),
           ("INV-003", "Greenleaf Studio", "Website copy – 5 pages", T(-38), None, 900, 400, None, "Deposit received"),
           ("INV-004", "Harbor Realty", "Listing photos", T(-20), None, 350, 0, None, "Sent reminder"),
           ("INV-005", "Bright Pixel Co.", "Social media kit", T(-25), None, 750, 0, None, ""),
           ("INV-006", "Northwind Cafe", "Holiday flyer", T(-3), 7, 200, 0, None, "")]
    for i, e in enumerate(vex):
        r = 6 + i
        for col, v in zip((1, 2, 3, 4, 5, 7, 8, 10, 14), e):
            iv.cell(row=r, column=col, value=v)
    for r in range(6, 6 + N):
        iv.cell(row=r, column=6, value=f'=IF(OR(B{r}="",NOT(ISNUMBER(D{r}))),"",D{r}+IF(ISNUMBER(E{r}),E{r},IFERROR(IF(VLOOKUP(B{r},Clients!$A$5:$D$54,4,FALSE)="",Settings!$B$8,VLOOKUP(B{r},Clients!$A$5:$D$54,4,FALSE)),Settings!$B$8)))')
        iv.cell(row=r, column=9, value=f'=IF(G{r}="","",G{r}-N(H{r}))')
        iv.cell(row=r, column=11, value=f'=IF(G{r}="","",IF(I{r}<=0,"Paid",IF(AND(ISNUMBER(F{r}),TODAY()>F{r}),"Overdue",IF(N(H{r})>0,"Partially Paid",IF(AND(ISNUMBER(F{r}),F{r}-TODAY()<=Settings!$B$6),"Due Soon","Open")))))')
        iv.cell(row=r, column=12, value=f'=IF(K{r}="Overdue",TODAY()-F{r},"")')
        iv.cell(row=r, column=13, value=f'=IF(AND(K{r}="Paid",ISNUMBER(J{r}),ISNUMBER(D{r})),J{r}-D{r},"")')
    grid(iv, 6, 5 + N, 1, 14, input_cols=(1, 2, 3, 4, 5, 7, 8, 10, 14),
         fmt={4: "yyyy-mm-dd", 6: "yyyy-mm-dd", 7: "#,##0.00", 8: "#,##0.00", 9: "#,##0.00", 10: "yyyy-mm-dd", 12: "0", 13: "0"})
    dv_list(iv, "=Clients!$A$5:$A$54", f"B6:B{5+N}")
    rng = f"A6:N{5+N}"
    iv.conditional_formatting.add(rng, FormulaRule(formula=['$K6="Overdue"'], fill=RED_FILL))
    iv.conditional_formatting.add(rng, FormulaRule(formula=['$K6="Due Soon"'], fill=AMBER_FILL))
    iv.conditional_formatting.add(rng, FormulaRule(formula=['$K6="Paid"'], fill=GREEN_FILL))
    iv.conditional_formatting.add(rng, FormulaRule(formula=['$K6="Partially Paid"'], fill=BLUE_FILL))
    iv.freeze_panes = "C6"
    iv.auto_filter.ref = f"A5:N{5+N}"
    widths(iv, {"A": 10, "B": 18, "C": 24, "D": 11, "E": 10, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11,
                "K": 13, "L": 9, "M": 8, "N": 20})
    landscape(iv); iv.print_area = "A1:N30"

    # Dashboard
    Y = "Settings!$B$5"
    B, Dd, G_, H_, I_, J_, K_ = (f"{I}${c}$6:${c}${5+N}" for c in "BDGHIJK")
    yr = f'{Dd},">="&DATE({Y},1,1),{Dd},"<="&DATE({Y},12,31)'
    title(db, "Freelance Income Dashboard", None)
    db["A2"] = '=Settings!B4&"   ·   Report year: "&Settings!B5&"   ·   Today: "&TEXT(TODAY(),"yyyy-mm-dd")'
    db["A2"].font = font(size=11, italic=True, color="6B7280")
    kpi(db, 4, 1, "Invoiced (year)", f"=SUMIFS({G_},{yr})", "#,##0.00")
    kpi(db, 4, 2, "Collected (year)", f'=SUMIFS({H_},{J_},">="&DATE({Y},1,1),{J_},"<="&DATE({Y},12,31))', "#,##0.00")
    kpi(db, 4, 3, "Outstanding (all)", f'=SUMIF({I_},">0")', "#,##0.00")
    kpi(db, 4, 4, "Overdue (all)", f'=SUMIF({K_},"Overdue",{I_})', "#,##0.00")
    kpi(db, 4, 5, "Overdue invoices", f'=COUNTIF({K_},"Overdue")')
    kpi(db, 7, 1, "Invoices sent (year)", f"=COUNTIFS({yr})")
    kpi(db, 7, 2, "Due soon", f'=COUNTIF({K_},"Due Soon")')
    kpi(db, 7, 3, "Avg days to pay", f'=IFERROR(AVERAGEIF({K_},"Paid",{I}$M$6:$M${5+N}),0)', "0")
    kpi(db, 7, 4, '="Set-aside ("&TEXT(Settings!B7,"0%")&" of collected)"', "=B5*Settings!B7", "#,##0.00")
    kpi(db, 7, 5, "Collection rate (year)", f'=IFERROR(SUMIFS({H_},{yr})/A5,0)', "0%")
    db.conditional_formatting.add("D5:E5", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))

    header_row(db, 11, ["Month", "Invoiced", "Collected", "Difference"])
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for i, m in enumerate(months):
        r = 12 + i
        lo = f"DATE({Y},{i+1},1)"; hi = f"DATE({Y},{i+2},1)" if i < 11 else f"DATE({Y}+1,1,1)"
        db.cell(row=r, column=1, value=m)
        db.cell(row=r, column=2, value=f'=SUMIFS({G_},{Dd},">="&{lo},{Dd},"<"&{hi})')
        db.cell(row=r, column=3, value=f'=SUMIFS({H_},{J_},">="&{lo},{J_},"<"&{hi})')
        db.cell(row=r, column=4, value=f"=C{r}-B{r}")
    db.cell(row=24, column=1, value="Total")
    for c, L_ in ((2, "B"), (3, "C"), (4, "D")):
        db.cell(row=24, column=c, value=f"=SUM({L_}12:{L_}23)")
    grid(db, 12, 24, 1, 4, fmt={2: "#,##0.00", 3: "#,##0.00", 4: "#,##0.00"})
    for c in range(1, 5):
        db.cell(row=24, column=c).font = font(bold=True)

    header_row(db, 11, ["Who owes you (top 5 outstanding)", "Outstanding"], start_col=6)
    for k in range(5):
        r = 12 + k
        db.cell(row=r, column=7, value=f'=IFERROR(LARGE(Clients!$H$5:$H$54,{k+1}),"")')
        db.cell(row=r, column=6, value=f'=IF(OR(G{r}="",N(G{r})<=0),"",INDEX(Clients!$A$5:$A$54,MATCH(G{r},Clients!$H$5:$H$54,0)))')
    grid(db, 12, 16, 6, 7, fmt={7: "#,##0.00"})
    db["F18"] = "Note: if two clients owe exactly the same amount, the first one is shown twice – see the Clients sheet for the full list."
    db["F18"].font = font(size=8, italic=True, color="6B7280")
    db["F18"].alignment = Alignment(wrap_text=True, vertical="top")
    db.merge_cells("F18:G20")

    ch = BarChart(); ch.type = "col"; ch.title = "Invoiced vs collected by month"; ch.style = 10
    ch.add_data(Reference(db, min_col=2, max_col=3, min_row=11, max_row=23), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=1, min_row=12, max_row=23))
    ch.height = 7.5; ch.width = 17
    ch.series[0].graphicalProperties.solidFill = "2F3E46"
    ch.series[1].graphicalProperties.solidFill = "F2C14E"
    db.add_chart(ch, "A26")
    widths(db, {"A": 20, "B": 20, "C": 20, "D": 24, "E": 20, "F": 30, "G": 14})
    landscape(db); db.page_setup.fitToHeight = 1; db.print_area = "A1:G42"
    wb.move_sheet("Dashboard", offset=-3)
    wb.active = 1
    wb.save(f"{OUT}/freelance-invoice-payment-tracker.xlsx")


# =====================================================================
# 3) 年前大掃除＆斷捨離計畫表
# =====================================================================
def cleaning():
    Z = True
    wb = Workbook()
    hw = wb.active; hw.title = "使用說明"
    se = wb.create_sheet("設定")
    tk = wb.create_sheet("大掃除清單")
    dc = wb.create_sheet("斷捨離清單")
    sp = wb.create_sheet("清潔用品採購")
    db = wb.create_sheet("進度總覽")
    howto(hw, "年前大掃除＆斷捨離計畫表－使用說明", [
        "## 內含工作表（共 6 張）",
        "使用說明：本頁。",
        "設定：填入目標完成日（例如除夕前一天）與家庭成員名字，下拉選單會自動帶入。",
        "大掃除清單：已預先列好 40 項常見打掃工作（依區域、由上而下排列），可直接分配負責人、預計日期與預估時間；狀態改成「完成」整列變綠色，逾期未完成會變紅色。",
        "斷捨離清單：記錄要處理的物品與方式（丟棄、回收、捐贈、二手出售、保留），自動加總二手出售的預估與實際收入。",
        "清潔用品採購：列出需要買的清潔用品，自動計算小計與總花費，打勾表示已買。",
        "進度總覽：距離目標日倒數天數、整體完成率、各區域進度、每位家人的分工（任務數、預估時間、完成數），附圖表。",
        "## 快速開始",
        "1. 到「設定」填入目標完成日與家人名字（最多 8 位）。",
        "2. 到「大掃除清單」刪掉不需要的項目、補上自己家的項目，再用下拉選單選負責人並填預計日期。",
        "3. 每做完一項，把狀態改成「完成」並填完成日。",
        "4. 隨時打開「進度總覽」看還剩多少、誰還有哪些沒做。",
        "## 小提醒",
        "淺黃色格子是可輸入欄位；白色格子有公式，請勿覆蓋。",
        "常見打掃原則：先整理（斷捨離）再清潔；由上而下、由內而外、由乾到濕。",
        "清潔劑請依產品標示使用，不同清潔劑不要混用，並保持通風。",
        "表內預估時間只是範例，請依自家情況調整。",
        "## 適用軟體",
        "Microsoft Excel 2010 以上（Windows／Mac）、Google 試算表（檔案 → 匯入）、LibreOffice Calc。不含巨集。",
        "Corn_Can 製作，僅供個人使用。",
    ], cjk=Z)

    title(se, "設定", "淺黃色格子請自行填寫。", cjk=Z)
    a = se.cell(row=4, column=1, value="目標完成日（例如除夕前一天）"); a.font = font(Z, bold=True); a.border = BORDER
    b = se.cell(row=4, column=2, value=D(2027, 2, 4)); b.fill = INPUT_FILL; b.border = BORDER; b.number_format = "yyyy-mm-dd"
    b.font = font(Z, size=12, bold=True)
    se["C4"] = "範例日期，請依自家計畫修改。"
    se["C4"].font = font(Z, size=9, italic=True, color="6B7280")
    header_row(se, 6, ["家庭成員"], cjk=Z)
    header_row(se, 6, ["打掃區域"], cjk=Z, start_col=3)
    members = ["爸爸", "媽媽", "小明", "小美", "全家"]
    areas = ["全屋", "客廳", "臥室", "廚房", "浴室", "玄關", "陽台", "窗戶", "寢具織品", "其他"]
    for i in range(8):
        c = se.cell(row=7 + i, column=1, value=members[i] if i < len(members) else None); c.fill = INPUT_FILL; c.border = BORDER; c.font = font(Z)
    for i in range(12):
        c = se.cell(row=7 + i, column=3, value=areas[i] if i < len(areas) else None); c.fill = INPUT_FILL; c.border = BORDER; c.font = font(Z)
    widths(se, {"A": 30, "B": 16, "C": 16})
    landscape(se)

    # 大掃除清單
    N = 120
    title(tk, "大掃除清單", None, cjk=Z)
    tk["A2"] = '="目標完成日："&TEXT(設定!B4,"yyyy-mm-dd")&"　｜　今天："&TEXT(TODAY(),"yyyy-mm-dd")'
    tk["A2"].font = font(Z, size=10, color="6B7280")
    header_row(tk, 4, ["區域", "項目", "負責人", "預計日期", "預估分鐘", "狀態", "完成日", "提醒", "備註"], cjk=Z)
    tasks = [
        ("全屋", "斷捨離：清出過期、損壞、一年沒用的物品", 120),
        ("全屋", "天花板、牆角蜘蛛網與灰塵", 40),
        ("全屋", "燈具、吊扇擦拭", 40),
        ("全屋", "冷氣濾網拆洗", 30),
        ("全屋", "門框、門把、電燈開關擦拭消毒", 30),
        ("全屋", "地板吸塵＋拖地（最後做）", 60),
        ("客廳", "高櫃頂部、書櫃除塵", 30),
        ("客廳", "沙發吸塵、抱枕套換洗", 30),
        ("客廳", "電視、電器與遙控器擦拭", 20),
        ("客廳", "茶几、電視櫃收納整理", 30),
        ("臥室", "衣櫃整理、換季衣物收納", 90),
        ("臥室", "床底清潔", 30),
        ("臥室", "梳妝台、書桌整理", 30),
        ("寢具織品", "床單、被套、枕套換洗", 60),
        ("寢具織品", "厚外套、圍巾清洗或送洗", 60),
        ("寢具織品", "窗簾拆洗", 90),
        ("廚房", "抽油煙機與濾網去油", 60),
        ("廚房", "瓦斯爐／電磁爐清潔", 30),
        ("廚房", "冰箱清空、丟過期食品、擦拭層架", 60),
        ("廚房", "櫥櫃內外擦拭、調味料檢查效期", 45),
        ("廚房", "流理台、水槽與排水孔清潔", 30),
        ("廚房", "微波爐、電鍋、烤箱清潔", 30),
        ("廚房", "廚房牆面磁磚去油", 30),
        ("浴室", "磁磚縫與矽利康除霉", 45),
        ("浴室", "馬桶、洗手台清潔", 30),
        ("浴室", "浴室鏡面與水龍頭除水垢", 20),
        ("浴室", "排水孔、地漏清潔", 20),
        ("浴室", "抽風扇除塵", 15),
        ("玄關", "鞋櫃整理、丟掉不穿的鞋", 40),
        ("玄關", "大門內外擦拭", 20),
        ("陽台", "洗衣機槽清潔", 30),
        ("陽台", "陽台地板刷洗、排水口清理", 40),
        ("陽台", "植物修剪、雜物整理", 30),
        ("窗戶", "玻璃窗內外擦拭", 60),
        ("窗戶", "紗窗拆洗", 45),
        ("窗戶", "窗溝清潔", 30),
        ("其他", "垃圾、資源回收分類清運", 30),
        ("其他", "汽機車內部清潔", 60),
        ("其他", "電腦與手機檔案整理（選做）", 45),
        ("其他", "貼春聯、佈置（打掃完成後）", 30),
    ]
    who = ["全家", "爸爸", "爸爸", "爸爸", "小美", "全家", "爸爸", "媽媽", "小明", "小明", "媽媽", "小明", "小美", "媽媽",
           "媽媽", "爸爸", "爸爸", "媽媽", "媽媽", "媽媽", "小美", "小美", "爸爸", "爸爸", "小明", "小美", "小明", "爸爸",
           "小明", "小美", "媽媽", "爸爸", "媽媽", "全家", "爸爸", "小明", "全家", "爸爸", "小明", "全家"]
    for i, (area, item, mins) in enumerate(tasks):
        r = 5 + i
        tk.cell(row=r, column=1, value=area); tk.cell(row=r, column=2, value=item)
        tk.cell(row=r, column=3, value=who[i]); tk.cell(row=r, column=5, value=mins)
        tk.cell(row=r, column=6, value="未開始")
    # a few example progress rows
    tk["D5"] = D(2027, 1, 16); tk["F5"] = "完成"; tk["G5"] = D(2027, 1, 16)
    tk["D6"] = D(2027, 1, 23); tk["F6"] = "完成"; tk["G6"] = D(2027, 1, 23)
    tk["D7"] = D(2027, 1, 23); tk["F7"] = "進行中"
    tk["D18"] = D(2027, 1, 17); tk["F18"] = "完成"; tk["G18"] = D(2027, 1, 17)
    tk["D21"] = D(2027, 1, 24)
    for r in range(5, 5 + N):
        tk.cell(row=r, column=8, value=f'=IF(B{r}="","",IF(F{r}="完成","✔ 已完成",IF(AND(ISNUMBER(D{r}),D{r}<TODAY()),"逾期 "&(TODAY()-D{r})&" 天",IF(ISNUMBER(D{r}),"還有 "&(D{r}-TODAY())&" 天","尚未排日期"))))')
    grid(tk, 5, 4 + N, 1, 9, cjk=Z, input_cols=(1, 2, 3, 4, 5, 6, 7, 9), fmt={4: "mm/dd", 5: "0", 7: "mm/dd"})
    dv_list(tk, "=設定!$C$7:$C$18", f"A5:A{4+N}")
    dv_list(tk, "=設定!$A$7:$A$14", f"C5:C{4+N}")
    dv_list(tk, '"未開始,進行中,完成"', f"F5:F{4+N}")
    rng = f"A5:I{4+N}"
    tk.conditional_formatting.add(rng, FormulaRule(formula=['$F5="完成"'], fill=GREEN_FILL))
    tk.conditional_formatting.add(rng, FormulaRule(formula=['AND($B5<>"",$F5<>"完成",ISNUMBER($D5),$D5<TODAY())'], fill=RED_FILL))
    tk.conditional_formatting.add(rng, FormulaRule(formula=['$F5="進行中"'], fill=BLUE_FILL))
    tk.freeze_panes = "C5"
    tk.auto_filter.ref = f"A4:I{4+N}"
    widths(tk, {"A": 10, "B": 40, "C": 9, "D": 10, "E": 9, "F": 9, "G": 9, "H": 14, "I": 20})
    landscape(tk); tk.print_area = "A1:I44"

    # 斷捨離清單
    title(dc, "斷捨離清單", "先整理再打掃：決定每樣東西的去處。", cjk=Z)
    header_row(dc, 4, ["物品", "區域", "處理方式", "數量", "二手預估售價", "實際售出金額", "已處理", "備註"], cjk=Z)
    dex = [("舊外套", "臥室", "捐贈", 3, None, None, "是"), ("不用的果汁機", "廚房", "二手出售", 1, 500, 400, "是"),
           ("過期化妝品", "浴室", "丟棄", 6, None, None, "否"), ("看完的書", "客廳", "二手出售", 12, 600, None, "否"),
           ("紙箱、舊雜誌", "陽台", "回收", 1, None, None, "否")]
    for i, e in enumerate(dex):
        for j, v in enumerate(e):
            dc.cell(row=5 + i, column=j + 1, value=v)
    grid(dc, 5, 104, 1, 8, cjk=Z, input_cols=range(1, 9), fmt={5: "#,##0", 6: "#,##0"})
    dv_list(dc, "=設定!$C$7:$C$18", "B5:B104")
    dv_list(dc, '"丟棄,回收,捐贈,二手出售,保留"', "C5:C104")
    dv_list(dc, '"是,否"', "G5:G104")
    dc.conditional_formatting.add("A5:H104", FormulaRule(formula=['$G5="是"'], fill=GREEN_FILL))
    # summary at right
    header_row(dc, 4, ["處理方式", "件數", "已處理"], cjk=Z, start_col=10)
    for i, m in enumerate(["丟棄", "回收", "捐贈", "二手出售", "保留"]):
        r = 5 + i
        dc.cell(row=r, column=10, value=m)
        dc.cell(row=r, column=11, value=f'=SUMIF($C$5:$C$104,J{r},$D$5:$D$104)')
        dc.cell(row=r, column=12, value=f'=SUMIFS($D$5:$D$104,$C$5:$C$104,J{r},$G$5:$G$104,"是")')
    dc["J10"] = "合計"; dc["K10"] = "=SUM(K5:K9)"; dc["L10"] = "=SUM(L5:L9)"
    dc["J12"] = "二手預估收入"; dc["K12"] = '=SUMIF($C$5:$C$104,"二手出售",$E$5:$E$104)'
    dc["J13"] = "二手實際收入"; dc["K13"] = "=SUM($F$5:$F$104)"
    grid(dc, 5, 10, 10, 12, cjk=Z)
    grid(dc, 12, 13, 10, 11, cjk=Z, fmt={11: "#,##0"})
    for c in ("J10", "K10", "L10", "J12", "J13", "K12", "K13"):
        dc[c].font = font(Z, size=10, bold=True)
    widths(dc, {"A": 22, "B": 10, "C": 11, "D": 7, "E": 13, "F": 13, "G": 8, "H": 20, "I": 3, "J": 14, "K": 10, "L": 9})
    dc.freeze_panes = "A5"
    landscape(dc); dc.print_area = "A1:L30"

    # 清潔用品採購
    title(sp, "清潔用品採購", "列出要買的用品，打勾已買，自動加總花費。", cjk=Z)
    header_row(sp, 4, ["品項", "用途／區域", "數量", "單價", "小計", "已買", "購買地點／備註"], cjk=Z)
    sex = [("小蘇打粉", "廚房、浴室", 1, 89, "是"), ("檸檬酸", "水垢、熱水壺", 1, 79, "是"), ("超細纖維抹布（組）", "全屋", 2, 99, "否"),
           ("浴室除霉劑", "浴室", 1, 159, "否"), ("大垃圾袋", "全屋", 2, 69, "否"), ("橡膠手套", "全屋", 2, 39, "否")]
    for i, e in enumerate(sex):
        for j, v in zip((1, 2, 3, 4, 6), e):
            sp.cell(row=5 + i, column=j, value=v)
    for r in range(5, 45):
        sp.cell(row=r, column=5, value=f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}*D{r},"")')
    grid(sp, 5, 44, 1, 7, cjk=Z, input_cols=(1, 2, 3, 4, 6, 7), fmt={4: "#,##0", 5: "#,##0"})
    dv_list(sp, '"是,否"', "F5:F44")
    sp.conditional_formatting.add("A5:G44", FormulaRule(formula=['$F5="是"'], fill=GREEN_FILL))
    sp["I4"] = "預估總花費"; sp["J4"] = "=SUM(E5:E44)"
    sp["I5"] = "已花費"; sp["J5"] = '=SUMIF(F5:F44,"是",E5:E44)'
    sp["I6"] = "還要花"; sp["J6"] = "=J4-J5"
    grid(sp, 4, 6, 9, 10, cjk=Z, fmt={10: "#,##0"})
    for c in ("I4", "I5", "I6"):
        sp[c].font = font(Z, bold=True); sp[c].fill = PatternFill("solid", fgColor=LIGHT)
    widths(sp, {"A": 22, "B": 16, "C": 7, "D": 9, "E": 10, "F": 7, "G": 22, "H": 3, "I": 12, "J": 12})
    landscape(sp); sp.print_area = "A1:J25"

    # 進度總覽
    T = "大掃除清單!"
    A_, B_, C_, E_, F_ = (f"{T}${c}$5:${c}${4+N}" for c in "ABCEF")
    title(db, "大掃除進度總覽", None, cjk=Z)
    db["A2"] = '="目標完成日："&TEXT(設定!B4,"yyyy-mm-dd")&"　｜　今天："&TEXT(TODAY(),"yyyy-mm-dd")'
    db["A2"].font = font(Z, size=11, color="6B7280")
    kpi(db, 4, 1, "距離目標日", '=IF(設定!B4-TODAY()>=0,設定!B4-TODAY()&" 天","已過 "&(TODAY()-設定!B4)&" 天")', "@", cjk=Z)
    kpi(db, 4, 2, "總任務數", f'=COUNTA({B_})', cjk=Z)
    kpi(db, 4, 3, "已完成", f'=COUNTIFS({B_},"<>",{F_},"完成")', cjk=Z)
    kpi(db, 4, 4, "完成率", "=IF(B5=0,0,C5/B5)", "0%", cjk=Z)
    kpi(db, 4, 5, "逾期未完成", f'=SUMPRODUCT(({B_}<>"")*({F_}<>"完成")*ISNUMBER({T}$D$5:$D${4+N})*({T}$D$5:$D${4+N}<TODAY()))', cjk=Z)
    kpi(db, 4, 6, "剩餘預估時間", f'=ROUND(SUMIFS({E_},{B_},"<>",{F_},"<>完成")/60,1)&" 小時"', "@", cjk=Z)
    db.conditional_formatting.add("D5", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F2C14E"))
    db.conditional_formatting.add("E5", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))

    header_row(db, 8, ["區域", "任務數", "已完成", "完成率"], cjk=Z)
    for i in range(12):
        r = 9 + i
        db.cell(row=r, column=1, value=f'=IF(設定!C{7+i}="","",設定!C{7+i})')
        db.cell(row=r, column=2, value=f'=IF(A{r}="","",COUNTIFS({A_},A{r},{B_},"<>"))')
        db.cell(row=r, column=3, value=f'=IF(A{r}="","",COUNTIFS({A_},A{r},{B_},"<>",{F_},"完成"))')
        db.cell(row=r, column=4, value=f'=IF(OR(A{r}="",N(B{r})=0),"",C{r}/B{r})')
    grid(db, 9, 20, 1, 4, cjk=Z, fmt={4: "0%"})
    db.conditional_formatting.add("D9:D20", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="9AD1A8"))

    header_row(db, 8, ["負責人", "任務數", "已完成", "未完成", "預估總分鐘"], cjk=Z, start_col=6)
    for i in range(8):
        r = 9 + i
        db.cell(row=r, column=6, value=f'=IF(設定!A{7+i}="","",設定!A{7+i})')
        db.cell(row=r, column=7, value=f'=IF(F{r}="","",COUNTIFS({C_},F{r},{B_},"<>"))')
        db.cell(row=r, column=8, value=f'=IF(F{r}="","",COUNTIFS({C_},F{r},{B_},"<>",{F_},"完成"))')
        db.cell(row=r, column=9, value=f'=IF(F{r}="","",G{r}-H{r})')
        db.cell(row=r, column=10, value=f'=IF(F{r}="","",SUMIFS({E_},{C_},F{r},{B_},"<>"))')
    grid(db, 9, 16, 6, 10, cjk=Z)

    ch = BarChart(); ch.type = "bar"; ch.title = "各區域完成率"; ch.style = 10
    ch.add_data(Reference(db, min_col=4, min_row=8, max_row=18), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=1, min_row=9, max_row=18))
    ch.legend = None; ch.height = 7.5; ch.width = 13
    ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1; ch.y_axis.numFmt = "0%"
    ch.series[0].graphicalProperties.solidFill = "F2C14E"
    db.add_chart(ch, "A23")
    ch2 = BarChart(); ch2.type = "col"; ch2.grouping = "stacked"; ch2.overlap = 100; ch2.title = "每人分工"; ch2.style = 10
    ch2.add_data(Reference(db, min_col=8, max_col=9, min_row=8, max_row=13), titles_from_data=True)
    ch2.set_categories(Reference(db, min_col=6, min_row=9, max_row=13))
    ch2.height = 7.5; ch2.width = 13
    ch2.series[0].graphicalProperties.solidFill = "9AD1A8"
    ch2.series[1].graphicalProperties.solidFill = "2F3E46"
    db.add_chart(ch2, "F23")
    widths(db, {"A": 16, "B": 14, "C": 14, "D": 14, "E": 14, "F": 16, "G": 10, "H": 10, "I": 10, "J": 12})
    landscape(db); db.page_setup.fitToHeight = 1; db.print_area = "A1:K39"
    wb.move_sheet("進度總覽", offset=-4)
    wb.active = 1
    wb.save(f"{OUT}/new-year-cleaning-planner-tw.xlsx")


reading(); freelance(); cleaning()
print("ok")
