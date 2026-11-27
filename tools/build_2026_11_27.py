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



def ref(sheet, col, r1, r2):
    return f"'{sheet}'!${col}${r1}:${col}${r2}"


# =====================================================================
# 1) Student Semester Planner & Grade Calculator
# =====================================================================
def student():
    wb = Workbook()
    hw = wb.active; hw.title = "How to Use"
    db = wb.create_sheet("Dashboard")
    co = wb.create_sheet("Courses")
    asg = wb.create_sheet("Assignments")
    st = wb.create_sheet("Setup")
    NC, NA = 12, 200
    C1, C2 = 5, 4 + NC
    A1_, A2_ = 5, 4 + NA

    howto(hw, "Student Semester Planner & Grade Calculator – How to Use", [
        "## What's inside (5 sheets)",
        "How to Use – this page.",
        "Dashboard – semester countdown, projected GPA, open/overdue assignments, the next 8 deadlines, and a grade table and chart for every course.",
        "Courses – up to 12 courses: credits, target grade, class times. Calculates current grade %, letter grade, grade points, and the average score you need on the remaining work to hit your target.",
        "Assignments – every assignment, quiz, exam and project (up to 200): due date, weight, status, score. Shows days left and flags OVERDUE / DUE SOON / THIS WEEK.",
        "Setup – semester dates, the grade scale (letter grades and grade points) and the drop-down lists. Edit the scale to match your school.",
        "## Quick start",
        "1. Setup: enter your semester start and end dates. Check the grade scale – a common US 4.0 scale is filled in, but schools differ, so change the minimum % and points if yours is different.",
        "2. Courses: replace the example courses with yours. Enter credits and your target grade (e.g. 90%). Light-yellow cells are inputs; white cells are formulas – please don't type over them.",
        "3. Assignments: add each graded item from your syllabus. Pick the course from the drop-down, enter the due date and its weight as a % of the course grade (e.g. Midterm = 25%).",
        "4. When you get a grade back, enter Score and Max score (e.g. 46 / 50). The course grade updates automatically.",
        "## How the grade math works",
        "Current grade = (sum of score % × weight) ÷ (sum of weights graded so far). Example: 92% on a 10% quiz and 80% on a 25% midterm → (9.2 + 20) ÷ 35 = 83.4%.",
        "Needed on the rest = (target − points earned so far) ÷ (weight still to come). This assumes the weights for a course add up to 100% – the 'Weights entered' column helps you check. Over 100% means the target can't be reached; 0% or less means it's already secured.",
        "Projected GPA = credit-weighted average of grade points for courses that have at least one grade. It's an estimate based on your own scale, not an official transcript calculation.",
        "## Tips",
        "Statuses: Not started, In progress, Submitted, Graded. Submitted and Graded items are never flagged as overdue.",
        "Extra credit: add it as its own row with a small weight, or raise the Score above Max score.",
        "Sort or filter the Assignments sheet by Due date or Course using the filter buttons in the header row.",
        "## Works with",
        "Microsoft Excel 2010 or later (Windows/Mac), Google Sheets (File → Import) and LibreOffice Calc. No macros, no add-ins.",
        "Made by Corn_Can. For personal use; please don't resell the file.",
    ])

    # Setup
    title(st, "Setup", "Semester dates, grade scale and drop-down lists. Yellow cells are editable.")
    setup_rows(st, 4, [("Semester name", "Spring 2027", "@"),
                       ("Semester start", D(2027, 1, 19), "yyyy-mm-dd"),
                       ("Semester end", D(2027, 5, 14), "yyyy-mm-dd")])
    header_row(st, 9, ["Min %", "Letter", "Grade points"])
    scale = [(0, "F", 0.0), (60, "D-", 0.7), (63, "D", 1.0), (67, "D+", 1.3), (70, "C-", 1.7), (73, "C", 2.0),
             (77, "C+", 2.3), (80, "B-", 2.7), (83, "B", 3.0), (87, "B+", 3.3), (90, "A-", 3.7), (93, "A", 4.0)]
    for i, (m, l, p) in enumerate(scale):
        st.cell(row=10 + i, column=1, value=m); st.cell(row=10 + i, column=2, value=l); st.cell(row=10 + i, column=3, value=p)
    grid(st, 10, 21, 1, 3, input_cols=(1, 2, 3), fmt={3: "0.0"})
    st["A22"] = "Keep the scale sorted from lowest to highest Min %."
    st["A22"].font = font(size=9, italic=True, color="6B7280")
    list_col(st, 9, 5, "Assignment types", ["Homework", "Quiz", "Exam", "Project", "Paper", "Lab", "Participation", "Presentation"], 12)
    list_col(st, 9, 7, "Statuses", ["Not started", "In progress", "Submitted", "Graded"], 4, inp=False)
    widths(st, {"A": 18, "B": 16, "C": 14, "D": 3, "E": 20, "F": 3, "G": 14})
    landscape(st)
    MIN, LET, PTS = "Setup!$A$10:$A$21", "Setup!$B$10:$B$21", "Setup!$C$10:$C$21"

    # Assignments
    title(asg, "Assignments", "One row per graded item. Weight = % of the course grade. Enter Score and Max score when graded.")
    hdr = ["Course", "Assignment", "Type", "Due date", "Weight", "Status", "Score", "Max score", "Score %",
           "Days left", "Flag", "Graded weight", "Weighted pts", "Upcoming key"]
    header_row(asg, 4, hdr)
    ex = [("BIO 101", "Lab report 1", "Lab", D(2027, 1, 29), 0.05, "Graded", 46, 50),
          ("BIO 101", "Midterm exam", "Exam", D(2027, 3, 5), 0.25, "Not started", None, None),
          ("ENG 102", "Essay 1: Argument", "Paper", D(2027, 2, 5), 0.15, "Graded", 85, 100),
          ("ENG 102", "Essay 2: Research", "Paper", D(2027, 3, 19), 0.20, "In progress", None, None),
          ("MATH 151", "Quiz 1", "Quiz", D(2027, 1, 28), 0.05, "Graded", 18, 20),
          ("MATH 151", "Problem set 3", "Homework", D(2027, 2, 12), 0.05, "Submitted", None, None),
          ("PSY 100", "Chapter 1–3 quiz", "Quiz", D(2027, 2, 3), 0.10, "Graded", 27, 30),
          ("PSY 100", "Group presentation", "Presentation", D(2027, 4, 9), 0.15, "Not started", None, None)]
    for i, e in enumerate(ex):
        for c, v in enumerate(e, start=1):
            if c == 7:
                asg.cell(row=A1_ + i, column=7, value=v)
            elif c == 8:
                asg.cell(row=A1_ + i, column=8, value=v)
            else:
                asg.cell(row=A1_ + i, column=c, value=v)
    for r in range(A1_, A2_ + 1):
        asg.cell(row=r, column=9, value=f'=IF(AND(ISNUMBER(G{r}),N(H{r})>0),G{r}/H{r},"")')
        asg.cell(row=r, column=10, value=f'=IF(OR(B{r}="",NOT(ISNUMBER(D{r}))),"",D{r}-TODAY())')
        asg.cell(row=r, column=11, value=f'=IF(OR(B{r}="",NOT(ISNUMBER(D{r}))),"",IF(OR(F{r}="Submitted",F{r}="Graded"),"✓",IF(J{r}<0,"OVERDUE",IF(J{r}<=2,"DUE SOON",IF(J{r}<=7,"THIS WEEK","")))))')
        asg.cell(row=r, column=12, value=f'=IF(ISNUMBER(I{r}),N(E{r}),0)')
        asg.cell(row=r, column=13, value=f'=IF(ISNUMBER(I{r}),I{r}*N(E{r}),0)')
        asg.cell(row=r, column=14, value=f'=IF(AND(B{r}<>"",ISNUMBER(D{r}),OR(F{r}="Not started",F{r}="In progress",F{r}="")),D{r}+ROW()/100000,"")')
    grid(asg, A1_, A2_, 1, 14, input_cols=(1, 2, 3, 4, 5, 6, 7, 8),
         fmt={4: "ddd, mmm d", 5: "0%", 9: "0.0%", 12: "0%", 13: "0.00"})
    for r in range(A1_, A2_ + 1):
        asg.cell(row=r, column=11).font = font(size=9, bold=True)
        for c in (12, 13, 14):
            asg.cell(row=r, column=c).font = font(size=8, color="9CA3AF")
    dv_list(asg, f"=Courses!$A${C1}:$A${C2}", f"A{A1_}:A{A2_}")
    dv_list(asg, "=Setup!$E$10:$E$21", f"C{A1_}:C{A2_}")
    dv_list(asg, '"Not started,In progress,Submitted,Graded"', f"F{A1_}:F{A2_}")
    dvp = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                         showErrorMessage=True, error="Enter the weight as a percentage, e.g. 15%.")
    asg.add_data_validation(dvp); dvp.add(f"E{A1_}:E{A2_}")
    asg.conditional_formatting.add(f"K{A1_}:K{A2_}", FormulaRule(formula=[f'K{A1_}="OVERDUE"'], fill=RED_FILL, font=Font(color="9B1C1C", bold=True)))
    asg.conditional_formatting.add(f"K{A1_}:K{A2_}", FormulaRule(formula=[f'K{A1_}="DUE SOON"'], fill=AMBER_FILL))
    asg.conditional_formatting.add(f"K{A1_}:K{A2_}", FormulaRule(formula=[f'K{A1_}="THIS WEEK"'], fill=BLUE_FILL))
    asg.conditional_formatting.add(f"K{A1_}:K{A2_}", FormulaRule(formula=[f'K{A1_}="✓"'], fill=GREEN_FILL))
    asg.conditional_formatting.add(f"I{A1_}:I{A2_}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="9AD1A8"))
    widths(asg, {"A": 12, "B": 26, "C": 13, "D": 13, "E": 8, "F": 12, "G": 8, "H": 9, "I": 9, "J": 9, "K": 11, "L": 8, "M": 8, "N": 10})
    asg.freeze_panes = "C5"
    asg.auto_filter.ref = f"A4:N{A2_}"
    landscape(asg); asg.print_area = f"A1:K40"

    # Courses
    AC, AL, AM, AE = (ref("Assignments", c, A1_, A2_) for c in "ALME")
    AF_, AB_ = ref("Assignments", "F", A1_, A2_), ref("Assignments", "B", A1_, A2_)
    title(co, "Courses", "Yellow = your input. Grades update from the Assignments sheet.")
    hdr = ["Course code", "Course name", "Instructor", "Credits", "Target %", "Class days & time", "Room",
           "Weights entered", "Graded so far", "Current %", "Letter", "Grade pts", "Needed on rest", "Open items", "QP", "GPA cr"]
    header_row(co, 4, hdr)
    cex = [("BIO 101", "Intro to Biology", "Dr. Kim", 4, 0.90, "Mon/Wed 9:00–10:15", "SCI 210"),
           ("ENG 102", "College Writing II", "Prof. Lopez", 3, 0.88, "Tue/Thu 11:00–12:15", "HUM 104"),
           ("MATH 151", "Calculus I", "Dr. Patel", 4, 0.85, "Mon/Wed/Fri 13:00–13:50", "MTH 301"),
           ("PSY 100", "Intro to Psychology", "Prof. Reed", 3, 0.90, "Tue/Thu 14:00–15:15", "SOC 115")]
    for i, e in enumerate(cex):
        for c, v in enumerate(e, start=1):
            co.cell(row=C1 + i, column=c, value=v)
    for r in range(C1, C2 + 1):
        co.cell(row=r, column=8, value=f'=IF(A{r}="","",SUMIFS({AE},{AC},A{r}))')
        co.cell(row=r, column=9, value=f'=IF(A{r}="","",SUMIFS({AL},{AC},A{r}))')
        co.cell(row=r, column=10, value=f'=IF(OR(A{r}="",N(I{r})=0),"",SUMIFS({AM},{AC},A{r})/I{r})')
        co.cell(row=r, column=11, value=f'=IF(J{r}="","",LOOKUP(J{r}*100,{MIN},{LET}))')
        co.cell(row=r, column=12, value=f'=IF(J{r}="","",LOOKUP(J{r}*100,{MIN},{PTS}))')
        co.cell(row=r, column=13, value=f'=IF(OR(A{r}="",NOT(ISNUMBER(E{r}))),"",IF(1-N(I{r})<=0,"",(E{r}-SUMIFS({AM},{AC},A{r}))/(1-I{r})))')
        co.cell(row=r, column=14, value=f'=IF(A{r}="","",COUNTIFS({AC},A{r},{AF_},"Not started")+COUNTIFS({AC},A{r},{AF_},"In progress"))')
        co.cell(row=r, column=15, value=f'=IF(ISNUMBER(L{r}),L{r}*N(D{r}),0)')
        co.cell(row=r, column=16, value=f'=IF(ISNUMBER(L{r}),N(D{r}),0)')
    grid(co, C1, C2, 1, 16, input_cols=(1, 2, 3, 4, 5, 6, 7),
         fmt={5: "0%", 8: "0%", 9: "0%", 10: "0.0%", 12: "0.0", 13: "0.0%"})
    for r in range(C1, C2 + 1):
        co.cell(row=r, column=11).font = font(size=11, bold=True)
        for c in (15, 16):
            co.cell(row=r, column=c).font = font(size=8, color="9CA3AF")
    co.conditional_formatting.add(f"H{C1}:H{C2}", FormulaRule(formula=[f'AND(ISNUMBER(H{C1}),ROUND(H{C1},4)<>1)'], fill=AMBER_FILL))
    co.conditional_formatting.add(f"M{C1}:M{C2}", FormulaRule(formula=[f'AND(ISNUMBER(M{C1}),M{C1}>1)'], fill=RED_FILL, font=Font(color="9B1C1C", bold=True)))
    co.conditional_formatting.add(f"M{C1}:M{C2}", FormulaRule(formula=[f'AND(ISNUMBER(M{C1}),M{C1}<=0)'], fill=GREEN_FILL))
    co.conditional_formatting.add(f"J{C1}:J{C2}", FormulaRule(formula=[f'AND(ISNUMBER(J{C1}),ISNUMBER(E{C1}),J{C1}>=E{C1})'], fill=GREEN_FILL))
    co.conditional_formatting.add(f"J{C1}:J{C2}", FormulaRule(formula=[f'AND(ISNUMBER(J{C1}),ISNUMBER(E{C1}),J{C1}<E{C1})'], fill=AMBER_FILL))
    dvc = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True)
    co.add_data_validation(dvc); dvc.add(f"E{C1}:E{C2}")
    co["A18"] = "Weights entered: should total 100% per course (amber if not).   Needed on rest: red = over 100% (target out of reach), green = target already secured."
    co["A18"].font = font(size=9, italic=True, color="6B7280")
    widths(co, {"A": 11, "B": 22, "C": 13, "D": 8, "E": 9, "F": 22, "G": 10, "H": 9, "I": 9, "J": 10, "K": 8, "L": 8, "M": 10, "N": 8, "O": 6, "P": 6})
    co.freeze_panes = "C5"
    landscape(co); co.print_area = "A1:N18"

    # Dashboard
    title(db, "Semester Dashboard")
    db["A2"] = '=Setup!B4&"   ·   "&TEXT(Setup!B5,"mmm d")&" – "&TEXT(Setup!B6,"mmm d, yyyy")&"   ·   Today: "&TEXT(TODAY(),"yyyy-mm-dd")'
    db["A2"].font = font(size=11, italic=True, color="6B7280")
    AK, AN, AD = ref("Assignments", "K", A1_, A2_), ref("Assignments", "N", A1_, A2_), ref("Assignments", "D", A1_, A2_)
    kpi(db, 4, 1, "Projected GPA", f'=IF(SUM(Courses!$P${C1}:$P${C2})=0,"–",SUM(Courses!$O${C1}:$O${C2})/SUM(Courses!$P${C1}:$P${C2}))', "0.00")
    kpi(db, 4, 2, "Courses / credits", f'=COUNTIF(Courses!$A${C1}:$A${C2},"?*")&" / "&SUM(Courses!$D${C1}:$D${C2})', "@")
    kpi(db, 4, 3, "Open items", f'=COUNTIF({AF_},"Not started")+COUNTIF({AF_},"In progress")')
    kpi(db, 4, 4, "Overdue", f'=COUNTIF({AK},"OVERDUE")')
    kpi(db, 4, 5, "Due in 7 days", f'=COUNTIF({AK},"DUE SOON")+COUNTIF({AK},"THIS WEEK")')
    kpi(db, 4, 6, "Semester progress", '=IF(TODAY()<Setup!B5,0,IF(TODAY()>Setup!B6,1,(TODAY()-Setup!B5)/(Setup!B6-Setup!B5)))', "0%")
    kpi(db, 4, 7, "Days to semester end", '=IF(TODAY()<Setup!B5,"Starts in "&(Setup!B5-TODAY())&" d",MAX(0,Setup!B6-TODAY()))', "0")
    db.conditional_formatting.add("D5", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
    db.conditional_formatting.add("E5", CellIsRule(operator="greaterThan", formula=["0"], fill=AMBER_FILL))
    db.conditional_formatting.add("F5", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F2C14E"))

    db["A8"] = "Next deadlines"; db["A8"].font = font(size=13, bold=True, color=DARK)
    header_row(db, 9, ["Due", "Course", "Assignment", "Type", "Weight", "Days left"])
    for k in range(8):
        r = 10 + k
        db.cell(row=r, column=1, value=f'=IFERROR(INT(SMALL({AN},{k+1})),"")')
        m = f'MATCH(SMALL({AN},{k+1}),{AN},0)'
        db.cell(row=r, column=2, value=f'=IFERROR(INDEX({AC},{m}),"")')
        db.cell(row=r, column=3, value=f'=IFERROR(INDEX({AB_},{m}),"")')
        db.cell(row=r, column=4, value=f'=IFERROR(INDEX({ref("Assignments","C",A1_,A2_)},{m}),"")')
        db.cell(row=r, column=5, value=f'=IFERROR(INDEX({AE},{m}),"")')
        db.cell(row=r, column=6, value=f'=IF(A{r}="","",A{r}-TODAY())')
    grid(db, 10, 17, 1, 6, fmt={1: "ddd, mmm d", 5: "0%"})
    db.conditional_formatting.add("F10:F17", FormulaRule(formula=['AND(ISNUMBER(F10),F10<=2)'], fill=AMBER_FILL))

    db["A20"] = "Grades by course"; db["A20"].font = font(size=13, bold=True, color=DARK)
    header_row(db, 21, ["Course", "Credits", "Target", "Current %", "Letter", "Needed on rest", "Open items"])
    for i in range(NC):
        r = 22 + i; s = C1 + i
        db.cell(row=r, column=1, value=f'=IF(Courses!A{s}="","",Courses!A{s})')
        db.cell(row=r, column=2, value=f'=IF(A{r}="","",Courses!D{s})')
        db.cell(row=r, column=3, value=f'=IF(A{r}="","",Courses!E{s})')
        db.cell(row=r, column=4, value=f'=IF(A{r}="","",Courses!J{s})')
        db.cell(row=r, column=5, value=f'=IF(A{r}="","",Courses!K{s})')
        db.cell(row=r, column=6, value=f'=IF(A{r}="","",Courses!M{s})')
        db.cell(row=r, column=7, value=f'=IF(A{r}="","",Courses!N{s})')
    grid(db, 22, 21 + NC, 1, 7, fmt={3: "0%", 4: "0.0%", 6: "0.0%"})
    db.conditional_formatting.add(f"D22:D{21+NC}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="9AD1A8"))
    db.conditional_formatting.add(f"F22:F{21+NC}", FormulaRule(formula=['AND(ISNUMBER(F22),F22>1)'], fill=RED_FILL))
    ch = BarChart(); ch.type = "col"; ch.title = "Current grade vs target"; ch.style = 10
    ch.add_data(Reference(db, min_col=3, min_row=21, max_row=25), titles_from_data=True)
    ch.add_data(Reference(db, min_col=4, min_row=21, max_row=25), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=1, min_row=22, max_row=25))
    ch.height = 7.5; ch.width = 14
    ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1; ch.y_axis.numFmt = "0%"
    ch.series[0].graphicalProperties.solidFill = "2F3E46"
    ch.series[1].graphicalProperties.solidFill = "F2C14E"
    db.add_chart(ch, "I8")
    widths(db, {"A": 15, "B": 15, "C": 26, "D": 13, "E": 12, "F": 15, "G": 18, "H": 3})
    landscape(db); db.page_setup.fitToHeight = 1; db.print_area = "A1:Q33"
    wb.active = 1
    wb.save(f"{OUT}/student-semester-planner-grade-calculator.xlsx")


# =====================================================================
# 2) Social Media Content Calendar
# =====================================================================
def content():
    wb = Workbook()
    hw = wb.active; hw.title = "How to Use"
    db = wb.create_sheet("Dashboard")
    cal = wb.create_sheet("Monthly Calendar")
    cp = wb.create_sheet("Content Plan")
    ib = wb.create_sheet("Idea Bank")
    st = wb.create_sheet("Setup")
    N = 300
    R1, R2 = 5, 4 + N

    howto(hw, "Social Media Content Calendar – How to Use", [
        "## What's inside (6 sheets)",
        "How to Use – this page.",
        "Dashboard – posts planned and published this month, posts scheduled for the next 7 days, late posts, average engagement rate, plus tables and a chart by platform, content pillar and status.",
        "Monthly Calendar – pick a year and month and see every day as a calendar box with the first post of the day and how many more are planned.",
        "Content Plan – your master list of posts (up to 300): date, time, platform, content pillar, format, hook/title, caption notes, hashtags/CTA, status, link and results (reach, likes, comments, shares, saves). Engagement rate is calculated for you.",
        "Idea Bank – park ideas before they get a date. Mark one 'Moved to plan' once it's scheduled.",
        "Setup – your platforms, content pillars, formats and statuses. The drop-downs everywhere read from these lists.",
        "## Quick start",
        "1. Setup: replace the example platforms and content pillars with your own (up to 10 of each).",
        "2. Content Plan: add one row per post. Light-yellow cells are inputs; white cells are formulas – please don't type over them.",
        "3. Move posts through the statuses: Idea → Drafting → Scheduled → Published.",
        "4. After a post has been live for a few days, copy the numbers from the platform's own insights into Reach, Likes, Comments, Shares and Saves.",
        "5. Monthly Calendar: choose the year and month in the yellow cells at the top.",
        "## Flags & formulas",
        "LATE = the publish date has passed and the post isn't marked Published.   PREP = it goes out within 3 days but is still an Idea or Drafting.",
        "Engagement rate = (likes + comments + shares + saves) ÷ reach. Platforms define reach and engagement differently, so compare a platform with itself over time rather than across platforms.",
        "Week = ISO week number, handy for weekly planning sessions.",
        "## Works with",
        "Microsoft Excel 2013 or later (Windows/Mac), Google Sheets (File → Import) and LibreOffice Calc. No macros, no add-ins, no account connections – you type or paste the numbers yourself.",
        "Made by Corn_Can. For personal and business use; please don't resell the file. Platform names are used only as labels and belong to their owners.",
    ])

    title(st, "Setup", "Edit the yellow lists. Drop-downs across the workbook use them.")
    list_col(st, 4, 1, "Platforms", ["Instagram", "TikTok", "YouTube", "Pinterest", "LinkedIn", "Facebook", "Blog", "Newsletter"], 10)
    list_col(st, 4, 3, "Content pillars", ["Education / tips", "Behind the scenes", "Product / offer", "Community", "Inspiration", "Personal story"], 10)
    list_col(st, 4, 5, "Formats", ["Reel / short video", "Carousel", "Single image", "Story", "Long video", "Text post", "Live", "Article"], 10)
    list_col(st, 4, 7, "Statuses", ["Idea", "Drafting", "Scheduled", "Published"], 4, inp=False)
    widths(st, {"A": 18, "B": 3, "C": 22, "D": 3, "E": 22, "F": 3, "G": 14})
    landscape(st)

    # Content plan
    title(cp, "Content Plan", "One row per post. Yellow = your input; Engagement, Week and Flag are calculated.")
    hdr = ["Publish date", "Time", "Platform", "Pillar", "Format", "Hook / title", "Caption notes", "Hashtags / CTA",
           "Status", "Link", "Reach", "Likes", "Comments", "Shares", "Saves", "Eng. rate", "Week", "Flag"]
    header_row(cp, 4, hdr)
    ex = [(D(2026, 12, 1), "09:00", "Instagram", "Education / tips", "Carousel", "5 ways to plan a week of posts in 30 min", "Slide per tip, end with checklist", "Save this for Sunday", "Published", "", 1840, 112, 9, 14, 61),
          (D(2026, 12, 2), "18:30", "TikTok", "Behind the scenes", "Reel / short video", "Packing orders: a day in the shop", "Time-lapse + voiceover", "Follow for more BTS", "Published", "", 5230, 388, 22, 31, 40),
          (D(2026, 12, 4), "12:00", "Pinterest", "Product / offer", "Single image", "Holiday gift guide under $25", "Vertical pin, link to blog", "#giftguide", "Published", "", 960, 18, 0, 3, 57),
          (D(2026, 12, 8), "08:00", "LinkedIn", "Personal story", "Text post", "What I learned from my first 100 customers", "3 lessons, short paragraphs", "Ask: what was your first win?", "Scheduled", "", None, None, None, None, None),
          (D(2026, 12, 9), "09:00", "Instagram", "Community", "Story", "Poll: which product should come back?", "2 options + sticker", "", "Drafting", "", None, None, None, None, None),
          (D(2026, 12, 9), "17:00", "YouTube", "Education / tips", "Long video", "Year-end review: my content numbers", "Outline in Idea Bank", "Subscribe", "Idea", "", None, None, None, None, None)]
    for i, e in enumerate(ex):
        for c, v in enumerate(e, start=1):
            cp.cell(row=R1 + i, column=c, value=v)
    for r in range(R1, R2 + 1):
        cp.cell(row=r, column=16, value=f'=IF(N(K{r})>0,(N(L{r})+N(M{r})+N(N{r})+N(O{r}))/K{r},"")')
        cp.cell(row=r, column=17, value=f'=IF(ISNUMBER(A{r}),"W"&TEXT(_xlfn.ISOWEEKNUM(A{r}),"00"),"")')
        cp.cell(row=r, column=18, value=f'=IF(OR(F{r}="",NOT(ISNUMBER(A{r}))),"",IF(I{r}="Published","✓",IF(A{r}<TODAY(),"LATE",IF(AND(A{r}-TODAY()<=3,OR(I{r}="Idea",I{r}="Drafting",I{r}="")),"PREP",""))))')
    grid(cp, R1, R2, 1, 18, input_cols=tuple(range(1, 16)),
         fmt={1: "ddd, mmm d", 11: "#,##0", 12: "#,##0", 13: "#,##0", 14: "#,##0", 15: "#,##0", 16: "0.0%"})
    for r in range(R1, R2 + 1):
        cp.cell(row=r, column=18).font = font(size=9, bold=True)
        for c in (6, 7, 8):
            cp.cell(row=r, column=c).alignment = Alignment(vertical="center", wrap_text=False)
    dv_list(cp, "=Setup!$A$5:$A$14", f"C{R1}:C{R2}")
    dv_list(cp, "=Setup!$C$5:$C$14", f"D{R1}:D{R2}")
    dv_list(cp, "=Setup!$E$5:$E$14", f"E{R1}:E{R2}")
    dv_list(cp, '"Idea,Drafting,Scheduled,Published"', f"I{R1}:I{R2}")
    cp.conditional_formatting.add(f"R{R1}:R{R2}", FormulaRule(formula=[f'R{R1}="LATE"'], fill=RED_FILL, font=Font(color="9B1C1C", bold=True)))
    cp.conditional_formatting.add(f"R{R1}:R{R2}", FormulaRule(formula=[f'R{R1}="PREP"'], fill=AMBER_FILL))
    cp.conditional_formatting.add(f"R{R1}:R{R2}", FormulaRule(formula=[f'R{R1}="✓"'], fill=GREEN_FILL))
    cp.conditional_formatting.add(f"I{R1}:I{R2}", FormulaRule(formula=[f'I{R1}="Scheduled"'], fill=BLUE_FILL))
    cp.conditional_formatting.add(f"I{R1}:I{R2}", FormulaRule(formula=[f'I{R1}="Published"'], fill=GREEN_FILL))
    widths(cp, {"A": 12, "B": 7, "C": 11, "D": 17, "E": 17, "F": 36, "G": 26, "H": 20, "I": 11, "J": 12,
                "K": 8, "L": 7, "M": 9, "N": 7, "O": 7, "P": 9, "Q": 6, "R": 8})
    cp.freeze_panes = "B5"
    cp.auto_filter.ref = f"A4:R{R2}"
    landscape(cp); cp.print_area = "A1:R40"

    # Idea bank
    title(ib, "Idea Bank", "Capture ideas now, schedule them later.")
    header_row(ib, 4, ["Date added", "Idea", "Platform", "Pillar", "Format", "Priority", "Source / notes", "Moved to plan?"])
    iex = [(D(2026, 11, 20), "Year-end review video: my content numbers", "YouTube", "Education / tips", "Long video", "High", "Pull stats from Dashboard", "Yes"),
           (D(2026, 11, 22), "Myth vs fact carousel for beginners", "Instagram", "Education / tips", "Carousel", "Medium", "Questions from comments", ""),
           (D(2026, 11, 25), "Customer spotlight interview", "LinkedIn", "Community", "Text post", "Low", "Ask 2 customers first", "")]
    for i, e in enumerate(iex):
        for c, v in enumerate(e, start=1):
            ib.cell(row=5 + i, column=c, value=v)
    grid(ib, 5, 104, 1, 8, input_cols=tuple(range(1, 9)), fmt={1: "yyyy-mm-dd"})
    dv_list(ib, "=Setup!$A$5:$A$14", "C5:C104")
    dv_list(ib, "=Setup!$C$5:$C$14", "D5:D104")
    dv_list(ib, "=Setup!$E$5:$E$14", "E5:E104")
    dv_list(ib, '"High,Medium,Low"', "F5:F104")
    dv_list(ib, '"Yes"', "H5:H104")
    ib.conditional_formatting.add("A5:H104", FormulaRule(formula=['$H5="Yes"'], fill=GREY_FILL, font=Font(color="9CA3AF")))
    ib.conditional_formatting.add("F5:F104", FormulaRule(formula=['F5="High"'], fill=AMBER_FILL))
    widths(ib, {"A": 12, "B": 42, "C": 12, "D": 18, "E": 18, "F": 9, "G": 28, "H": 14})
    ib.freeze_panes = "A5"
    landscape(ib); ib.print_area = "A1:H30"

    # Monthly calendar
    PD, PC, PF, PS = (ref("Content Plan", c, R1, R2) for c in "ACFI")
    cal["A1"] = "Monthly Content Calendar"
    cal["A1"].font = font(size=18, bold=True, color=DARK)
    cal["A2"] = "Year"; cal["B2"] = 2026; cal["C2"] = "Month"; cal["D2"] = 12
    for c in ("A2", "C2"):
        cal[c].font = font(bold=True); cal[c].alignment = Alignment(horizontal="right")
    for c in ("B2", "D2"):
        cal[c].fill = INPUT_FILL; cal[c].border = BORDER; cal[c].font = font(size=12, bold=True)
        cal[c].alignment = Alignment(horizontal="center")
    dv_list(cal, '"2026,2027,2028,2029,2030"', "B2")
    dv_list(cal, '"1,2,3,4,5,6,7,8,9,10,11,12"', "D2")
    cal["E2"] = '=TEXT(DATE(B2,D2,1),"mmmm yyyy")&"  ·  "&COUNTIFS(' + PD + ',">="&DATE(B2,D2,1),' + PD + ',"<"&DATE(B2,D2+1,1))&" posts planned"'
    cal["E2"].font = font(size=13, bold=True, color=DARK)
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    header_row(cal, 4, days)
    START = "(DATE($B$2,$D$2,1)-WEEKDAY(DATE($B$2,$D$2,1),2)+1)"
    for w in range(6):
        r0 = 5 + w * 3
        for d in range(7):
            col = CL(1 + d)
            dc = cal.cell(row=r0, column=1 + d, value=f"={START}+{w*7+d}")
            dc.number_format = "d"
            dc.font = font(size=11, bold=True, color=DARK)
            dc.alignment = Alignment(horizontal="right", vertical="top")
            dref = f"{col}{r0}"
            t = cal.cell(row=r0 + 1, column=1 + d, value=f'=IFERROR(INDEX({PC},MATCH({dref},{PD},0))&": "&INDEX({PF},MATCH({dref},{PD},0)),"")')
            t.font = font(size=8); t.alignment = Alignment(wrap_text=True, vertical="top")
            m = cal.cell(row=r0 + 2, column=1 + d, value=f'=IF(COUNTIF({PD},{dref})>1,"+"&(COUNTIF({PD},{dref})-1)&" more","")')
            m.font = font(size=8, bold=True, color="B7791F")
            for rr in range(r0, r0 + 3):
                cell = cal.cell(row=rr, column=1 + d)
                cell.border = Border(left=thin, right=thin, top=thin if rr == r0 else None, bottom=thin if rr == r0 + 2 else None)
        cal.row_dimensions[r0].height = 18
        cal.row_dimensions[r0 + 1].height = 46
        cal.row_dimensions[r0 + 2].height = 14
    rng = "A5:G22"
    # anchor per cell: date row for the block = row - MOD(row-5,3)
    anchor = 'INDEX($A$5:$G$22,ROW()-4-MOD(ROW()-5,3),COLUMN())'
    cal.conditional_formatting.add(rng, FormulaRule(formula=[f'MONTH({anchor})<>$D$2'], fill=GREY_FILL, font=Font(color="C0C4CC"), stopIfTrue=True))
    cal.conditional_formatting.add(rng, FormulaRule(formula=[f'{anchor}=TODAY()'], fill=PatternFill("solid", fgColor="FBD5D5")))
    cal.conditional_formatting.add(rng, FormulaRule(formula=[f'COUNTIF({PD},{anchor})>0'], fill=PatternFill("solid", fgColor=LIGHT)))
    for i in range(7):
        cal.column_dimensions[CL(1 + i)].width = 19
    cal["A24"] = "Pick the year and month in the yellow cells. Each box shows the first post of the day (Platform: hook) and how many more are planned. Today is highlighted in pink."
    cal["A24"].font = font(size=9, italic=True, color="6B7280")
    landscape(cal); cal.page_setup.fitToHeight = 1; cal.print_area = "A1:G24"

    # Dashboard
    title(db, "Content Dashboard")
    db["A2"] = '="This month: "&TEXT(TODAY(),"mmmm yyyy")&"   ·   Today: "&TEXT(TODAY(),"yyyy-mm-dd")'
    db["A2"].font = font(size=11, italic=True, color="6B7280")
    PR, PP_, PK, PDp = ref("Content Plan", "R", R1, R2), ref("Content Plan", "P", R1, R2), ref("Content Plan", "K", R1, R2), ref("Content Plan", "D", R1, R2)
    m0, m1 = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)", "DATE(YEAR(TODAY()),MONTH(TODAY())+1,1)"
    kpi(db, 4, 1, "Planned this month", f'=COUNTIFS({PD},">="&{m0},{PD},"<"&{m1})')
    kpi(db, 4, 2, "Published this month", f'=COUNTIFS({PD},">="&{m0},{PD},"<"&{m1},{PS},"Published")')
    kpi(db, 4, 3, "Next 7 days", f'=COUNTIFS({PD},">="&TODAY(),{PD},"<="&TODAY()+7)')
    kpi(db, 4, 4, "Late", f'=COUNTIF({PR},"LATE")')
    kpi(db, 4, 5, "Needs prep (≤3 days)", f'=COUNTIF({PR},"PREP")')
    kpi(db, 4, 6, "Avg engagement", f'=IFERROR(AVERAGE({PP_}),"–")', "0.0%")
    db.conditional_formatting.add("D5", CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL))
    db.conditional_formatting.add("E5", CellIsRule(operator="greaterThan", formula=["0"], fill=AMBER_FILL))

    header_row(db, 8, ["Platform", "Total posts", "Published", "Scheduled", "Total reach", "Avg eng. rate"])
    for i in range(10):
        r = 9 + i
        db.cell(row=r, column=1, value=f'=IF(Setup!A{5+i}="","",Setup!A{5+i})')
        db.cell(row=r, column=2, value=f'=IF(A{r}="","",COUNTIFS({PC},A{r},{PF},"?*"))')
        db.cell(row=r, column=3, value=f'=IF(A{r}="","",COUNTIFS({PC},A{r},{PS},"Published"))')
        db.cell(row=r, column=4, value=f'=IF(A{r}="","",COUNTIFS({PC},A{r},{PS},"Scheduled"))')
        db.cell(row=r, column=5, value=f'=IF(A{r}="","",SUMIFS({PK},{PC},A{r}))')
        db.cell(row=r, column=6, value=f'=IF(A{r}="","",IFERROR(AVERAGEIFS({PP_},{PC},A{r}),""))')
    grid(db, 9, 18, 1, 6, fmt={5: "#,##0", 6: "0.0%"})
    header_row(db, 8, ["Content pillar", "Posts", "Share of mix"], start_col=8)
    for i in range(10):
        r = 9 + i
        db.cell(row=r, column=8, value=f'=IF(Setup!C{5+i}="","",Setup!C{5+i})')
        db.cell(row=r, column=9, value=f'=IF(H{r}="","",COUNTIFS({PDp},H{r},{PF},"?*"))')
        db.cell(row=r, column=10, value=f'=IF(OR(H{r}="",SUM($I$9:$I$18)=0),"",I{r}/SUM($I$9:$I$18))')
    grid(db, 9, 18, 8, 10, fmt={10: "0%"})
    db.conditional_formatting.add("J9:J18", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F2C14E"))
    header_row(db, 20, ["Status", "Posts"], start_col=8)
    for i, s in enumerate(["Idea", "Drafting", "Scheduled", "Published"]):
        db.cell(row=21 + i, column=8, value=s)
        db.cell(row=21 + i, column=9, value=f'=COUNTIFS({PS},H{21+i},{PF},"?*")')
    grid(db, 21, 24, 8, 9)
    ch = BarChart(); ch.type = "col"; ch.title = "Posts by platform"; ch.style = 10; ch.grouping = "stacked"; ch.overlap = 100
    ch.add_data(Reference(db, min_col=3, min_row=8, max_row=16), titles_from_data=True)
    ch.add_data(Reference(db, min_col=4, min_row=8, max_row=16), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=1, min_row=9, max_row=16))
    ch.height = 7.5; ch.width = 15
    ch.series[0].graphicalProperties.solidFill = "4CAF7A"
    ch.series[1].graphicalProperties.solidFill = "F2C14E"
    db.add_chart(ch, "A20")
    widths(db, {"A": 18, "B": 16, "C": 18, "D": 12, "E": 18, "F": 16, "G": 3, "H": 20, "I": 9, "J": 13})
    landscape(db); db.page_setup.fitToHeight = 1; db.print_area = "A1:J36"
    wb.active = 1
    wb.save(f"{OUT}/social-media-content-calendar.xlsx")


# =====================================================================
# 3) 尾牙籌備＆抽獎表
# =====================================================================
def weiya():
    Z = True
    wb = Workbook()
    hw = wb.active; hw.title = "使用說明"
    ov = wb.create_sheet("總覽")
    se = wb.create_sheet("活動設定")
    bu = wb.create_sheet("預算表")
    pr = wb.create_sheet("獎品清單")
    pe = wb.create_sheet("人員名單")
    dr = wb.create_sheet("抽獎")
    tk = wb.create_sheet("籌備清單")
    NP, NZ, NB, NT, ND = 300, 30, 40, 50, 150
    P1, P2 = 5, 4 + NP

    howto(hw, "尾牙籌備＆抽獎表 – 使用說明", [
        "## 內容（共 8 個工作表）",
        "使用說明 – 本頁。",
        "總覽 – 尾牙倒數天數、出席／未回覆／素食人數、建議桌數、預算與實際支出、獎品總額、籌備進度，以及各部門出席統計與圖表。",
        "活動設定 – 活動名稱、日期、地點、總預算、每桌人數、抽獎資格（是否須完成報到），以及部門與預算類別清單。",
        "預算表 – 每個支出項目的單價、數量、預估金額、實際金額與差額，付款狀態一目了然（最多 40 項）。",
        "獎品清單 – 依抽出順序列出獎項、獎品、名額與金額，自動算出每個獎項對應的抽獎序號（最多 30 個獎項）。",
        "人員名單 – 員工、部門、出席回覆、葷素、桌號與現場報到（最多 300 人），自動判斷誰有抽獎資格。",
        "抽獎 – 用 Excel 亂數一次抽出所有名額，每人最多中一次；另有「正式得獎人」欄位讓你鎖定結果。",
        "籌備清單 – 訂場地、發邀請、採購獎品等工作，追蹤負責人、期限與進度，逾期自動標紅。",
        "## 快速開始",
        "1. 活動設定：填入尾牙日期、總預算、每桌人數，並把部門清單改成貴公司的部門。",
        "2. 人員名單：貼上員工名單，收集出席與葷素回覆。淺黃色是輸入欄，白色是公式欄，請不要覆蓋。",
        "3. 獎品清單：由小獎到大獎依序輸入（抽獎時由第 1 列開始抽），填好名額與單價。",
        "4. 預算表、籌備清單：列出所有支出與工作，活動前每週更新一次。",
        "## 抽獎怎麼用",
        "• 有抽獎資格的人（出席，且依設定已報到）會自動取得一個隨機數字並排出順位，順位 1 的人得到第 1 個名額，依此類推，所以同一人不會重複中獎。",
        "• 亂數會在每次重新計算時改變：Windows Excel 按 F9、Mac 按 Fn+F9 或 Command+=，就會重新抽一次。也就是說，現場只要按一次就能抽出全部名單。",
        "• 抽完請馬上鎖定結果：選取「抽獎」工作表的 D 欄中獎人 → 複製 → 在 F 欄「正式得獎人」用「選擇性貼上 → 值」貼上。之後再怎麼重算，F 欄都不會變。",
        "• 想要一個獎項一個獎項慢慢抽？先抽完全部、鎖定 F 欄，再依序公布即可；或用投影只顯示該獎項那幾列。",
        "• 名額多於有資格人數時，中獎人會顯示「（人數不足）」。",
        "## 小提醒",
        "• 獎金或獎品價值較高時，可能涉及所得申報與扣繳，請依國稅局規定或與公司會計確認，本範本不計算稅額。",
        "• 範例資料（人名、金額）皆為虛構，請直接覆蓋或刪除。",
        "## 適用軟體",
        "Microsoft Excel 2010 以上（Windows／Mac）、Google 試算表（檔案 → 匯入）、LibreOffice Calc。不含巨集。",
        "Corn_Can 製作。限個人與公司內部使用，請勿轉售本檔案。",
    ], cjk=Z)

    # 活動設定
    title(se, "活動設定", "黃色儲存格可修改。", cjk=Z)
    setup_rows(se, 4, [("活動名稱", "2026 年度尾牙暨春酒", "@"),
                       ("活動日期", D(2027, 1, 22), "yyyy-mm-dd"),
                       ("開始時間", "18:30", "@"),
                       ("地點", "（餐廳名稱／地址）", "@"),
                       ("總預算（元）", 300000, "#,##0"),
                       ("每桌人數", 10, "0"),
                       ("抽獎須完成報到", "是", "@")], cjk=Z)
    dv_list(se, '"是,否"', "B10")
    list_col(se, 13, 1, "部門", ["業務部", "行銷部", "研發部", "財務部", "人資部", "客服部", "管理部"], 15, cjk=Z)
    list_col(se, 13, 3, "預算類別", ["餐飲", "場地布置", "主持表演", "音響燈光", "獎品", "交通", "攝影", "其他"], 12, cjk=Z)
    widths(se, {"A": 20, "B": 26, "C": 16})
    landscape(se)

    # 人員名單
    title(pe, "人員名單", "出席回覆、葷素、報到；最右側兩欄為抽獎用公式，請勿修改。", cjk=Z)
    header_row(pe, 4, ["員工編號", "姓名", "部門", "出席回覆", "葷素", "桌號", "現場報到", "備註", "抽獎資格", "隨機數", "抽獎順位"], cjk=Z)
    pex = [("A001", "王小明", "業務部", "出席", "葷", 1, "是", ""),
           ("A002", "陳怡君", "行銷部", "出席", "素", 1, "是", ""),
           ("A003", "林志豪", "研發部", "出席", "葷", 2, "是", ""),
           ("A004", "黃雅婷", "財務部", "不出席", "", None, "", "出差"),
           ("A005", "張家瑋", "研發部", "出席", "葷", 2, "是", ""),
           ("A006", "李佳穎", "人資部", "未回覆", "", None, "", "")]
    for i, e in enumerate(pex):
        for c, v in enumerate(e, start=1):
            pe.cell(row=P1 + i, column=c, value=v)
    for r in range(P1, P2 + 1):
        pe.cell(row=r, column=9, value=f'=IF(AND(B{r}<>"",D{r}="出席",OR(活動設定!$B$10="否",G{r}="是")),"是","")')
        pe.cell(row=r, column=10, value=f'=IF(I{r}="是",RAND(),"")')
        pe.cell(row=r, column=11, value=f'=IF(I{r}="是",RANK(J{r},$J${P1}:$J${P2}),"")')
    grid(pe, P1, P2, 1, 11, cjk=Z, input_cols=(1, 2, 3, 4, 5, 6, 7, 8), fmt={10: "0.0000"})
    for r in range(P1, P2 + 1):
        pe.cell(row=r, column=10).font = font(Z, size=8, color="9CA3AF")
    dv_list(pe, "=活動設定!$A$14:$A$28", f"C{P1}:C{P2}")
    dv_list(pe, '"出席,不出席,未回覆"', f"D{P1}:D{P2}")
    dv_list(pe, '"葷,素,蛋奶素"', f"E{P1}:E{P2}")
    dv_list(pe, '"是"', f"G{P1}:G{P2}")
    pe.conditional_formatting.add(f"D{P1}:D{P2}", FormulaRule(formula=[f'D{P1}="未回覆"'], fill=AMBER_FILL))
    pe.conditional_formatting.add(f"D{P1}:D{P2}", FormulaRule(formula=[f'D{P1}="不出席"'], fill=GREY_FILL))
    pe.conditional_formatting.add(f"I{P1}:I{P2}", FormulaRule(formula=[f'I{P1}="是"'], fill=GREEN_FILL))
    widths(pe, {"A": 10, "B": 12, "C": 11, "D": 10, "E": 8, "F": 7, "G": 9, "H": 18, "I": 9, "J": 9, "K": 9})
    pe.freeze_panes = "C5"
    pe.auto_filter.ref = f"A4:K{P2}"
    landscape(pe); pe.print_area = "A1:K40"

    # 獎品清單
    title(pr, "獎品清單", "由第 1 列開始抽：建議由小獎排到大獎。起訖序號自動計算。", cjk=Z)
    header_row(pr, 4, ["抽出順序", "獎項", "獎品", "名額", "單價", "小計", "提供者", "起始序號", "結束序號"], cjk=Z)
    zex = [("三獎", "氣炸鍋", 2, 3000, "公司"),
           ("二獎", "現金 10,000 元", 1, 10000, "總經理加碼"),
           ("頭獎", "現金 30,000 元", 1, 30000, "董事長加碼")]
    for i, e in enumerate(zex):
        for c, v in enumerate(e, start=2):
            pr.cell(row=5 + i, column=c, value=v)
    for i in range(NZ):
        r = 5 + i
        pr.cell(row=r, column=1, value=i + 1)
        pr.cell(row=r, column=6, value=f'=IF(B{r}="","",N(D{r})*N(E{r}))')
        pr.cell(row=r, column=8, value=f'=IF(OR(B{r}="",N(D{r})=0),"",SUM($D$4:D{r-1})+1)' if r > 5 else f'=IF(OR(B{r}="",N(D{r})=0),"",1)')
        pr.cell(row=r, column=9, value=f'=IF(H{r}="","",H{r}+D{r}-1)')
    grid(pr, 5, 4 + NZ, 1, 9, cjk=Z, input_cols=(2, 3, 4, 5, 7), fmt={5: "#,##0", 6: "#,##0"})
    pr["K4"] = "獎品總名額"; pr["L4"] = f"=SUM(D5:D{4+NZ})"
    pr["K5"] = "獎品總金額"; pr["L5"] = f"=SUM(F5:F{4+NZ})"
    grid(pr, 4, 5, 11, 12, cjk=Z, fmt={12: "#,##0"})
    for c in ("K4", "K5"):
        pr[c].font = font(Z, bold=True); pr[c].fill = PatternFill("solid", fgColor=LIGHT)
    widths(pr, {"A": 9, "B": 12, "C": 24, "D": 7, "E": 10, "F": 11, "G": 13, "H": 9, "I": 9, "J": 3, "K": 12, "L": 11})
    landscape(pr); pr.print_area = "A1:L30"

    # 抽獎
    PB, PC_, PK = ref("人員名單", "B", P1, P2), ref("人員名單", "C", P1, P2), ref("人員名單", "K", P1, P2)
    ZB, ZC, ZH = "獎品清單!$B$5:$B$34", "獎品清單!$C$5:$C$34", "獎品清單!$H$5:$H$34"
    title(dr, "抽獎", "按 F9 重新抽獎；確定後把 D 欄複製，以「貼上值」貼到 F 欄鎖定。", cjk=Z)
    dr["H4"] = "可抽獎人數"; dr["I4"] = f'=COUNTIF({ref("人員名單","I",P1,P2)},"是")'
    dr["H5"] = "獎品總名額"; dr["I5"] = "=獎品清單!L4"
    dr["H6"] = "狀態"; dr["I6"] = '=IF(I5=0,"尚未輸入獎品",IF(I4<I5,"名額多於人數","OK"))'
    grid(dr, 4, 6, 8, 9, cjk=Z)
    for c in ("H4", "H5", "H6"):
        dr[c].font = font(Z, bold=True); dr[c].fill = PatternFill("solid", fgColor=LIGHT)
    header_row(dr, 4, ["序號", "獎項", "獎品", "中獎人（亂數）", "部門", "正式得獎人（貼上值）"], cjk=Z)
    for k in range(ND):
        r = 5 + k
        dr.cell(row=r, column=1, value=k + 1)
        dr.cell(row=r, column=2, value=f'=IF(A{r}>獎品清單!$L$4,"",INDEX({ZB},MATCH(A{r},{ZH},1)))')
        dr.cell(row=r, column=3, value=f'=IF(B{r}="","",INDEX({ZC},MATCH(A{r},{ZH},1)))')
        dr.cell(row=r, column=4, value=f'=IF(B{r}="","",IFERROR(INDEX({PB},MATCH(A{r},{PK},0)),"（人數不足）"))')
        dr.cell(row=r, column=5, value=f'=IF(OR(B{r}="",D{r}="（人數不足）"),"",INDEX({PC_},MATCH(A{r},{PK},0)))')
    grid(dr, 5, 4 + ND, 1, 6, cjk=Z, input_cols=(6,))
    for r in range(5, 5 + ND):
        dr.cell(row=r, column=4).font = font(Z, size=11, bold=True)
    dr.conditional_formatting.add(f"A5:F{4+ND}", FormulaRule(formula=['AND($B5<>"",MOD(MATCH($B5,' + ZB + ',0),2)=0)'], fill=PatternFill("solid", fgColor=LIGHT)))
    dr.conditional_formatting.add("I6", FormulaRule(formula=['I6<>"OK"'], fill=RED_FILL))
    dr.conditional_formatting.add("I6", FormulaRule(formula=['I6="OK"'], fill=GREEN_FILL))
    widths(dr, {"A": 7, "B": 12, "C": 24, "D": 16, "E": 11, "F": 24, "G": 3, "H": 12, "I": 14})
    dr.freeze_panes = "A5"
    landscape(dr); dr.print_area = "A1:F40"

    # 預算表
    title(bu, "預算表", "預估金額 = 單價 × 數量；填入實際金額後自動算差額。", cjk=Z)
    header_row(bu, 4, ["類別", "項目", "單價", "數量", "預估金額", "實際金額", "差額（實際－預估）", "付款狀態", "廠商／備註"], cjk=Z)
    bex = [("餐飲", "桌菜（含服務費）", 12000, 15, None, "已付訂金", "訂金 3 萬"),
           ("場地布置", "背板與氣球布置", 18000, 1, 16500, "已付清", ""),
           ("主持表演", "專業主持人", 15000, 1, None, "未付款", ""),
           ("音響燈光", "音響、投影設備", 12000, 1, None, "未付款", "場地加購"),
           ("獎品", "抽獎獎品（連動獎品清單）", "=獎品清單!L5", 1, None, "未付款", "自動帶入獎品總額"),
           ("交通", "接駁車", 6000, 2, None, "未付款", "")]
    for i, e in enumerate(bex):
        r = 5 + i
        for c, v in zip((1, 2, 3, 4, 6, 8, 9), e):
            bu.cell(row=r, column=c, value=v)
    for r in range(5, 5 + NB):
        bu.cell(row=r, column=5, value=f'=IF(B{r}="","",N(C{r})*N(D{r}))')
        bu.cell(row=r, column=7, value=f'=IF(ISNUMBER(F{r}),F{r}-N(E{r}),"")')
    grid(bu, 5, 4 + NB, 1, 9, cjk=Z, input_cols=(1, 2, 3, 4, 6, 8, 9), fmt={3: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0;[Red]-#,##0"})
    dv_list(bu, "=活動設定!$C$14:$C$25", f"A5:A{4+NB}")
    dv_list(bu, '"未付款,已付訂金,已付清"', f"H5:H{4+NB}")
    bu.conditional_formatting.add(f"G5:G{4+NB}", FormulaRule(formula=['AND(ISNUMBER(G5),G5>0)'], fill=RED_FILL))
    bu.conditional_formatting.add(f"H5:H{4+NB}", FormulaRule(formula=['H5="已付清"'], fill=GREEN_FILL))
    r = 5 + NB
    bu.cell(row=r, column=2, value="合計")
    bu.cell(row=r, column=5, value=f"=SUM(E5:E{4+NB})")
    bu.cell(row=r, column=6, value=f"=SUM(F5:F{4+NB})")
    grid(bu, r, r, 1, 9, cjk=Z, fmt={5: "#,##0", 6: "#,##0"})
    for c in range(1, 10):
        bu.cell(row=r, column=c).font = font(Z, bold=True); bu.cell(row=r, column=c).fill = PatternFill("solid", fgColor=LIGHT)
    widths(bu, {"A": 11, "B": 26, "C": 10, "D": 7, "E": 11, "F": 11, "G": 14, "H": 11, "I": 20})
    bu.freeze_panes = "A5"
    landscape(bu); bu.print_area = "A1:I25"

    # 籌備清單
    title(tk, "籌備清單", "期限已過且未完成會標紅；7 天內到期標黃。", cjk=Z)
    header_row(tk, 4, ["工作項目", "負責人", "期限", "狀態", "剩餘天數", "提醒", "備註"], cjk=Z)
    tex = [("確認日期與預算", "福委會", D(2026, 11, 20), "已完成", ""),
           ("訂場地與桌菜", "人資部", D(2026, 11, 30), "已完成", "已付訂金"),
           ("發出席調查（含葷素）", "人資部", D(2026, 12, 15), "進行中", "表單連結"),
           ("確認獎品與採購", "管理部", D(2027, 1, 8), "未開始", ""),
           ("主持流程與節目表", "行銷部", D(2027, 1, 12), "未開始", ""),
           ("現場報到與抽獎彩排", "人資部", D(2027, 1, 21), "未開始", "")]
    for i, e in enumerate(tex):
        for c, v in zip((1, 2, 3, 4, 7), e):
            tk.cell(row=5 + i, column=c, value=v)
    for r in range(5, 5 + NT):
        tk.cell(row=r, column=5, value=f'=IF(OR(A{r}="",NOT(ISNUMBER(C{r})),D{r}="已完成"),"",C{r}-TODAY())')
        tk.cell(row=r, column=6, value=f'=IF(A{r}="","",IF(D{r}="已完成","✓",IF(NOT(ISNUMBER(E{r})),"",IF(E{r}<0,"逾期",IF(E{r}<=7,"快到期","")))))')
    grid(tk, 5, 4 + NT, 1, 7, cjk=Z, input_cols=(1, 2, 3, 4, 7), fmt={3: "yyyy-mm-dd"})
    dv_list(tk, '"未開始,進行中,已完成"', f"D5:D{4+NT}")
    tk.conditional_formatting.add(f"F5:F{4+NT}", FormulaRule(formula=['F5="逾期"'], fill=RED_FILL, font=Font(color="9B1C1C", bold=True)))
    tk.conditional_formatting.add(f"F5:F{4+NT}", FormulaRule(formula=['F5="快到期"'], fill=AMBER_FILL))
    tk.conditional_formatting.add(f"F5:F{4+NT}", FormulaRule(formula=['F5="✓"'], fill=GREEN_FILL))
    widths(tk, {"A": 26, "B": 11, "C": 12, "D": 10, "E": 10, "F": 9, "G": 22})
    tk.freeze_panes = "A5"
    landscape(tk); tk.print_area = "A1:G30"

    # 總覽
    PD_, PE_, PI_ = ref("人員名單", "D", P1, P2), ref("人員名單", "E", P1, P2), ref("人員名單", "I", P1, P2)
    TA, TD, TF = ref("籌備清單", "A", 5, 4 + NT), ref("籌備清單", "D", 5, 4 + NT), ref("籌備清單", "F", 5, 4 + NT)
    title(ov, None)
    ov["A1"] = "=活動設定!B4&\"｜總覽\""
    ov["A1"].font = font(Z, size=18, bold=True, color=DARK)
    ov["A2"] = '="活動日期："&TEXT(活動設定!B5,"yyyy/mm/dd")&" "&活動設定!B6&"　地點："&活動設定!B7&"　今天："&TEXT(TODAY(),"yyyy/mm/dd")'
    ov["A2"].font = font(Z, size=10, italic=True, color="6B7280")
    kpi(ov, 4, 1, "尾牙倒數（天）", '=IF(活動設定!B5-TODAY()<0,"已結束",活動設定!B5-TODAY())', "0", cjk=Z)
    kpi(ov, 4, 2, "名單人數", f'=COUNTIF({PB},"?*")', "0", cjk=Z)
    kpi(ov, 4, 3, "出席", f'=COUNTIF({PD_},"出席")', "0", cjk=Z)
    kpi(ov, 4, 4, "未回覆", f'=COUNTIFS({PB},"?*",{PD_},"未回覆")+COUNTIFS({PB},"?*",{PD_},"")', "0", cjk=Z)
    kpi(ov, 4, 5, "素食（含蛋奶素）", f'=COUNTIFS({PD_},"出席",{PE_},"素")+COUNTIFS({PD_},"出席",{PE_},"蛋奶素")', "0", cjk=Z)
    kpi(ov, 4, 6, "建議桌數", '=IF(N(活動設定!B9)=0,"",ROUNDUP(C5/活動設定!B9,0))', "0", cjk=Z)
    kpi(ov, 7, 1, "總預算", "=活動設定!B8", "#,##0", cjk=Z)
    kpi(ov, 7, 2, "預估支出", f"=預算表!E{5+NB}", "#,##0", cjk=Z)
    kpi(ov, 7, 3, "已花費（實際）", f"=預算表!F{5+NB}", "#,##0", cjk=Z)
    kpi(ov, 7, 4, "預算剩餘（依預估）", "=A8-B8", "#,##0;[Red]-#,##0", cjk=Z)
    kpi(ov, 7, 5, "獎品總額／名額", '=TEXT(獎品清單!L5,"#,##0")&" / "&獎品清單!L4', "@", cjk=Z)
    kpi(ov, 7, 6, "籌備進度", f'=IFERROR(COUNTIF({TD},"已完成")/COUNTIF({TA},"?*"),0)', "0%", cjk=Z)
    ov.conditional_formatting.add("D8", CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
    ov.conditional_formatting.add("D5", CellIsRule(operator="greaterThan", formula=["0"], fill=AMBER_FILL))
    ov.conditional_formatting.add("F8", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="F2C14E"))
    ov["A10"] = f'=IF(COUNTIF({TF},"逾期")>0,"⚠ 有 "&COUNTIF({TF},"逾期")&" 項籌備工作已逾期，請查看籌備清單。","")'
    ov["A10"].font = font(Z, size=11, bold=True, color="9B1C1C")
    header_row(ov, 12, ["部門", "名單人數", "出席", "未回覆", "素食", "可抽獎"], cjk=Z)
    for i in range(10):
        r = 13 + i
        ov.cell(row=r, column=1, value=f'=IF(活動設定!A{14+i}="","",活動設定!A{14+i})')
        ov.cell(row=r, column=2, value=f'=IF(A{r}="","",COUNTIFS({PC_},A{r},{PB},"?*"))')
        ov.cell(row=r, column=3, value=f'=IF(A{r}="","",COUNTIFS({PC_},A{r},{PD_},"出席"))')
        ov.cell(row=r, column=4, value=f'=IF(A{r}="","",COUNTIFS({PC_},A{r},{PB},"?*",{PD_},"未回覆"))')
        ov.cell(row=r, column=5, value=f'=IF(A{r}="","",COUNTIFS({PC_},A{r},{PD_},"出席",{PE_},"素")+COUNTIFS({PC_},A{r},{PD_},"出席",{PE_},"蛋奶素"))')
        ov.cell(row=r, column=6, value=f'=IF(A{r}="","",COUNTIFS({PC_},A{r},{PI_},"是"))')
    grid(ov, 13, 22, 1, 6, cjk=Z)
    ch = BarChart(); ch.type = "col"; ch.title = "各部門出席人數"; ch.style = 10
    ch.add_data(Reference(ov, min_col=3, min_row=12, max_row=19), titles_from_data=True)
    ch.set_categories(Reference(ov, min_col=1, min_row=13, max_row=19))
    ch.legend = None; ch.height = 7; ch.width = 13
    ch.series[0].graphicalProperties.solidFill = "F2C14E"
    ov.add_chart(ch, "H12")
    widths(ov, {"A": 16, "B": 14, "C": 16, "D": 18, "E": 18, "F": 14, "G": 3})
    landscape(ov); ov.page_setup.fitToHeight = 1; ov.print_area = "A1:O27"
    wb.active = 1
    wb.save(f"{OUT}/year-end-party-planner-tw.xlsx")


student(); content(); weiya()
print("ok")
