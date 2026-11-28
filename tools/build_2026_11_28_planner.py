"""2027 printable planner PDF (Letter & A4, Sunday- or Monday-start)."""
import sys, calendar, datetime as dt
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.colors import HexColor

OUT = sys.argv[1]
DARK = HexColor("#2F3E46"); CORN = HexColor("#F2C14E"); CREAM = HexColor("#FFF6DD")
LINE = HexColor("#C9CED6"); MUTED = HexColor("#6B7280"); LIGHT = HexColor("#F3F4F6")
Y = 2027
MONTHS = list(calendar.month_name)[1:]


def header(c, W, H, title, sub=None, m=36):
    c.setFillColor(DARK); c.rect(0, H - 64, W, 64, stroke=0, fill=1)
    c.setFillColor(CORN); c.rect(m, H - 64, 8, 64, stroke=0, fill=1)
    c.setFillColor(HexColor("#FFFFFF")); c.setFont("Helvetica-Bold", 22)
    c.drawString(m + 20, H - 42, title)
    if sub:
        c.setFont("Helvetica", 11); c.setFillColor(CORN)
        c.drawRightString(W - m, H - 40, sub)


def footer(c, W, m=36):
    c.setFont("Helvetica", 7); c.setFillColor(MUTED)
    c.drawRightString(W - m, 18, "Corn_Can  |  2027 Planner")


def lines(c, x, y_top, w, y_bot, step=18):
    c.setStrokeColor(LINE); c.setLineWidth(0.5)
    y = y_top - step
    while y >= y_bot:
        c.line(x, y, x + w, y); y -= step


def box(c, x, y, w, h, label=None, fill=None):
    if fill:
        c.setFillColor(fill); c.rect(x, y, w, h, stroke=0, fill=1)
    c.setStrokeColor(LINE); c.setLineWidth(0.8); c.rect(x, y, w, h, stroke=1, fill=0)
    if label:
        c.setFillColor(DARK); c.setFont("Helvetica-Bold", 9); c.drawString(x + 6, y + h - 13, label)


def month_matrix(year, month, first):
    cal = calendar.Calendar(firstweekday=first)
    return cal.monthdayscalendar(year, month)


def dow_names(first):
    names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    return names[first:] + names[:first]


def cover(c, W, H, first):
    c.setFillColor(CREAM); c.rect(0, 0, W, H, stroke=0, fill=1)
    c.setFillColor(DARK); c.rect(0, H * 0.42, W, H * 0.58, stroke=0, fill=1)
    c.setFillColor(CORN); c.setFont("Helvetica-Bold", 120); c.drawCentredString(W / 2, H * 0.68, str(Y))
    c.setFillColor(HexColor("#FFFFFF")); c.setFont("Helvetica-Bold", 30)
    c.drawCentredString(W / 2, H * 0.58, "Monthly & Weekly Planner")
    c.setFont("Helvetica", 13); c.setFillColor(CORN)
    c.drawCentredString(W / 2, H * 0.53, ("Sunday" if first == 6 else "Monday") + " start  ·  Printable")
    c.setFillColor(DARK); c.setFont("Helvetica", 12)
    c.drawCentredString(W / 2, H * 0.30, "This planner belongs to")
    c.setStrokeColor(DARK); c.line(W * 0.25, H * 0.25, W * 0.75, H * 0.25)
    c.setFont("Helvetica-Bold", 10); c.drawCentredString(W / 2, 40, "Corn_Can")
    c.showPage()


def year_glance(c, W, H, first, m=36):
    header(c, W, H, f"{Y} at a Glance")
    gw = (W - 2 * m - 2 * 18) / 3; gh = (H - 64 - 40 - 50 - 3 * 14) / 4
    for i in range(12):
        col, row = i % 3, i // 3
        x = m + col * (gw + 18); top = H - 64 - 30 - row * (gh + 14)
        c.setFillColor(DARK); c.setFont("Helvetica-Bold", 12); c.drawString(x, top - 12, MONTHS[i])
        cw = gw / 7; c.setFont("Helvetica-Bold", 7); c.setFillColor(MUTED)
        for j, d in enumerate(dow_names(first)):
            c.drawCentredString(x + cw * j + cw / 2, top - 28, d[:2])
        c.setFont("Helvetica", 9); c.setFillColor(DARK)
        rh = (gh - 40) / 6
        for r, wk in enumerate(month_matrix(Y, i + 1, first)):
            for j, d in enumerate(wk):
                if d:
                    c.drawCentredString(x + cw * j + cw / 2, top - 42 - r * rh, str(d))
    footer(c, W); c.showPage()


def goals(c, W, H, m=36):
    header(c, W, H, f"{Y} Goals & Intentions")
    top = H - 64 - 24; bw = (W - 2 * m - 14) / 2; bh = (top - 50 - 14 * 2) / 3
    labels = ["Personal", "Health & Fitness", "Career & Learning", "Money & Savings", "Home & Family", "Fun & Travel"]
    for i, lab in enumerate(labels):
        col, row = i % 2, i // 2
        x = m + col * (bw + 14); y = top - (row + 1) * bh - row * 14
        box(c, x, y, bw, bh, lab, None)
        lines(c, x + 8, y + bh - 20, bw - 16, y + 8, 20)
    footer(c, W); c.showPage()


def month_page(c, W, H, mi, first, m=36):
    header(c, W, H, MONTHS[mi], str(Y))
    side = 150 if W > H else 0
    top = H - 64 - 20; bottom = 160
    gw = W - 2 * m; cw = gw / 7
    c.setFont("Helvetica-Bold", 9); c.setFillColor(DARK)
    for j, d in enumerate(dow_names(first)):
        c.setFillColor(LIGHT); c.rect(m + j * cw, top - 18, cw, 18, stroke=0, fill=1)
        c.setFillColor(DARK); c.drawCentredString(m + j * cw + cw / 2, top - 13, d.upper())
    weeks = month_matrix(Y, mi + 1, first)
    rows = len(weeks); rh = (top - 18 - bottom) / rows
    c.setStrokeColor(LINE); c.setLineWidth(0.8)
    for r, wk in enumerate(weeks):
        for j, d in enumerate(wk):
            x = m + j * cw; y = top - 18 - (r + 1) * rh
            c.setFillColor(HexColor("#FAFAFA") if not d else HexColor("#FFFFFF"))
            c.rect(x, y, cw, rh, stroke=1, fill=1)
            if d:
                c.setFillColor(DARK); c.setFont("Helvetica-Bold", 11); c.drawString(x + 5, y + rh - 14, str(d))
    # bottom: priorities + notes
    bw = (gw - 14) / 2
    box(c, m, 40, bw, bottom - 54, "Top priorities this month")
    for k in range(5):
        yy = 40 + bottom - 54 - 32 - k * 16
        c.setStrokeColor(DARK); c.rect(m + 8, yy, 8, 8, stroke=1, fill=0)
        c.setStrokeColor(LINE); c.line(m + 22, yy, m + bw - 8, yy)
    box(c, m + bw + 14, 40, bw, bottom - 54, "Notes")
    lines(c, m + bw + 22, 40 + bottom - 54 - 18, bw - 16, 46, 16)
    footer(c, W); c.showPage()


def week_page(c, W, H, start, m=36):
    end = start + dt.timedelta(days=6)
    t = f"Week of {start.strftime('%b')} {start.day}" + (f", {start.year}" if start.year != Y else "")
    header(c, W, H, t, f"{start.strftime('%b %d')} – {end.strftime('%b %d, %Y')}")
    top = H - 64 - 18; gw = W - 2 * m
    colw = (gw - 14) / 2
    bh = (top - 40 - 3 * 10) / 4
    days = [start + dt.timedelta(days=i) for i in range(7)]
    slots = [(0, 0), (0, 1), (0, 2), (0, 3), (1, 0), (1, 1), (1, 2)]
    for d, (col, row) in zip(days, slots):
        x = m + col * (colw + 14); y = top - (row + 1) * bh - row * 10
        box(c, x, y, colw, bh)
        c.setFillColor(CORN); c.rect(x, y + bh - 20, colw, 20, stroke=0, fill=1)
        c.setFillColor(DARK); c.setFont("Helvetica-Bold", 10)
        c.drawString(x + 6, y + bh - 14, d.strftime("%A").upper())
        c.setFont("Helvetica", 10); c.drawRightString(x + colw - 6, y + bh - 14, f"{d.strftime('%b')} {d.day}")
        lines(c, x + 6, y + bh - 20, colw - 12, y + 4, 17)
    # to-do box in the last slot
    x = m + colw + 14; y = top - 4 * bh - 3 * 10
    box(c, x, y, colw, bh, "To-do / Notes", HexColor("#FFFBEA"))
    for k in range(int((bh - 26) // 17)):
        yy = y + bh - 34 - k * 17
        c.setStrokeColor(DARK); c.rect(x + 8, yy, 7, 7, stroke=1, fill=0)
        c.setStrokeColor(LINE); c.line(x + 20, yy, x + colw - 8, yy)
    footer(c, W); c.showPage()


def notes(c, W, H, m=36):
    header(c, W, H, "Notes")
    lines(c, m, H - 64 - 10, W - 2 * m, 40, 20)
    footer(c, W); c.showPage()


def build(size, first, fname):
    W, H = size
    c = canvas.Canvas(f"{OUT}/{fname}", pagesize=size)
    c.setTitle(f"{Y} Monthly & Weekly Planner – Corn_Can"); c.setAuthor("Corn_Can")
    cover(c, W, H, first)
    year_glance(c, W, H, first)
    goals(c, W, H)
    jan1 = dt.date(Y, 1, 1)
    start_dow = 6 if first == 6 else 0  # python weekday: Mon=0..Sun=6
    s = jan1 - dt.timedelta(days=(jan1.weekday() - start_dow) % 7)
    weeks = []
    while s <= dt.date(Y, 12, 31):
        weeks.append(s); s += dt.timedelta(days=7)
    for mi in range(12):
        month_page(c, W, H, mi, first)
        for w in weeks:
            # place each week after the month in which it starts (first week goes after January)
            ws_month = w.month if w.year == Y else 1
            if ws_month == mi + 1:
                week_page(c, W, H, w)
    for _ in range(3):
        notes(c, W, H)
    c.save()
    return len(weeks)


for size, sname in ((letter, "letter"), (A4, "a4")):
    for first, fn in ((6, "sunday"), (0, "monday")):
        n = build(size, first, f"planner-2027-{sname}-{fn}-start.pdf")
        print(sname, fn, "weeks", n)
