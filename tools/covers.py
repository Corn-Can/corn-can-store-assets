from PIL import Image, ImageDraw, ImageFont, ImageFilter
import os

HI = "/home/claude/out/pdf/hi/"
OUT = "/home/claude/out/img/"
os.makedirs(OUT, exist_ok=True)
DARK = (47, 62, 70)
CORN = (242, 193, 78)
CREAM = (255, 246, 221)
WHITE = (255, 255, 255)
SANS = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
BOLD = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
TC = 3  # index of Traditional Chinese face in the .ttc collection


def F(size, bold=False):
    return ImageFont.truetype(BOLD if bold else SANS, size, index=TC)


def shot(name, top=0.0, bottom=1.0, left=0.0, right=1.0):
    im = Image.open(HI + name).convert("RGB")
    w, h = im.size
    return im.crop((int(w * left), int(h * top), int(w * right), int(h * bottom)))


def card(canvas, im, box, radius=14):
    """Paste screenshot scaled to fit box (x, y, w, h) with border + shadow."""
    x, y, bw, bh = box
    s = min(bw / im.width, bh / im.height)
    im = im.resize((int(im.width * s), int(im.height * s)), Image.LANCZOS)
    sh = Image.new("RGBA", (im.width + 40, im.height + 40), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle((20, 24, im.width + 20, im.height + 24), radius, fill=(0, 0, 0, 90))
    sh = sh.filter(ImageFilter.GaussianBlur(10))
    canvas.paste(sh, (x - 20, y - 20), sh)
    frame = Image.new("RGB", (im.width + 12, im.height + 12), WHITE)
    frame.paste(im, (6, 6))
    mask = Image.new("L", frame.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, frame.width, frame.height), radius, fill=255)
    canvas.paste(frame, (x - 6, y - 6), mask)
    return im.width, im.height


def wrap(draw, text, fnt, maxw):
    lines, cur = [], ""
    tokens = list(text) if any("一" <= ch <= "鿿" for ch in text) else text.split(" ")
    sep = "" if len(tokens) and len(tokens[0]) <= 1 and any("一" <= ch <= "鿿" for ch in text) else " "
    for t in tokens:
        trial = (cur + sep + t) if cur else t
        if draw.textlength(trial, font=fnt) <= maxw:
            cur = trial
        else:
            lines.append(cur)
            cur = t
    if cur:
        lines.append(cur)
    return lines


def text_block(draw, xy, text, fnt, fill, maxw, gap=1.18):
    x, y = xy
    for line in wrap(draw, text, fnt, maxw):
        draw.text((x, y), line, font=fnt, fill=fill)
        y += int(fnt.size * gap)
    return y


def brand(draw, x, y, color=WHITE, size=26):
    draw.rounded_rectangle((x, y + 4, x + 22, y + 26), 6, fill=CORN)
    draw.text((x + 32, y - 2), "Corn_Can", font=F(size, True), fill=color)


def main_cover(p):
    W, H = 1280, 720
    c = Image.new("RGB", (W, H), CREAM)
    d = ImageDraw.Draw(c)
    d.rectangle((0, 0, 500, H), fill=DARK)
    d.rounded_rectangle((44, 44, 44 + d.textlength(p["tag"], font=F(22, True)) + 32, 86), 21, fill=CORN)
    d.text((60, 47), p["tag"], font=F(22, True), fill=DARK)
    y = text_block(d, (44, 110), p["title"], F(p.get("tsize", 50), True), WHITE, 420, 1.15)
    y = text_block(d, (44, y + 14), p["sub"], F(24), CORN, 420, 1.3)
    y += 18
    for b in p["bullets"][:4]:
        d.ellipse((46, y + 12, 56, y + 22), fill=CORN)
        y = text_block(d, (70, y), b, F(21), WHITE, 400, 1.3) + 6
    brand(d, 44, H - 70)
    d.text((W - 24, H - 34), p["apps"], font=F(18), fill=DARK, anchor="rs")
    a, b2 = p["shots"][0], p["shots"][1]
    card(c, a, (560, 60, 680, 380))
    card(c, b2, (640, 330, 600, 320))
    c.save(OUT + p["slug"] + "-cover.png", optimize=True)


def pin_cover(p):
    W, H = 1000, 1500
    c = Image.new("RGB", (W, H), CREAM)
    d = ImageDraw.Draw(c)
    d.rectangle((0, 0, W, 470), fill=DARK)
    d.rounded_rectangle((60, 56, 60 + d.textlength(p["tag"], font=F(30, True)) + 44, 112), 28, fill=CORN)
    d.text((82, 60), p["tag"], font=F(30, True), fill=DARK)
    y = text_block(d, (60, 140), p["pin_title"], F(p.get("ptsize", 72), True), WHITE, 880, 1.12)
    text_block(d, (60, y + 22), p["sub"], F(32), CORN, 880, 1.3)
    a, b2 = p["shots"][0], p["shots"][1]
    _, h1 = card(c, a, (60, 530, 880, 400))
    y2 = 530 + h1 + 50
    _, h2 = card(c, b2, (100, y2, 840, 1210 - y2))
    y = max(1250, y2 + h2 + 50)
    for b in p["bullets"][:3]:
        d.ellipse((64, y + 15, 78, y + 29), fill=CORN)
        y = text_block(d, (96, y), b, F(30), DARK, 840, 1.25) + 4
    brand(d, 60, H - 62, DARK, 30)
    d.text((W - 60, H - 30), p["apps"], font=F(24), fill=DARK, anchor="rs")
    c.save(OUT + p["slug"] + "-pin.png", optimize=True)


products = [
    dict(slug="subscription-tracker", tag="Excel template",
         title="Subscription Tracker", pin_title="Subscription Tracker Spreadsheet",
         sub="See what every subscription really costs – and when it renews.",
         bullets=["Monthly & yearly cost of every subscription", "Renewal alerts: 7 days red, 30 days amber",
                  "Free-trial warnings & 'Worth it?' ratings", "Dashboard with category chart"],
         apps="Excel · Google Sheets · LibreOffice",
         shots=[shot("sub_dash-4_c.png", 0, 1, 0, 1), shot("sub_list-1_c.png", 0, 0.25, 0, 1)]),
    dict(slug="trip-planner", tag="Excel template",
         title="Trip Planner & Travel Budget", pin_title="Trip Planner & Travel Budget Spreadsheet", tsize=46, ptsize=66,
         sub="Itinerary, expenses, budget and who-owes-who in one file.",
         bullets=["Day-by-day itinerary up to 30 days", "Expense log with currency conversion",
                  "Budget vs actual + settle up for groups", "38-item packing list with progress"],
         apps="Excel · Google Sheets · LibreOffice",
         shots=[shot("trip_budget-15_c.png"), shot("trip_itin-02_c.png", 0, 0.3, 0, 1)]),
    dict(slug="invoice-lottery", tag="Excel 範本",
         title="統一發票對獎紀錄表", pin_title="統一發票對獎紀錄表", tsize=50, ptsize=80,
         sub="輸入發票號碼，自動對獎、算獎金、提醒領獎期限。",
         bullets=["自動判斷期別與獎別（特別獎～六獎）", "預估實領：所得稅與印花稅", "領獎期限倒數，過期紅字提醒",
                  "已填入 115年07-08月中獎號碼"],
         apps="Excel · Google 試算表 · LibreOffice",
         shots=[shot("inv_list-02_c.png", 0, 0.29, 0, 1), shot("inv_win-01_c.png", 0, 1, 0, 1)]),
]
for p in products:
    main_cover(p)
    pin_cover(p)
print(os.listdir(OUT))
