"""
Build a synthetic MDF-shaped PDF to exercise cases the one real sample does not
cover: a data table long enough to span a page break (with the header repeated),
a merged full-width row, and stacked label/value boxes either side of the break.
"""
from reportlab.lib.colors import Color, black
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas

BLUE = Color(0.678431372, 0.847058823, 0.901960784)
INK = Color(0.251, 0.255, 0.263)
W, H = landscape(A4)
L, R = 39.7, 809.2
BODY, SUB, SEC, TITLE = 6.624, 8.153, 10.191, 13.758


class Form:
    def __init__(self, path):
        self.c = canvas.Canvas(path, pagesize=landscape(A4))
        self.y = 28.5

    def newpage(self):
        self.c.showPage()
        self.y = 28.5

    def room(self, need):
        if self.y + need > H - 30:
            self.newpage()
            return True
        return False

    def text(self, s, size, bold=False, x=L):
        self.room(size * 2)
        self.y += size * 1.6
        self.c.setFillColor(INK)
        self.c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        self.c.drawString(x, H - self.y, s)
        self.y += size * 0.4

    def row(self, cols, widths, fills, bold=None, centre=None, h=15.3):
        """Draw one table row as filled rects plus hairline borders."""
        self.room(h + 2)
        top = self.y
        x = L
        for i, (txt, w) in enumerate(zip(cols, widths)):
            if fills[i]:
                self.c.setFillColor(BLUE)
                self.c.rect(x, H - top - h, w, h, stroke=0, fill=1)
            self.c.setFillColor(black)
            for rx, ry, rw, rh in ((x, H - top - h, 0.51, h), (x + w - 0.51, H - top - h, 0.51, h),
                                   (x, H - top - h, w, 0.51), (x, H - top - 0.51, w, 0.51)):
                self.c.rect(rx, ry, rw, rh, stroke=0, fill=1)
            self.c.setFillColor(INK)
            b = bold[i] if bold else False
            self.c.setFont("Helvetica-Bold" if b else "Helvetica", BODY)
            tw = self.c.stringWidth(txt, "Helvetica-Bold" if b else "Helvetica", BODY)
            tx = x + (w - tw) / 2 if (centre and centre[i]) else x + 3
            self.c.drawString(tx, H - top - h + 4.5, txt)
            x += w
        self.y = top + h

    def gap(self, n=7.2):
        self.y += n

    def save(self):
        self.c.save()


f = Form("fixture_long.pdf")
f.text("MDF-2026/27", BODY, True)
f.text("Module development form", TITLE, True)
f.text("Synthetic fixture used to exercise the converter.", BODY)
f.text("Module details", SEC, True)

FW = [193.0, 576.5]
for label, value in [("Module code", "XX9Z9999"), ("CIS unit code", "ZZZ001"),
                     ("Module title", "A Module With A Considerably Longer Title Than Usual"),
                     ("Credit value", "40")]:
    f.row([label, value], FW, [True, False], bold=[True, False])
    f.gap()

f.text("Learning and teaching", SEC, True)
f.text("Learning outcomes", SUB, True)

LO = [62.0, 44.5, 663.0]
HEAD = ["ID", "Type", "Learning outcome (by the end of the module the student will be able to...)"]


def header():
    f.row(HEAD, LO, [True, True, True], bold=[True, True, True],
          centre=[True, True, True])


header()
for i in range(1, 29):                       # long enough to cross a page break
    if f.y + 17 > H - 30:
        f.newpage()
        header()                             # the browser print repeats the header
    f.row(["%010d" % i, "MODULE",
           "Outcome %d: demonstrate an understanding of the topic under study." % i],
          LO, [False, False, False])

f.text("Requisites", SEC, True)
RQ = [213.0, 202.3, 101.9, 182.5, 69.6]
f.row(["Type of requisite", "Rule taking flag", "Module", "Module status", "Rule"],
      RQ, [True] * 5, bold=[True] * 5, centre=[True] * 5)
f.row(["None Found"], [sum(RQ)], [False], centre=[True])   # merged full-width row

f.save()
print("wrote fixture_long.pdf")
