"""True North 450 customer, partner and legal document template.
Matches TrueNorth450_LOI_Template.pdf exactly: US Letter, Poppins, navy/copper, 58pt side margins.
Usage: from tn_docs import *; build(path, header_label, flowables)
"""
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
                                KeepTogether, Flowable)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor

FONT_DIR = os.environ.get('TN_FONT_DIR', '/usr/share/fonts/truetype/google-fonts/')
pdfmetrics.registerFont(TTFont('P', FONT_DIR + 'Poppins-Regular.ttf'))
pdfmetrics.registerFont(TTFont('PM', FONT_DIR + 'Poppins-Medium.ttf'))
pdfmetrics.registerFont(TTFont('PB', FONT_DIR + 'Poppins-Bold.ttf'))
registerFontFamily('P', normal='P', bold='PB', italic='P', boldItalic='PB')

NAVY = HexColor('#0B2545')       # text, titles, rules
COPPER = HexColor('#C4763A')     # title accent word, bullets
COPPER_D = HexColor('#9A5424')   # section numbers, list numbers
SLATE = HexColor('#4A5873')      # secondary text, labels, footer
PANEL = HexColor('#EEF1F8')      # summary panel, table header
LINE = HexColor('#D3D9E4')       # hairlines
LOGO = os.environ.get('TN_LOGO', os.path.join(os.path.dirname(__file__), 'logo.png'))

W, H = letter
L, R = 58, 555                    # 497pt text width
TAGLINE = 'Getting Social Security benefits to the people who need help most.'

body = ParagraphStyle('body', fontName='P', fontSize=10, leading=14.6, textColor=NAVY, spaceAfter=8)
note = ParagraphStyle('note', fontName='P', fontSize=8.5, leading=13, textColor=SLATE, spaceAfter=8)
h_sec = ParagraphStyle('sec', fontName='PB', fontSize=12.5, leading=16, textColor=NAVY, spaceBefore=12, spaceAfter=6)
item = ParagraphStyle('item', parent=body, spaceAfter=0)


class Title(Flowable):
    """'Letter of intent.' style: navy words then one copper word, 30pt Poppins Bold."""
    def __init__(self, navy, copper):
        super().__init__(); self.a, self.b = navy, copper
    def wrap(self, aw, ah): return aw, 36
    def draw(self):
        c = self.canv; c.setFont('PB', 30); c.setFillColor(NAVY); c.drawString(0, 8, self.a + ' ')
        c.setFillColor(COPPER); c.drawString(c.stringWidth(self.a + ' ', 'PB', 30), 8, self.b)


def prepared_for(name, lead='Prepared for:'):
    st = ParagraphStyle('pf', fontName='P', fontSize=15, leading=20, textColor=SLATE)
    return Paragraph(f'{lead} <font name="PB" color="#0B2545">{name}</font>', st)


class Hair(Flowable):
    def __init__(self, before=10, after=8, color=LINE, width=1):
        super().__init__(); self.b, self.a, self.c, self.w = before, after, color, width
    def wrap(self, aw, ah): self.aw = aw; return aw, self.b + self.a
    def draw(self):
        self.canv.setStrokeColor(self.c); self.canv.setLineWidth(self.w); self.canv.line(0, self.a, self.aw, self.a)


class Panel(Flowable):
    """Rounded summary panel: up to 4 label/value columns."""
    def __init__(self, cols):
        super().__init__(); self.cols = cols
        self.lab = ParagraphStyle('pl', fontName='PM', fontSize=8, leading=11, textColor=SLATE)
        self.val = ParagraphStyle('pv', fontName='PB', fontSize=10, leading=14, textColor=NAVY)
    def wrap(self, aw, ah):
        self.aw = aw; cw = (aw - 24) / len(self.cols); self.cw = cw
        self.ps = [(Paragraph(a, self.lab), Paragraph(b, self.val)) for a, b in self.cols]
        hv = max(v.wrap(cw - 10, 200)[1] for _, v in self.ps)
        self.h = 12 + 11 + 4 + hv + 14; return aw, self.h + 2
    def draw(self):
        c = self.canv; c.setFillColor(PANEL); c.roundRect(0, 2, self.aw, self.h, 6, stroke=0, fill=1)
        for i, (lab, val) in enumerate(self.ps):
            x = 12 + i * self.cw
            lab.wrap(self.cw - 10, 50); lab.drawOn(c, x, 2 + self.h - 12 - 11)
            _, hv = val.wrap(self.cw - 10, 200); val.drawOn(c, x, 2 + self.h - 12 - 11 - 4 - hv)


def section(num, title):
    return Paragraph(f'<font color="#9A5424">{num}</font>  {title}', h_sec)


def _row(marker, flow, indent, mstyle):
    t = Table([[Paragraph(marker, mstyle), flow]], colWidths=[indent, R - L - indent])
    t.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0),
                           ('RIGHTPADDING', (0, 0), (-1, -1), 0), ('TOPPADDING', (0, 0), (-1, -1), 0),
                           ('BOTTOMPADDING', (0, 0), (-1, -1), 3)]))
    return t


def numbered(items):
    """items: list of (bold lead or None, text). Copper-dark numerals, 16pt indent."""
    ms = ParagraphStyle('nm', fontName='PB', fontSize=10, leading=14.6, textColor=COPPER_D)
    return [_row(str(i), Paragraph((f'<font name="PB">{lead}</font> ' if lead else '') + text, item), 16, ms)
            for i, (lead, text) in enumerate(items, 1)]


def bullets(items):
    """items: list of (bold lead or None, text). Copper bullets, 12pt indent."""
    ms = ParagraphStyle('bm', fontName='PB', fontSize=10, leading=14.6, textColor=COPPER, leftIndent=1)
    return [_row('•', Paragraph((f'<font name="PB">{lead}</font> ' if lead else '') + text, item), 12, ms)
            for lead, text in items]


def terms_table(rows, header=('Term', 'Proposed'), first_col=129):
    hs = ParagraphStyle('th', fontName='PM', fontSize=8.4, leading=11, textColor=SLATE)
    k = ParagraphStyle('tk', fontName='PM', fontSize=9.2, leading=13, textColor=NAVY)
    v = ParagraphStyle('tv', fontName='P', fontSize=9.2, leading=13, textColor=NAVY)
    data = [[Paragraph(header[0], hs), Paragraph(header[1], hs)]] + [[Paragraph(a, k), Paragraph(b, v)] for a, b in rows]
    t = Table(data, colWidths=[first_col, R - L - first_col], repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), PANEL), ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('LINEBELOW', (0, 1), (-1, -1), 1, LINE), ('LEFTPADDING', (0, 0), (-1, -1), 6),
                           ('RIGHTPADDING', (0, 0), (-1, -1), 6), ('TOPPADDING', (0, 0), (-1, -1), 5),
                           ('BOTTOMPADDING', (0, 0), (-1, -1), 5)]))
    return t


def grid_table(rows, widths, num_cols=()):
    """Multi-column data table in the same style (header shaded, hairlines, first column medium)."""
    hs = ParagraphStyle('gh', fontName='PM', fontSize=8.4, leading=11, textColor=SLATE)
    hr = ParagraphStyle('ghr', parent=hs, alignment=2)
    k = ParagraphStyle('gk', fontName='PM', fontSize=9.2, leading=13, textColor=NAVY)
    v = ParagraphStyle('gv', fontName='P', fontSize=9.2, leading=13, textColor=NAVY)
    vr = ParagraphStyle('gvr', parent=v, alignment=2)
    data = []
    for i, r in enumerate(rows):
        data.append([c if isinstance(c, Flowable) else
                     Paragraph(str(c), (hr if j in num_cols else hs) if i == 0 else (k if j == 0 else (vr if j in num_cols else v)))
                     for j, c in enumerate(r)])
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), PANEL), ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                           ('LINEBELOW', (0, 1), (-1, -1), 1, LINE), ('LEFTPADDING', (0, 0), (-1, -1), 6),
                           ('RIGHTPADDING', (0, 0), (-1, -1), 6), ('TOPPADDING', (0, 0), (-1, -1), 5),
                           ('BOTTOMPADDING', (0, 0), (-1, -1), 5)]))
    return t


class Box(Flowable):
    def wrap(self, aw, ah): return 10, 12
    def draw(self):
        self.canv.setStrokeColor(NAVY); self.canv.setLineWidth(0.8); self.canv.rect(0, 1, 9, 9)


class Signatures(Flowable):
    """Two-column signature block with invisible e-signature anchors (\\s1\\ \\n1\\ \\t1\\ \\d1\\ ...)."""
    def __init__(self, left, right):
        super().__init__(); self.l, self.r = left, right
    def wrap(self, aw, ah): return aw, 175
    def draw(self):
        c = self.canv; top = 175
        for col, (x, w, name) in enumerate([(7, 231, self.l), (259, 230, self.r)]):
            c.setFont('PB', 11); c.setFillColor(NAVY); c.drawString(x, top - 12, name)
            for i, (lab, tag) in enumerate([('Signature', 's'), ('Name', 'n'), ('Title', 't'), ('Date', 'd')]):
                y = top - 47 - i * 40
                c.setFont('P', 7); c.setFillColorRGB(1, 1, 1); c.drawString(x + 2, y + 13, f'\\{tag}{col + 1}\\')
                c.setStrokeColor(NAVY); c.setLineWidth(1); c.line(x, y, x + w, y)
                c.setFont('P', 7.8); c.setFillColor(SLATE); c.drawString(x, y - 10, lab)


def _group(flow):
    """Keep each section heading with its content, and a closing note with the signature block."""
    out, cur = [], None
    for f in flow:
        if isinstance(f, Paragraph) and f.style.name == 'sec':
            if cur: out.append(KeepTogether(cur))
            cur = [f]
        elif isinstance(f, Signatures) and cur is not None and len(cur) > 0 and isinstance(cur[-1], Paragraph):
            cur = cur[:-1] + [KeepTogether([cur[-1], f])]
        elif cur is not None:
            cur.append(f)
        else:
            out.append(f)
    if cur: out.append(KeepTogether(cur))
    return out


def build(path, label, flow, footer_note='Draft for discussion', title=None):
    flow = _group(flow)
    def page(c, doc):
        c.saveState()
        c.drawImage(LOGO, L, H - 42 - 76, 76, 76, mask='auto')
        c.setFont('PB', 8.5); c.setFillColor(NAVY); c.drawRightString(R, H - 90 - 7, label)
        c.setFont('P', 7); c.setFillColor(SLATE); c.drawString(L, H - 724 - 6, TAGLINE)
        c.setStrokeColor(NAVY); c.setLineWidth(1); c.line(L, H - 735.5, R - 1, H - 735.5)
        c.setFont('P', 7.5); c.drawString(L, H - 743 - 7, f'Page {doc.page} · {footer_note}')
        c.setFont('PB', 7.5); c.setFillColor(NAVY); c.drawRightString(R, H - 743 - 7, 'True North 450')
        c.restoreState()
    doc = BaseDocTemplate(path, pagesize=letter, title=title or label, author='True North 450, Inc.')
    frame = Frame(L, 82, R - L, H - 120 - 82, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id='p', frames=[frame], onPage=page)])
    doc.build(flow)
