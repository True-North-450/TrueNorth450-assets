"""True North 450 deck template (matches TrueNorth450_Oct2026.pdf).
16:9, 960x540pt, white slides, Poppins. Editable PowerPoint built with python-pptx.
Usage: d = Deck('True North 450 · PHCA · November 2026'); s = d.slide('EYEBROW', 'Navy part', 'copper part.'); ...; d.save(path)
"""
import os
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from PIL import ImageFont

FONT_DIR = os.environ.get('TN_FONT_DIR', '/usr/share/fonts/truetype/google-fonts/')
LOGO = os.environ.get('TN_LOGO', os.path.join(os.path.dirname(__file__), 'logo.png'))
NAVY, COPPER = RGBColor(0x0B, 0x25, 0x45), RGBColor(0xC4, 0x76, 0x3A)
SKY, LINE = RGBColor(0xEE, 0xF1, 0xF8), RGBColor(0xD9, 0xDE, 0xE5)
GRAY, INK, WHITE = RGBColor(0x6B, 0x72, 0x80), RGBColor(0x1F, 0x29, 0x37), RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xC9, 0xD3, 0xE3)
P = Pt
_bold = ImageFont.truetype(FONT_DIR + 'Poppins-Bold.ttf', 300)


def _nostyle(shp):
    st = shp._element.find(qn('p:style'))
    if st is not None: shp._element.remove(st)
    return shp


def text_width(t, size):
    return _bold.getlength(t) * size / 300


class Deck:
    def __init__(self, footer):
        self.prs = Presentation(); self.prs.slide_width = Emu(12192000); self.prs.slide_height = Emu(6858000)
        self.blank = self.prs.slide_layouts[6]; self.footer = footer; self.n = 0

    # ---------- primitives ----------
    def text(self, s, x, y, w, h, paras, size=13, color=INK, bold=False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, spacing=0, tracking=0, line=None):
        """paras: str | list of paragraphs; a paragraph is str or list of (text, color, bold, size)."""
        box = s.shapes.add_textbox(P(x), P(y), P(w), P(h)); tf = box.text_frame; tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0; tf.vertical_anchor = anchor
        for i, para in enumerate(paras if isinstance(paras, list) else [paras]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); p.alignment = align
            if spacing: p.space_after = P(spacing)
            if line: p.line_spacing = line
            for seg in (para if isinstance(para, list) else [(para,)]):
                t, c, b, sz = (list(seg) + [None] * 3)[:4]
                r = p.add_run(); r.text = t; f = r.font; f.name = 'Poppins'
                f.size = P(sz or size); f.bold = bold if b is None else b; f.color.rgb = c or color
                if tracking: r._r.get_or_add_rPr().set('spc', str(int(tracking * 100)))
        return box

    def rect(self, s, x, y, w, h, fill=None, line=None, lw=1.25, radius=True):
        shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE, P(x), P(y), P(w), P(h))
        if radius: shp.adjustments[0] = min(0.5, 7.0 / min(w, h))
        if fill is None: shp.fill.background()
        else: shp.fill.solid(); shp.fill.fore_color.rgb = fill
        if line is None: shp.line.fill.background()
        else: shp.line.color.rgb = line; shp.line.width = P(lw)
        shp.shadow.inherit = False
        return _nostyle(shp)

    def hline(self, s, x1, y, x2, color, w):
        return self.rect(s, x1, y - w / 2, x2 - x1, w, color, None, radius=False)

    def circle(self, s, x, y, d, fill, label=None, size=14):
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, P(x), P(y), P(d), P(d)); c.fill.solid(); c.fill.fore_color.rgb = fill
        c.line.fill.background(); c.shadow.inherit = False; _nostyle(c)
        if label is not None:
            self.text(s, x, y, d, d, str(label), size=size, color=WHITE, bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        return c

    # ---------- components ----------
    def card(self, s, x, y, w, h, kind='outline', title=None, body=None, tag=None, center=False,
             title_size=13, body_size=10.5, pad=18):
        """kind: outline (white, navy border) | accent (white, copper border) | panel (sky fill) | navy (filled)."""
        fill, line = {'outline': (WHITE, NAVY), 'accent': (WHITE, COPPER), 'panel': (SKY, LINE), 'navy': (NAVY, LINE)}[kind]
        self.rect(s, x, y, w, h, fill, line, 1.25 if kind in ('outline', 'accent') else 0.75)
        tc = WHITE if kind == 'navy' else NAVY; bc = SOFT if kind == 'navy' else INK
        al = PP_ALIGN.CENTER if center else PP_ALIGN.LEFT
        ty = y + pad
        if tag: self.pill(s, x + w - pad - 70, y + 10, 70, tag, COPPER if kind == 'accent' else NAVY); ty += 8
        paras = []
        if title: paras.append([(title, tc, True, title_size)])
        if body: paras.append([(body, bc, False, body_size)])
        if paras: self.text(s, x + pad, ty, w - 2 * pad, h - (ty - y) - pad, paras, align=al, spacing=6, line=1.15)

    def pill(self, s, x, y, w, label, fill=NAVY):
        self.rect(s, x, y, w, 13, fill, None)
        self.text(s, x, y, w, 13, label.upper(), size=6.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE, tracking=0.6)

    def band(self, s, x, y, w, label, fill=NAVY):
        self.rect(s, x, y, w, 24, fill, None, radius=True)
        self.text(s, x, y, w, 24, label.upper(), size=9, color=WHITE, bold=True, align=PP_ALIGN.CENTER,
                  anchor=MSO_ANCHOR.MIDDLE, tracking=1.5)

    def stat(self, s, x, y, w, h, number, caption, kind='panel', num_color=None, num_size=38):
        fill, line = {'panel': (SKY, LINE), 'navy': (NAVY, LINE), 'outline': (WHITE, NAVY)}[kind]
        self.rect(s, x, y, w, h, fill, line, 0.75 if kind != 'outline' else 1.25)
        nc = num_color or (COPPER if kind == 'navy' else NAVY)
        self.text(s, x + 18, y + 40, w - 36, num_size + 14, number, size=num_size, color=nc, bold=True)
        self.text(s, x + 18, y + 54 + num_size, w - 36, h - 70 - num_size, caption, size=14,
                  color=WHITE if kind == 'navy' else NAVY, bold=True, line=1.1)

    # ---------- slides ----------
    def slide(self, eyebrow, navy, copper='', source=''):
        s = self.prs.slides.add_slide(self.blank); self.n += 1
        s.shapes.add_picture(LOGO, P(32), P(25), P(54), P(54))
        self.text(s, 97, 21, 820, 28, eyebrow.upper(), size=20, color=COPPER, bold=True, tracking=2)
        full = navy + (' ' + copper if copper else '')
        two = text_width(full, 30) > 820
        self.text(s, 97, 52, 830, 80, [[(navy + (' ' if copper else ''), NAVY, True, 30), (copper, COPPER, True, 30)]], line=1.0)
        self.top = 155 if two else 120
        self.hline(s, 32, 140 if two else 104, 927, NAVY, 1.5)
        self.hline(s, 32, 503, 927, LINE, 0.75)
        self.text(s, 32, 512, 820, 14, self.footer + (' · ' + source if source else ''), size=8, color=GRAY)
        self.text(s, 880, 512, 47, 14, str(self.n), size=8, color=GRAY, bold=True, align=PP_ALIGN.RIGHT)
        return s

    def title_slide(self, kicker, headline_navy, headline_copper, rows):
        """Cover in the Oct 2026 style: sky background, logo lockup, copper bar, 44pt headline, label/value rows."""
        s = self.prs.slides.add_slide(self.blank); self.n += 1
        bg = s.background.fill; bg.solid(); bg.fore_color.rgb = SKY
        s.shapes.add_picture(LOGO, P(43), P(30), P(86), P(86))
        self.text(s, 143, 50, 600, 24, 'True North 450', size=18, color=NAVY, bold=True)
        self.text(s, 143, 77, 700, 20, kicker, size=12.5, color=GRAY, bold=True)
        self.rect(s, 72, 158, 79, 5, COPPER, None, radius=False)
        self.text(s, 72, 192, 816, 120, [[(headline_navy, NAVY, True, 44)], [(headline_copper, COPPER, True, 44)]], line=1.05)
        self.hline(s, 72, 320, 888, NAVY, 1.25)
        y = 340
        for i, (lab, val) in enumerate(rows):
            self.text(s, 72, y + 22, 150, 20, lab.upper(), size=13, color=COPPER, bold=True, tracking=1)
            self.text(s, 238, y + 10, 650, 60, val, size=22, color=NAVY, bold=True, line=1.05)
            y += 79
            if i < len(rows) - 1: self.hline(s, 72, y, 888, LINE, 0.75)
        return s

    def save(self, path):
        self.prs.save(path)
