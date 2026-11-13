"""GTM update slides built inside the real seed deck: clone slides 10 and 8, swap text only."""
import copy
from pptx import Presentation
from pptx.util import Pt

SRC = '/home/claude/seed/seed.pptx'
OUT = '/home/claude/out/TrueNorth450_GTMUpdate_Nov2026.pptx'
FOOT = 'True North 450 · GTM update · November 2026'
prs = Presentation(SRC)
R_IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'


def clone(n):
    src = prs.slides[n - 1]
    dst = prs.slides.add_slide(src.slide_layout)
    for shp in list(dst.shapes):
        shp._element.getparent().remove(shp._element)
    rmap = {}
    for rid, rel in src.part.rels.items():
        if rel.reltype == R_IMG:
            rmap[rid] = dst.part.relate_to(rel.target_part, R_IMG)
    for el in src.shapes._spTree.iterchildren():
        if el.tag.endswith('}nvGrpSpPr') or el.tag.endswith('}grpSpPr'):
            continue
        e = copy.deepcopy(el)
        for node in e.iter():
            for k, v in list(node.attrib.items()):
                if k.endswith('}embed') and v in rmap:
                    node.set(k, rmap[v])
        dst.shapes._spTree.append(e)
    return dst


def put(shape, *parts):
    """Replace text keeping run formatting. parts: strings, one per existing run (extra runs dropped)."""
    p = shape.text_frame.paragraphs[0]
    runs = p.runs
    for r, t in zip(runs, parts):
        r.text = t
    for r in runs[len(parts):]:
        r._r.getparent().remove(r._r)
    for extra in shape.text_frame.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)


# ---------- Slide A: clone of seed slide 10 (four stat tiles) ----------
a = clone(10)
