#!/usr/bin/env python3
"""Build an editable, Canva-importable version of the roll-up banner.

Canva imports .pptx files as fully editable designs (text boxes, shapes, images),
so this script lays the banner out as ONE slide in the banner's 850 x 2000 mm
proportions using fonts that exist in Canva's library (Fraunces + Manrope) and the
2x upscaled figures in assets/hires/.

The PowerPoint file format caps a slide at 56 in (1422 mm) per side, so the slide
is emitted at 70 % scale (595 x 1400 mm) — the largest in-spec size. Everything is
vector/text, so printing at 850 x 2000 mm (143 %) or resizing in Canva loses nothing.

Usage:   python3 build_canva_pptx.py          # 595 x 1400 mm  (recommended for Canva)
         python3 build_canva_pptx.py --half   # 425 x 1000 mm  (smaller file, same layout)
Requires: pip install python-pptx pillow
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

HERE = Path(__file__).parent
HIRES = HERE.parent / "assets" / "hires"

S = 0.5 if "--half" in sys.argv else 0.7            # global scale factor (0.7 = OOXML 56-inch limit)
W_MM, H_MM = 850, 2000                               # design coordinates are always in banner mm
OUT = HERE / f"giant-axonal-neuropathy-banner-canva-{W_MM * S:.0f}x{H_MM * S:.0f}mm.pptx"

# ------------------------------------------------------------------ theme
PAPER, PAPER_DEEP = RGBColor(0xF6, 0xF3, 0xEC), RGBColor(0xEC, 0xE7, 0xDB)
INK, INK_SOFT = RGBColor(0x14, 0x21, 0x3D), RGBColor(0x4A, 0x55, 0x70)
NAVY, NAVY_2, NAVY_GLOW = RGBColor(0x0F, 0x1B, 0x33), RGBColor(0x18, 0x27, 0x47), RGBColor(0x12, 0x30, 0x45)
TEAL, TEAL_SOFT, TEAL_LIGHT = RGBColor(0x1A, 0x9C, 0x8E), RGBColor(0xDD, 0xEF, 0xEC), RGBColor(0x7F, 0xE0, 0xD4)
CORAL, GOLD, GOLD_DARK = RGBColor(0xE4, 0x57, 0x2E), RGBColor(0xE9, 0xB4, 0x4C), RGBColor(0xB7, 0x80, 0x1E)
WHITE, LINE, LINE_DARK = RGBColor(0xFF, 0xFF, 0xFF), RGBColor(0xDC, 0xD8, 0xCE), RGBColor(0x2A, 0x3A, 0x5C)
PAPER_DIM = RGBColor(0xB8, 0xBC, 0xC8)

HEAD, BODY = "Fraunces", "Manrope"                  # both available in Canva's font library

# type scale — "slightly bigger" than the HTML/PDF banner
T_TITLE, T_H2, T_H3, T_BODY, T_SMALL, T_KICK, T_STAT, T_LEAD = 118, 54, 29, 24, 16, 18, 56, 26

TOKEN = re.compile(r"(\*\*.+?\*\*|__.+?__)")


# ------------------------------------------------------------------ helpers
def mm(v):
    return Emu(int(round(v * S * 36000)))


def pt(v):
    return Pt(v * S)


def add_runs(p, text, size, color, font=BODY, bold=False, italic=False):
    for part in TOKEN.split(text):
        if not part:
            continue
        b, i, f = bold, italic, font
        if part.startswith("**"):
            part, b = part[2:-2], True
        elif part.startswith("__"):
            part, i, f = part[2:-2], True, HEAD
        r = p.add_run()
        r.text = part
        r.font.name, r.font.size, r.font.bold, r.font.italic = f, pt(size), b, i
        r.font.color.rgb = color


def text(slide, x, y, w, h, paras, *, size=T_BODY, color=INK, font=BODY, bold=False, italic=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, ls=1.15, after=3):
    """paras: list of str | dict(text, size, color, bold, font, after)."""
    box = slide.shapes.add_textbox(mm(x), mm(y), mm(w), mm(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    for i, item in enumerate(paras):
        spec = item if isinstance(item, dict) else {"text": item}
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = spec.get("align", align)
        p.line_spacing = spec.get("ls", ls)
        p.space_after = pt(spec.get("after", after))
        add_runs(p, spec["text"], spec.get("size", size), spec.get("color", color),
                 spec.get("font", font), spec.get("bold", bold), spec.get("italic", italic))
    return box


def bullets(slide, x, y, w, h, items, *, size=T_BODY, color=INK_SOFT, dot=TEAL, gap=7, ls=1.15, strong=INK):
    box = slide.shapes.add_textbox(mm(x), mm(y), mm(w), mm(h))
    tf = box.text_frame
    tf.word_wrap = True
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = ls
        p.space_after = pt(gap)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(mm(11))))
        pPr.set("indent", str(-int(mm(11))))
        r = p.add_run()
        r.text = "●\u2002"
        r.font.name, r.font.size = BODY, pt(size * 0.7)
        r.font.color.rgb = dot
        add_runs(p, item, size, color)
    for p in tf.paragraphs:                       # bold fragments in ink colour for contrast
        for r in p.runs:
            if r.font.bold and r.font.color.rgb == color:
                r.font.color.rgb = strong
    return box


def rect(slide, x, y, w, h, fill, *, line=None, rounded=True, radius=0.05):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                                 mm(x), mm(y), mm(w), mm(h))
    if rounded:
        shp.adjustments[0] = radius
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = pt(1.2)
    shp.shadow.inherit = False
    return shp


def oval(slide, x, y, w, h, fill, *, line=None, lw=2.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, mm(x), mm(y), mm(w), mm(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = pt(lw)
    shp.shadow.inherit = False
    return shp


def picture(slide, path, x, y, w=None, h=None):
    """Place an image by width OR height (mm), keeping aspect ratio. Returns (x, y, w, h) in mm."""
    with Image.open(path) as im:
        iw, ih = im.size
    if w is None:
        w = h * iw / ih
    elif h is None:
        h = w * ih / iw
    slide.shapes.add_picture(str(path), mm(x), mm(y), mm(w), mm(h))
    return x, y, w, h


def caption(slide, x, y, w, label, body, *, dark=False):
    box = text(slide, x, y, w, 14, [""], size=T_SMALL)
    p = box.text_frame.paragraphs[0]
    add_runs(p, label.upper() + "   ", T_SMALL * 0.85, PAPER if dark else INK, bold=True)
    add_runs(p, body, T_SMALL, PAPER_DIM if dark else INK_SOFT)
    return box


def kicker(slide, x, y, w, label, color=TEAL):
    rect(slide, x, y + 3.2, 18, 1.4, color, rounded=False)
    text(slide, x + 24, y, w, 10, [label.upper()], size=T_KICK, color=color, bold=True)


def section_chrome(slide, y, number, *, dark=False, color=TEAL):
    """Numbered node on the left-hand 'axon' spine."""
    oval(slide, 7, y + 22, 24, 24, NAVY if dark else PAPER, line=color, lw=9)
    text(slide, 7, y + 22, 24, 24, [str(number)], size=22, color=PAPER if dark else INK, font=HEAD,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def stat(slide, x, y, w, h, value, label, *, vcolor=INK, vsize=T_STAT, dark=False):
    rect(slide, x, y, w, h, NAVY_2 if dark else WHITE, line=None if dark else LINE)
    vh = vsize * 0.353 * 1.2  # value line height in mm (1 pt = 0.353 mm)
    text(slide, x + 10, y + 8, w - 20, vh, [value], size=vsize, color=vcolor, font=HEAD, ls=1.0)
    text(slide, x + 10, y + 8 + vh + 2, w - 20, h - vh - 18, [label], size=T_SMALL + 1,
         color=PAPER_DIM if dark else INK_SOFT, ls=1.2)


# ------------------------------------------------------------------ build
def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = mm(W_MM), mm(H_MM)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = PAPER
    X, CW = 38, W_MM - 2 * 38                        # safe margin, content width

    # vertical "axon" spine (three solid segments echo the teal→gold→coral gradient)
    rect(s, 17.5, 200, 3, 600, TEAL, rounded=False)
    rect(s, 17.5, 800, 3, 600, GOLD, rounded=False)
    rect(s, 17.5, 1400, 3, 530, CORAL, rounded=False)

    # ---------------- HEADER 0–200
    rect(s, 0, 0, W_MM, 200, NAVY, rounded=False)
    oval(s, 640, -230, 400, 400, NAVY_GLOW)          # decorative glows stay inside the header band
    oval(s, -150, -30, 220, 220, RGBColor(0x2A, 0x22, 0x3A))
    rect(s, 0, 162, W_MM, 1.2, LINE_DARK, rounded=False)
    oval(s, 380, 152, 220, 34, RGBColor(0x1B, 0x2E, 0x4C))
    rect(s, 0, 168, W_MM, 4, RGBColor(0x2B, 0x4A, 0x62), rounded=False)
    kicker(s, X, 22, 600, "Rare neurogenetic disease · OMIM #256850")
    text(s, X, 30, 620, 100, [{"text": "Giant Axonal", "after": 0}, {"text": "__Neuropathy__", "after": 0}],
         size=T_TITLE, color=PAPER, font=HEAD, ls=0.98)
    text(s, X, 132, 560, 28, ["How the loss of one small protein, gigaxonin, makes nerve fibres swell and die — "
                              "and how one family's fight produced the world's first intrathecal gene therapy."],
         size=T_LEAD, color=PAPER_DIM, ls=1.2)
    tx = X
    for label, col, border in [("Gene: GAN", TEAL_LIGHT, TEAL), ("Chromosome 16q23.2", PAPER_DIM, LINE_DARK),
                               ("Autosomal recessive", PAPER_DIM, LINE_DARK), ("Onset ≈ 3 years", PAPER_DIM, LINE_DARK),
                               ("Prevalence < 1 in 1,000,000", PAPER_DIM, LINE_DARK)]:
        w = 4.2 * len(label) + 16
        rect(s, tx, 180, w, 13, NAVY, line=border, radius=0.5)
        text(s, tx, 180, w, 13, [label.upper()], size=T_SMALL, color=col, bold=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tx += w + 6

    # ---------------- 1 · GENETIC ROOT CAUSE 200–500
    y = 200
    section_chrome(s, y, 1)
    kicker(s, X, y + 18, 600, "The genetic root cause")
    text(s, X, y + 30, CW, 24, ["One gene, two broken copies"], size=T_H2, color=INK, font=HEAD)
    rect(s, X, y + 60, 432, 86, TEAL_SOFT)
    ry = y + 68
    for k, v in [("Gene", "**GAN** — chromosome 16, band 16q23.2"),
                 ("Protein", "**Gigaxonin** — 597 aa, BTB/Kelch family"),
                 ("Function", "Adaptor for a **Cul3 E3 ubiquitin ligase**"),
                 ("Mechanism", "Loss of function — **>100** variants")]:
        text(s, X + 12, ry + 1.5, 70, 12, [k.upper()], size=T_SMALL, color=TEAL, bold=True)
        text(s, X + 84, ry, 336, 14, [v], size=T_BODY, color=INK)
        ry += 18
    bullets(s, X, y + 156, 432, 56, [
        "**Autosomal recessive:** disease appears only when both inherited copies of __GAN__ are defective.",
        "**Carrier parents are healthy** — one working copy is enough. Each pregnancy carries a 25% risk of an affected child.",
    ])
    _, _, lw_, lh_ = picture(s, HIRES / "locus.png", X, y + 214, h=66)
    caption(s, X, y + 284, 432, "Mutation locus", "Long arm of chromosome 16, band q23.2 — home of the GAN gene")
    ix, iy, iw, ih = picture(s, HIRES / "inheritance.png", 486, y + 58, h=226)
    caption(s, ix, iy + ih + 5, 330, "Inheritance pattern", "Two carriers → 1 in 4 affected, 2 in 4 carriers, 1 in 4 unaffected")

    # ---------------- 2 · MECHANISM 500–785
    y = 500
    rect(s, 0, y, W_MM, 285, PAPER_DEEP, rounded=False)
    rect(s, 17.5, y, 3, 285, TEAL, rounded=False)
    section_chrome(s, y, 2)
    kicker(s, X, y + 18, 700, "What gigaxonin does — and what fails without it")
    text(s, X, y + 30, CW, 24, ["The cell's filament recycler goes missing"], size=T_H2, color=INK, font=HEAD)
    px, py, pw, ph = picture(s, HIRES / "pathway.png", X, y + 60, h=192)
    caption(s, px, py + ph + 5, pw + 20, "Affected pathway", "With gigaxonin → Shh ON · Without gigaxonin → Shh OFF")
    cx = X + pw + 22
    cw = W_MM - 38 - cx
    bullets(s, cx, y + 60, cw, 110, [
        "**Normal role:** gigaxonin tags **intermediate filament (IF)** proteins — neurofilaments, vimentin, peripherin, GFAP, keratins — for degradation by the proteasome.",
        "**Without gigaxonin:** filaments are never cleared and pile up into dense aggregates that choke axonal transport.",
        "**Second hit:** gigaxonin also degrades the receptor **Patched**. When Patched accumulates it keeps **Smoothened** off, so **Sonic hedgehog signalling stays OFF**.",
    ])
    fx, fy, fw, fh = picture(s, HIRES / "filament.png", cx, y + 172, h=100)
    text(s, cx + fw + 12, y + 196, cw - fw - 12, 50, [
        {"text": "INTERMEDIATE FILAMENT AGGREGATE", "size": T_SMALL * 0.85, "color": INK, "bold": True, "after": 2},
        {"text": "Patient cell — the hallmark of gigaxonin loss (arrow). N = nucleus.", "size": T_SMALL, "color": INK_SOFT},
    ], ls=1.25)

    # ---------------- 3 · PHENOTYPE 785–1125
    y = 785
    section_chrome(s, y, 3, color=CORAL)
    kicker(s, X, y + 18, 600, "The phenotype, cell by cell", CORAL)
    text(s, X, y + 30, CW, 24, ["Axons that swell into __giants__"], size=T_H2, color=INK, font=HEAD)
    bullets(s, X, y + 60, 470, 130, [
        "**Focal axonal swellings** several times the normal diameter, packed with tangled neurofilaments — the \"giant axons\" that name the disease.",
        "**Myelin thins and degenerates** over each swelling; conduction slows, then fails.",
        "**Both nervous systems are hit:** peripheral nerves (PNS) and the cerebellum, brainstem and white-matter tracts (CNS).",
        "**Not only neurons:** GFAP aggregates in astrocytes (Rosenthal fibres); disordered keratin gives the characteristic **tightly curled \"kinky\" hair**.",
    ], dot=CORAL)
    sw = (470 - 2 * 10) / 3
    for i, (v, l, c, vs) in enumerate([("3–5×", "typical diameter of a swollen \"giant\" axon vs. a healthy one", CORAL, T_STAT),
                                        ("PNS + CNS", "a length-dependent axonopathy that becomes a whole-brain disease", INK, 40),
                                        ("MRI", "cerebellar and brainstem white-matter change supports diagnosis", TEAL, T_STAT)]):
        stat(s, X + i * (sw + 10), y + 232, sw, 80, v, l, vcolor=c, vsize=vs)
    nx, ny, nw, nh = picture(s, HIRES / "neuron.png", 546, y + 58, h=250)
    caption(s, nx, ny + nh + 5, 266, "Normal vs GAN", "Neurofilament accumulation → axonal swelling → myelin loss, in CNS and PNS")

    # ---------------- 4 · SYMPTOMS 1125–1290
    y = 1125
    rect(s, 0, y, W_MM, 165, PAPER_DEEP, rounded=False)
    rect(s, 17.5, y, 3, 165, GOLD, rounded=False)
    section_chrome(s, y, 4, color=GOLD_DARK)
    kicker(s, X, y + 18, 600, "Symptoms and clinical course", CORAL)
    text(s, X, y + 30, CW, 24, ["A slow, relentless loss — from the feet upward"], size=T_H2, color=INK, font=HEAD)
    rect(s, X, y + 66, CW, 1.6, LINE, rounded=False)
    colw = (CW - 3 * 12) / 4
    stages = [("Onset · ≈ 3 years", "The clumsy walker", "Unsteady gait, frequent falls, weak feet and hands, high-arched feet, diminished reflexes. Tightly curled hair unlike either parent is a visible clue.", TEAL),
              ("Childhood · 5–10 years", "Peripheral neuropathy", "Loss of vibration and position sense, absent reflexes, muscle wasting below knees and elbows. Nerve biopsy shows giant axons; genetic testing confirms __GAN__ variants.", GOLD_DARK),
              ("Adolescence", "Central nervous system", "Wheelchair dependence, cerebellar ataxia, nystagmus, slurred speech, scoliosis, seizures and cognitive decline as brain white matter degenerates.", CORAL),
              ("Early adulthood", "Late stage", "Swallowing and breathing muscles weaken; most patients die of respiratory failure in the 2nd–3rd decade. Until 2015 there was **no treatment** beyond supportive care.", INK_SOFT)]
    for i, (when, title, body, col) in enumerate(stages):
        cx = X + i * (colw + 12)
        oval(s, cx, y + 61, 12, 12, PAPER_DEEP, line=col, lw=7)
        text(s, cx, y + 80, colw, 9, [when.upper()], size=T_SMALL, color=col, bold=True)
        text(s, cx, y + 90, colw, 14, [title], size=T_H3, color=INK, bold=True)
        text(s, cx, y + 106, colw, 56, [body], size=T_BODY - 3, color=INK_SOFT, ls=1.2)

    # ---------------- 5 · HANNAH 1290–1500 (dark)
    y = 1290
    rect(s, 0, y, W_MM, 210, NAVY, rounded=False)
    oval(s, 720, y + 40, 160, 160, RGBColor(0x1E, 0x2A, 0x3E))
    rect(s, 17.5, y, 3, 210, GOLD, rounded=False)
    section_chrome(s, y, 5, dark=True, color=GOLD)
    hx, hy, hw, hh = picture(s, HIRES / "hannah.png", X, y + 20, w=270)
    rect(s, hx + 8, hy + hh - 30, 200, 22, NAVY, radius=0.15)
    text(s, hx + 14, hy + hh - 27, 190, 18, [
        {"text": "HANNAH SAMES", "size": T_SMALL * 0.85, "color": TEAL_LIGHT, "bold": True, "after": 1},
        {"text": "Rexford, New York · diagnosed 24 March 2008, aged 4", "size": T_SMALL * 0.9, "color": PAPER},
    ])
    cx = X + 270 + 22
    cw = W_MM - 38 - cx
    kicker(s, cx, y + 18, 500, "The case that changed the field", GOLD)
    text(s, cx, y + 30, cw, 24, ["\"We decided to __fight__.\""], size=T_H2, color=PAPER, font=HEAD)
    bullets(s, cx, y + 60, cw, 96, [
        "**2007:** at age 3 Hannah's walking lags behind; a year of tests ends with a GAN diagnosis in **March 2008**.",
        "**The hardest genotype:** two deletions — Hannah makes **no gigaxonin at all** (CRIM-negative), so her immune system had never \"seen\" the protein.",
        "**Hannah's Hope Fund**, founded within months by Lori and Matt Sames, convenes the first GAN symposium (2008), raises **>$8 million** and funds Steven Gray (UNC) to build the vector.",
    ], size=T_BODY - 2, color=PAPER_DIM, strong=PAPER)
    rect(s, cx, y + 160, 2.5, 40, TEAL, rounded=False)
    text(s, cx + 10, y + 158, cw - 10, 44, [
        {"text": "Hannah's parents found themselves in the unimaginable situation of having funded a clinical trial that might have to exclude their own child.",
         "size": 25, "color": PAPER, "font": HEAD, "after": 3},
        {"text": "DNA SCIENCE, JULY 2016", "size": T_SMALL * 0.85, "color": PAPER_DIM, "bold": True},
    ], ls=1.2)

    # ---------------- 6 · GENE THERAPY 1500–1740
    y = 1500
    section_chrome(s, y, 6)
    kicker(s, X, y + 18, 760, "The genetic treatment · scAAV9/JeT-GAN · first-in-human intrathecal AAV trial, NIH Clinical Center 2015 (NCT02362438)")
    text(s, X, y + 30, CW, 24, ["A working gene, delivered into the spinal fluid"], size=T_H2, color=INK, font=HEAD)
    steps = [("1", "Build the vector", "**scAAV9**: a self-complementary AAV9 shell that enters neurons and spreads through spinal cord and brain. No viral genes; does not integrate."),
             ("2", "Load the cargo", "A **codon-optimised human** __GAN__ gene driven by **JeT**, a tiny synthetic promoter that fits the compact genome."),
             ("3", "Single lumbar injection", "One dose — up to **3.5 × 10¹⁴** vector genomes — into the CSF, bathing motor neurons and dorsal root ganglia and bypassing the blood–brain barrier."),
             ("4", "Restore gigaxonin", "Transduced neurons make gigaxonin within ~2 weeks, clearing aggregates and letting damaged sensory nerves **regenerate**.")]
    cwid = (CW - 3 * 12) / 4
    for i, (n, title, body) in enumerate(steps):
        cx = X + i * (cwid + 12)
        rect(s, cx, y + 62, cwid, 104, WHITE, line=LINE)
        text(s, cx + 12, y + 70, 30, 20, [n], size=44, color=TEAL, font=HEAD, ls=1.0)
        text(s, cx + 12, y + 90, cwid - 24, 14, [title], size=T_H3 - 4, color=INK, bold=True)
        text(s, cx + 12, y + 105, cwid - 24, 58, [body], size=T_BODY - 5, color=INK_SOFT, ls=1.2)
        if i < 3:
            text(s, cx + cwid - 2, y + 84, 16, 14, ["→"], size=24, color=TEAL, bold=True, align=PP_ALIGN.CENTER)
    sw3 = (CW - 2 * 12) / 3
    for i, (v, l, c) in enumerate([("May 2015", "first human ever to receive a therapeutic gene via the spinal fluid — a GAN patient", TEAL),
                                    ("21 July 2016", "Hannah, aged 12, treated as the 5th child (≈120 trillion vector genomes)", GOLD_DARK),
                                    ("CRIM-negative", "making no gigaxonin, Hannah needed a dedicated immunosuppression protocol", CORAL)]):
        stat(s, X + i * (sw3 + 12), y + 176, sw3, 56, v, l, vcolor=c, vsize=40)

    # ---------------- 7 · RESULTS 1740–1930 (dark)
    y = 1740
    rect(s, 0, y, W_MM, 260, NAVY, rounded=False)
    oval(s, -140, y + 10, 250, 250, NAVY_GLOW)
    rect(s, 17.5, y, 3, 190, CORAL, rounded=False)
    section_chrome(s, y, 7, dark=True)
    kicker(s, X, y + 18, 600, "What the trial showed — NEJM, March 2024")
    text(s, X, y + 30, CW, 24, ["Bending the curve of an \"untreatable\" disease"], size=T_H2, color=PAPER, font=HEAD)
    stats = [("14", "children aged 6–14 treated at four dose levels, followed for a median of 68.7 months"),
             ("1", "serious adverse event (fever) possibly related; CSF inflammation managed with steroids"),
             ("−7.2 → −1.9", "MFM-32 points lost per year before vs. after the 1.8 × 10¹⁴ vg dose"),
             ("99%", "posterior probability that motor decline slowed at that dose — above the 95% bar"),
             ("6 / 14", "regained, stabilised or newly recordable sensory nerve responses after dosing"),
             ("3–48×", "more regenerating nerve-fibre clusters in sensory-nerve biopsies at one year")]
    gw, gh = (440 - 2 * 10) / 3, 58
    for i, (v, l) in enumerate(stats):
        col, row = i % 3, i // 3
        stat(s, X + col * (gw + 10), y + 60 + row * (gh + 8), gw, gh, v, l, vcolor=PAPER,
             vsize=34 if len(v) > 6 else 44, dark=True)
    cx = X + 440 + 22
    cw = W_MM - 38 - cx
    bullets(s, cx, y + 60, cw, 78, [
        "**Not a cure — yet.** Existing damage is not reversed; benefit is largest when children are treated **young and early**.",
        "**A template for the field:** intrathecal AAV9 delivery pioneered here now underpins trials for many other childhood brain diseases.",
        "**Next:** programme licensed to industry (Taysha, TSHA-120); Hannah's Hope Fund funds a second injection site.",
    ], size=T_BODY - 3, color=PAPER_DIM, strong=PAPER, gap=5)
    rect(s, cx, y + 146, 2.5, 36, TEAL, rounded=False)
    text(s, cx + 10, y + 144, cw - 10, 40, [
        {"text": "One family with no medical background turned a diagnosis into a first-in-human therapy — proof of what patient-driven science can do.",
         "size": 23, "color": PAPER, "font": HEAD, "after": 3},
        {"text": "TAKE-HOME MESSAGE", "size": T_SMALL * 0.85, "color": PAPER_DIM, "bold": True},
    ], ls=1.2)

    # ---------------- FOOTER 1930–2000 (low-visibility zone at the base)
    y = 1930
    rect(s, 0, y, W_MM, 0.8, LINE_DARK, rounded=False)
    text(s, X, y + 14, 560, 44, [
        "Sources: Bharucha-Goebel DX et al. __Intrathecal Gene Therapy for Giant Axonal Neuropathy.__ N Engl J Med 2024;390:1092-1104 · "
        "ClinicalTrials.gov NCT02362438 · NIH News, March 2024 · Hannah's Hope Fund (hannahshopefund.org) · The Children's Inn at NIH, "
        "\"Hannah's Story\" · DNA Science (PLOS), July 2016. Figures adapted from the source poster."],
        size=T_SMALL * 0.9, color=PAPER_DIM, ls=1.3)
    text(s, 600, y + 16, 212, 40, [
        {"text": "Giant Axonal Neuropathy", "size": 30, "color": PAPER, "font": HEAD, "align": PP_ALIGN.RIGHT, "after": 2},
        {"text": "FROM GENE TO GENE THERAPY", "size": T_SMALL * 0.85, "color": TEAL_LIGHT, "bold": True, "align": PP_ALIGN.RIGHT},
    ])

    prs.save(OUT)
    print(f"wrote {OUT.name}  ({OUT.stat().st_size // 1024} KB)  slide {W_MM * S:.0f} x {H_MM * S:.0f} mm")


if __name__ == "__main__":
    build()
