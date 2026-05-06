"""
Master's thesis defense slides — Option A (18 slides, ~20 min)
Hyunho Kook -- Stabilizing Direct Training of SNNs

Design follows reference deck MP-Init_TrSG_Defense.pptx:
  - Navy / teal / amber / coral palette
  - Serif title (Cambria), sans body (Calibri)
  - Decorative circles on cover/conclusion
  - Section label · slide title · content
  - Audience: advisor + ML/CS people who don't know SNN
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy
import os

ASSETS = "ppt_assets"

# 16:9 widescreen
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ─── Palette (from reference) ──────────────────────────────────────
NAVY     = RGBColor(0x0F, 0x25, 0x40)
NAVY_DK  = RGBColor(0x09, 0x1A, 0x30)
TEAL     = RGBColor(0x1D, 0x7E, 0x7E)
TEAL_LT  = RGBColor(0xE5, 0xF5, 0xF0)
AMBER    = RGBColor(0xE0, 0xB8, 0x5F)
AMBER_LT = RGBColor(0xFA, 0xEE, 0xCB)
CORAL    = RGBColor(0xE5, 0x69, 0x45)
CORAL_LT = RGBColor(0xFA, 0xEE, 0xE0)
INK      = RGBColor(0x1A, 0x1A, 0x1A)
GRAY     = RGBColor(0x7A, 0x7A, 0x7A)
GRAY_LT  = RGBColor(0xE5, 0xE5, 0xE0)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
CREAM    = RGBColor(0xF8, 0xF7, 0xF2)

SERIF = "Cambria"
SANS  = "Calibri"

# ─── Helpers ───────────────────────────────────────────────────────
def add_blank():
    return prs.slides.add_slide(prs.slide_layouts[6])

def fill_bg(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    bg.shadow.inherit = False
    # send to back
    spTree = slide.shapes._spTree
    spTree.remove(bg._element); spTree.insert(2, bg._element)
    return bg

def add_circle(slide, x, y, d, color):
    """Decorative circle for title/cover slides."""
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = color
    c.line.fill.background()
    return c

def text(slide, x, y, w, h, txt, *, font=SANS, size=18, bold=False, italic=False,
         color=INK, align=None, anchor=None, spacing=None):
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tx.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    p = tf.paragraphs[0]
    p.text = txt
    p.font.name = font
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    if align: p.alignment = align
    if anchor: tf.vertical_anchor = anchor
    if spacing:
        for r in p.runs:
            rPr = r._r.get_or_add_rPr()
            rPr.set('spc', str(spacing))
    return tx

def section_label(slide, txt):
    """Top section label, ALL CAPS letter-spaced teal."""
    text(slide, 0.5, 0.4, 12, 0.3, txt.upper(),
         font=SANS, size=11, bold=True, color=TEAL, spacing=300)

def slide_title(slide, txt, sub=None):
    text(slide, 0.5, 0.85, 12.3, 0.85, txt,
         font=SERIF, size=30, bold=True, color=NAVY)
    if sub:
        text(slide, 0.5, 1.75, 12.3, 0.4, sub,
             font=SANS, size=14, italic=True, color=GRAY)

def page_footer(slide, n, total=18):
    text(slide, 0.5, 7.10, 6, 0.3, "Defense — Hyunho Kook",
         font=SANS, size=10, color=GRAY)
    text(slide, 7.0, 7.10, 5.8, 0.3, f"{n} / {total}",
         font=SANS, size=10, color=GRAY, align=PP_ALIGN.RIGHT)

def amber_underline(slide, x, y, w=0.7):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(0.04))
    bar.fill.solid(); bar.fill.fore_color.rgb = AMBER
    bar.line.fill.background()
    return bar

def callout_box(slide, x, y, w, h, *, fill=NAVY, accent=AMBER,
                title=None, body=None, title_size=14, body_size=12,
                title_color=AMBER, body_color=WHITE):
    # vertical accent line
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(0.06), Inches(h))
    bar.fill.solid(); bar.fill.fore_color.rgb = accent
    bar.line.fill.background()
    # background box
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(x + 0.06), Inches(y), Inches(w - 0.06), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = fill
    box.line.fill.background()
    # text
    if title or body:
        tx = slide.shapes.add_textbox(
            Inches(x + 0.35), Inches(y + 0.15),
            Inches(w - 0.5), Inches(h - 0.3))
        tf = tx.text_frame; tf.word_wrap = True
        if title:
            p = tf.paragraphs[0]
            p.text = title
            p.font.name = SANS; p.font.size = Pt(title_size); p.font.bold = True
            p.font.color.rgb = title_color
        if body:
            p2 = tf.add_paragraph() if title else tf.paragraphs[0]
            p2.text = body
            p2.font.name = SANS; p2.font.size = Pt(body_size)
            p2.font.color.rgb = body_color
            p2.space_before = Pt(6) if title else Pt(0)

def card(slide, x, y, w, h, *, fill=WHITE, border=TEAL, border_w=1.5):
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    rect.fill.solid(); rect.fill.fore_color.rgb = fill
    rect.line.color.rgb = border; rect.line.width = Pt(border_w)
    return rect

def bullets(slide, x, y, w, h, items, *, size=14, color=INK,
            bullet="•", spacing=6):
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tx.text_frame; tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"{bullet}  {item}"
        p.font.name = SANS; p.font.size = Pt(size); p.font.color.rgb = color
        if i > 0: p.space_before = Pt(spacing)


def img(slide, path, x, y, w=None, h=None):
    if not os.path.exists(path):
        return None
    kw = {}
    if w is not None: kw["width"] = Inches(w)
    if h is not None: kw["height"] = Inches(h)
    return slide.shapes.add_picture(path, Inches(x), Inches(y), **kw)


# ════════════════════════════════════════════════════════════════════
# S1 — Title
# ════════════════════════════════════════════════════════════════════
s = add_blank()
fill_bg(s, NAVY)
# Decorative circles
add_circle(s, -2.0, -2.0, 5.5, NAVY_DK)
add_circle(s, 10.5, 4.5, 4.0, NAVY_DK)

text(s, 0.7, 0.9, 10, 0.4, "MASTER'S THESIS DEFENSE",
     font=SANS, size=12, bold=True, color=AMBER, spacing=400)
text(s, 0.7, 1.7, 12, 1.3,
     "Enhancing Stability in Direct Training",
     font=SERIF, size=44, bold=True, color=WHITE)
text(s, 0.7, 2.55, 12, 1.3,
     "of Spiking Neural Networks",
     font=SERIF, size=44, bold=True, color=WHITE)
text(s, 0.7, 3.7, 12, 0.5,
     "Membrane Potential Initialization & Threshold-Robust Surrogate Gradients",
     font=SERIF, size=16, italic=True, color=AMBER)

# Amber underline
amber_underline(s, 0.7, 4.45, w=1.0)

text(s, 0.7, 4.7, 9, 0.4, "Hyunho Kook  (국 현 호)",
     font=SANS, size=18, bold=True, color=WHITE)
text(s, 0.7, 5.2, 9, 0.35, "Advisor: Prof. Eunhyeok Park",
     font=SANS, size=12, color=RGBColor(0xCB, 0xD8, 0xE0))
text(s, 0.7, 5.55, 9, 0.35, "Committee: Prof. Jungseul Ok · Prof. Jae-Joon Kim",
     font=SANS, size=12, color=RGBColor(0xCB, 0xD8, 0xE0))

text(s, 0.7, 7.05, 12, 0.3,
     "Department of Computer Science and Engineering · POSTECH · January 2026",
     font=SANS, size=10, color=RGBColor(0x80, 0x95, 0xA8))


# ════════════════════════════════════════════════════════════════════
# S2 — Why SNNs?
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Motivation")
slide_title(s, "Why Spiking Neural Networks?")

# Left column: 3 rows of bullet points
y = 2.4
text(s, 0.5, y, 7.2, 0.45,
     "DNNs are powerful but expensive",
     font=SANS, size=15, bold=True, color=NAVY)
text(s, 0.5, y+0.5, 7.2, 1.0,
     "Training and serving large models routinely consumes megawatt-scale power; "
     "even edge inference is hard under tight energy budgets.",
     font=SANS, size=12, color=INK)

text(s, 0.5, y+1.6, 7.2, 0.45,
     "SNNs: an event-driven alternative",
     font=SANS, size=15, bold=True, color=NAVY)
text(s, 0.5, y+2.1, 7.2, 1.0,
     "Information flows as discrete spikes over time. Paired with neuromorphic hardware, "
     "computation skips inactive neurons — orders of magnitude lower energy.",
     font=SANS, size=12, color=INK)

text(s, 0.5, y+3.2, 7.2, 0.45,
     "But direct training is unstable",
     font=SANS, size=15, bold=True, color=CORAL)
text(s, 0.5, y+3.7, 7.2, 1.0,
     "Recent direct-training methods narrowed the accuracy gap, yet two persistent instabilities "
     "still hold SNNs back.",
     font=SANS, size=12, color=INK)

# Right column: dark callout panel
panel_x, panel_y, panel_w, panel_h = 8.5, 2.3, 4.4, 4.4
box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
    Inches(panel_x), Inches(panel_y), Inches(panel_w), Inches(panel_h))
box.fill.solid(); box.fill.fore_color.rgb = NAVY; box.line.fill.background()

text(s, panel_x + 0.3, panel_y + 0.25, panel_w - 0.5, 0.3,
     "THIS THESIS", font=SANS, size=10, bold=True, color=AMBER, spacing=400)
text(s, panel_x + 0.3, panel_y + 0.7, panel_w - 0.5, 0.5,
     "Two instabilities,",
     font=SERIF, size=22, bold=True, color=WHITE)
text(s, panel_x + 0.3, panel_y + 1.15, panel_w - 0.5, 0.5,
     "two principled fixes.",
     font=SERIF, size=22, bold=True, color=WHITE)

# Two numbered items inside panel
nx, ny1, ny2 = panel_x + 0.3, panel_y + 2.0, panel_y + 3.2
# accent bar 1
ab1 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
    Inches(nx), Inches(ny1), Inches(0.04), Inches(0.85))
ab1.fill.solid(); ab1.fill.fore_color.rgb = TEAL; ab1.line.fill.background()
text(s, nx + 0.15, ny1, panel_w - 0.6, 0.3,
     "①  Temporal Covariate Shift",
     font=SANS, size=13, bold=True, color=WHITE)
text(s, nx + 0.15, ny1 + 0.35, panel_w - 0.6, 0.5,
     "in the membrane potential — invisible to existing remedies.",
     font=SANS, size=10, color=RGBColor(0xCB, 0xD8, 0xE0))
# accent bar 2
ab2 = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
    Inches(nx), Inches(ny2), Inches(0.04), Inches(0.95))
ab2.fill.solid(); ab2.fill.fore_color.rgb = AMBER; ab2.line.fill.background()
text(s, nx + 0.15, ny2, panel_w - 0.6, 0.3,
     "②  Gradient pathology",
     font=SANS, size=13, bold=True, color=WHITE)
text(s, nx + 0.15, ny2 + 0.35, panel_w - 0.6, 0.6,
     "when the firing threshold becomes a learnable parameter — never explicitly characterized before.",
     font=SANS, size=10, color=RGBColor(0xCB, 0xD8, 0xE0))

page_footer(s, 2)


# ════════════════════════════════════════════════════════════════════
# S3 — SNNs in 60 seconds (LIF + surrogate gradient)
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Background · 1/2")
slide_title(s, "Spiking Neurons in 60 Seconds",
            "How information flows, and how we get gradients through it")

# Left: LIF
text(s, 0.5, 2.5, 6.2, 0.4,
     "Leaky Integrate-and-Fire neuron",
     font=SANS, size=14, bold=True, color=NAVY)

eq_y = 3.0
def equation_box(slide, x, y, w, h, label, eq):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = CREAM
    box.line.color.rgb = GRAY_LT; box.line.width = Pt(0.5)
    text(slide, x + 0.15, y + 0.05, 1.4, 0.3,
         label, font=SANS, size=10, bold=True, color=TEAL)
    text(slide, x + 1.5, y + 0.05, w - 1.6, h - 0.1, eq,
         font="Cambria Math", size=12, color=INK,
         anchor=MSO_ANCHOR.MIDDLE)

equation_box(s, 0.5, eq_y,        6.2, 0.6,
             "Integrate", "M[t] = (1−1/τ)·U[t−1]  +  (1/τ)·I_in[t]")
equation_box(s, 0.5, eq_y + 0.75, 6.2, 0.6,
             "Fire",      "S[t] = H( M[t] − V_thr )       (1 if exceeds, else 0)")
equation_box(s, 0.5, eq_y + 1.5,  6.2, 0.6,
             "Reset",     "U[t] = M[t] − V_thr · S[t]   (soft reset)")

text(s, 0.5, eq_y + 2.2, 6.2, 0.5,
     "Two key learnable parameters:  τ  (leakage)  and  V_thr  (threshold).",
     font=SANS, size=11, italic=True, color=GRAY)

# Right: Surrogate gradient
text(s, 7.2, 2.5, 5.6, 0.4,
     "Heaviside is non-differentiable",
     font=SANS, size=14, bold=True, color=NAVY)
text(s, 7.2, 2.95, 5.6, 0.5,
     "S = H(M − V_thr)   ⇒   H'(x) = δ(x)   ⇒   no gradient flow",
     font="Cambria Math", size=12, color=INK)

text(s, 7.2, 3.7, 5.6, 0.4,
     "Surrogate gradient (SG)",
     font=SANS, size=14, bold=True, color=NAVY)
text(s, 7.2, 4.15, 5.6, 1.5,
     "Replace H with a smooth surrogate f near the threshold.\n"
     "Use f'(x) ≈ H'(x) only in backward pass — forward pass\n"
     "still emits binary spikes.",
     font=SANS, size=12, color=INK)

# Bottom callout
callout_box(s, 0.5, 6.05, 12.3, 0.85,
            title="Everything in this thesis turns on M[t] and V_thr.",
            body="The membrane potential drives TCS;  the threshold drives the gradient pathology.",
            title_size=13, body_size=11)

page_footer(s, 3)


# ════════════════════════════════════════════════════════════════════
# S4 — Diagnosis ① TCS in membrane potential
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Background · 2/2")
slide_title(s, "Diagnosis ①: TCS Lives in the Membrane Potential")

# Left: prose
text(s, 0.5, 2.5, 6.5, 0.4,
     "What everyone notices",
     font=SANS, size=14, bold=True, color=NAVY)
text(s, 0.5, 2.95, 6.5, 1.5,
     "Activation distributions drift across timesteps. Batch-norm statistics estimated\n"
     "at one timestep no longer fit at another — accuracy drops.",
     font=SANS, size=12, color=INK)

text(s, 0.5, 4.2, 6.5, 0.4,
     "What existing fixes do",
     font=SANS, size=14, bold=True, color=NAVY)
text(s, 0.5, 4.65, 6.5, 1.0,
     "TEBN, TAB add per-timestep BN parameters.\n"
     "More compute, more parameters — and they don't reach the cause.",
     font=SANS, size=12, color=INK)

# Bottom callout — the claim
callout_box(s, 0.5, 6.0, 6.5, 0.95,
            fill=NAVY, accent=CORAL,
            title="Our claim",
            body="TCS originates one level deeper — in M[t] itself.",
            title_color=CORAL, title_size=13, body_size=12)

# Right: figure
img(s, ASSETS + "/density_plot_membrane_potential_tdbn.png", 7.5, 2.4, w=5.4)
text(s, 7.5, 6.3, 5.4, 0.3,
     "M[t] under tdBN — distribution clearly drifts as t grows",
     font=SANS, size=10, italic=True, color=CORAL, align=PP_ALIGN.CENTER)

page_footer(s, 4)


# ════════════════════════════════════════════════════════════════════
# S5 — Diagnosis ② V_thr training breaks gradient
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Setup")
slide_title(s, "Diagnosis ②: Trainable V_thr Breaks the Gradient")

text(s, 0.5, 2.5, 12.3, 0.5,
     "V_thr controls when a neuron fires — we want it to learn.",
     font=SANS, size=15, bold=True, color=NAVY)
text(s, 0.5, 3.05, 12.3, 0.6,
     "Modern methods (PLIF, LTMD, DIET-SNN) all make V_thr trainable. But the surrogate gradient flowing\n"
     "through V_thr is highly sensitive to its scale.",
     font=SANS, size=12, color=INK)

# Two failure scenarios
fy = 4.1

# Left scenario
left_x, right_x, w, h = 0.5, 6.95, 5.85, 2.0
rect_l = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
    Inches(left_x), Inches(fy), Inches(w), Inches(h))
rect_l.fill.solid(); rect_l.fill.fore_color.rgb = CORAL_LT
rect_l.line.color.rgb = CORAL; rect_l.line.width = Pt(1.2)
text(s, left_x + 0.3, fy + 0.2, 1.3, 0.3,
     "V_thr ≪ 1", font=SANS, size=12, bold=True, color=CORAL)
text(s, left_x + 0.3, fy + 0.6, w - 0.5, 1.4,
     "AS-SG floods the layer (window covers most of the distribution).\n"
     "RS-SG explodes through the 1/V_thr factor in the chain rule.",
     font=SANS, size=12, color=INK)

# Right scenario
rect_r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
    Inches(right_x), Inches(fy), Inches(w), Inches(h))
rect_r.fill.solid(); rect_r.fill.fore_color.rgb = CORAL_LT
rect_r.line.color.rgb = CORAL; rect_r.line.width = Pt(1.2)
text(s, right_x + 0.3, fy + 0.2, 1.3, 0.3,
     "V_thr ≫ 1", font=SANS, size=12, bold=True, color=CORAL)
text(s, right_x + 0.3, fy + 0.6, w - 0.5, 1.4,
     "AS-SG starves (window misses most of the distribution).\n"
     "RS-SG vanishes through the same 1/V_thr factor.",
     font=SANS, size=12, color=INK)

text(s, 0.5, 6.4, 12.3, 0.5,
     "These are not corner cases — real trainings visit them.",
     font=SANS, size=13, italic=True, color=NAVY, align=PP_ALIGN.CENTER)

page_footer(s, 5)


# ════════════════════════════════════════════════════════════════════
# S6 — Theory: Markov chain
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Part I · MP-Init")
slide_title(s, "Why Does the Membrane Potential Drift?",
            "Because it's a Markov chain pulled toward a unique distribution")

# Recurrence
callout_box(s, 0.5, 2.4, 12.3, 0.9,
            fill=CREAM, accent=TEAL,
            title="The recurrence is a Markov chain",
            body="U[t+1]  =  f( U[t],  I_in[t+1] )         — i.i.d. inputs ⇒ Markov chain",
            title_color=TEAL, body_color=INK,
            title_size=14, body_size=14)

# Theorem
callout_box(s, 0.5, 3.5, 12.3, 1.6,
            fill=NAVY, accent=AMBER,
            title="Theorem  (Doeblin minorization)",
            body="Under (i.i.d. inputs) + (bounded U[t]),  the chain has a unique stationary "
                 "distribution π.  Moreover,\n\n"
                 "      ‖ Pⁿ(x, ·)  −  π(·) ‖_TV   ≤   C · (1 − ε)ⁿ        (geometric convergence)",
            title_color=AMBER, body_color=WHITE,
            title_size=14, body_size=12)

# Implications
text(s, 0.5, 5.4, 12.3, 0.5,
     "Implications",
     font=SANS, size=14, bold=True, color=CORAL)
bullets(s, 0.5, 5.85, 12.3, 1.5, [
    "Initial state ≠ π   ⇒   TCS is mathematically inevitable",
    "Standard practice  U[0] = 0  is almost always far from π",
    "Don't fix the normalization layer — fix the initial state",
], size=13, color=INK, bullet="▸")

page_footer(s, 6)


# ════════════════════════════════════════════════════════════════════
# S7 — MP-Init: Algorithm
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Part I · MP-Init")
slide_title(s, "MP-Init — Initialize at the Stationary Mean",
            "U[T] is the closest sample to π → use a running mean of E[U[T]]")

# Left: lemma + estimate
text(s, 0.5, 2.5, 6.5, 0.4,
     "Lemma — best constant is the mean",
     font=SANS, size=14, bold=True, color=NAVY)
callout_box(s, 0.5, 2.95, 6.5, 0.7,
            fill=CREAM, accent=TEAL,
            body="argmin_c   E[ (π − c)² ]   =   E[π]",
            body_color=INK, body_size=14, title_size=12)

text(s, 0.5, 3.95, 6.5, 0.4,
     "Practical estimator",
     font=SANS, size=14, bold=True, color=NAVY)
text(s, 0.5, 4.4, 6.5, 1.5,
     "U[T] is the closest sample to π. Maintain a per-layer running mean\n"
     "μ ← β · μ  +  (1 − β) · mean( U[T] )    over training, β = 0.9.\n"
     "At every batch, set U[0] = μ.",
     font=SANS, size=12, color=INK)

# Right: pseudocode
code_x = 7.4
text(s, code_x, 2.5, 5.4, 0.4,
     "Algorithm  (training-only)",
     font=SANS, size=14, bold=True, color=NAVY)

code_box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
    Inches(code_x), Inches(2.95), Inches(5.4), Inches(2.7))
code_box.fill.solid(); code_box.fill.fore_color.rgb = NAVY
code_box.line.fill.background()

code_lines = [
    "init  μ_l ← 0   for each layer l",
    "",
    "for each batch:",
    "    U_l[0] ← μ_l",
    "    run T LIF steps as usual",
    "    m ← mean( U_l[T] of active neurons )",
    "    μ_l ← β · μ_l  +  (1 − β) · m",
]
ctx = s.shapes.add_textbox(Inches(code_x + 0.25), Inches(3.05),
                            Inches(5.0), Inches(2.5))
tf = ctx.text_frame; tf.word_wrap = True
for i, line in enumerate(code_lines):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = line
    p.font.name = "Consolas"; p.font.size = Pt(11.5); p.font.color.rgb = WHITE

# Bottom strip
callout_box(s, 0.5, 6.05, 12.3, 0.9,
            fill=TEAL, accent=AMBER,
            title="Inference: μ frozen → standard LIF dynamics",
            body="Zero changes to BN, no per-timestep parameters, full neuromorphic-hardware compatibility.",
            title_color=WHITE, body_color=RGBColor(0xE0, 0xF0, 0xEA),
            title_size=13, body_size=11)

page_footer(s, 7)


# ════════════════════════════════════════════════════════════════════
# S8 — MP-Init: Result + Mechanism
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Part I · MP-Init")
slide_title(s, "MP-Init Works — at Both the Distribution and the Metric")

# Top row: 4 panels showing distribution alignment
text(s, 0.5, 2.4, 12.3, 0.4,
     "Membrane potential distributions across timesteps  (ResNet-19 / CIFAR-100)",
     font=SANS, size=11, italic=True, color=GRAY)
img(s, ASSETS + "/density_plot_membrane_potential_tdbn.png", 0.5, 2.8, w=2.95)
img(s, ASSETS + "/density_plot_membrane_potential_tebn.png", 3.6, 2.8, w=2.95)
img(s, ASSETS + "/density_plot_membrane_potential_tab.png",  6.7, 2.8, w=2.95)
img(s, ASSETS + "/density_plot_membrane_potential_mpinit.png", 9.8, 2.8, w=2.95)
text(s, 0.5, 4.65, 2.95, 0.3, "tdBN", font=SANS, size=10, color=GRAY, align=PP_ALIGN.CENTER)
text(s, 3.6, 4.65, 2.95, 0.3, "TEBN", font=SANS, size=10, color=GRAY, align=PP_ALIGN.CENTER)
text(s, 6.7, 4.65, 2.95, 0.3, "TAB",  font=SANS, size=10, color=GRAY, align=PP_ALIGN.CENTER)
text(s, 9.8, 4.65, 2.95, 0.3, "MP-Init (ours)",
     font=SANS, size=10, bold=True, color=TEAL, align=PP_ALIGN.CENTER)

# Bottom: gain table card
text(s, 0.5, 5.15, 12.3, 0.4,
     "Consistent accuracy gain on every baseline",
     font=SANS, size=12, bold=True, color=NAVY)

# Compact gain table
gx, gy, gw, gh = 0.5, 5.55, 7.0, 1.4
card(s, gx, gy, gw, gh, fill=CREAM, border=GRAY_LT, border_w=0.5)
gain_lines = [
    ("CIFAR-100, T=4   tdBN",        "75.55", "76.09"),
    ("                 TEBN",        "75.96", "76.45"),
    ("                 TAB",         "76.25", "77.24"),
    ("DVS-CIFAR10, T=10  tdBN",      "76.60", "77.37"),
]
ctx = s.shapes.add_textbox(Inches(gx + 0.2), Inches(gy + 0.1),
                            Inches(gw - 0.4), Inches(gh - 0.2))
tf = ctx.text_frame; tf.word_wrap = True
for i, (label, base, ours) in enumerate(gain_lines):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = f"{label:<32s} {base}  →  {ours}"
    p.font.name = "Consolas"; p.font.size = Pt(11); p.font.color.rgb = INK

# Right: highlighted +1.66
hx = 7.8
hbox = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
    Inches(hx), Inches(5.55), Inches(5.0), Inches(1.4))
hbox.fill.solid(); hbox.fill.fore_color.rgb = AMBER_LT
hbox.line.color.rgb = AMBER; hbox.line.width = Pt(1.2)
text(s, hx + 0.25, 5.65, 4.5, 0.55,
     "+1.66 %pt at T = 2",
     font=SERIF, size=22, bold=True, color=NAVY)
text(s, hx + 0.25, 6.2, 4.5, 0.65,
     "Biggest gain in the most latency-critical regime —\nexactly where SNN energy efficiency matters.",
     font=SANS, size=10, color=INK)

page_footer(s, 8)


# ════════════════════════════════════════════════════════════════════
# S9 — Two SG families
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Part II · TrSG")
slide_title(s, "Two Surrogate-Gradient Families",
            "AS-SG and RS-SG — same idea, two parameterizations")

card(s, 0.5, 2.5, 6.0, 4.0, fill=CREAM, border=NAVY, border_w=1.5)
text(s, 0.7, 2.65, 5.6, 0.4, "AS-SG  ·  Absolute-Scale",
     font=SANS, size=14, bold=True, color=NAVY)
text(s, 0.7, 3.1, 5.6, 0.5,
     "x = M[t] − V_thr",
     font="Cambria Math", size=18, color=NAVY)
bullets(s, 0.7, 3.85, 5.6, 2.5, [
    "Window width fixed at γ — ignores V_thr",
    "Magnitude independent of V_thr",
    "Used by tdBN, TET, PLIF, LTMD, …",
], size=12, color=INK, bullet="▸")

card(s, 6.85, 2.5, 6.0, 4.0, fill=CREAM, border=NAVY, border_w=1.5)
text(s, 7.05, 2.65, 5.6, 0.4, "RS-SG  ·  Relative-Scale",
     font=SANS, size=14, bold=True, color=NAVY)
text(s, 7.05, 3.1, 5.6, 0.5,
     "x = M[t] / V_thr − 1",
     font="Cambria Math", size=18, color=NAVY)
bullets(s, 7.05, 3.85, 5.6, 2.5, [
    "Window scales with V_thr  ✓",
    "But chain rule introduces a 1/V_thr factor in magnitude  ✗",
    "Used by Meng 2023, DIET-SNN",
], size=12, color=INK, bullet="▸")

callout_box(s, 0.5, 6.65, 12.3, 0.45,
            fill=NAVY, accent=CORAL,
            title="Treated as essentially equivalent. Once V_thr is trainable, they break in opposite ways.",
            title_color=AMBER, title_size=12, body_size=10)

page_footer(s, 9)


# ════════════════════════════════════════════════════════════════════
# S10 — Failure modes (table-style)
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Part II · TrSG")
slide_title(s, "Four Failure Modes — and One Fix")

text(s, 0.5, 2.4, 12.3, 0.4,
     "Window width × gradient magnitude — both must scale correctly with V_thr.",
     font=SANS, size=12, italic=True, color=GRAY)

# Header columns labels
text(s, 1.7, 2.85, 5.5, 0.3, "V_thr ≪ 1",
     font=SANS, size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER, spacing=200)
text(s, 7.4, 2.85, 5.5, 0.3, "V_thr ≫ 1",
     font=SANS, size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER, spacing=200)

# Row 1: AS-SG (coral)
ry = 3.2; rh = 1.0
card(s, 0.5, ry, 1.1, rh, fill=CORAL, border=CORAL, border_w=0)
text(s, 0.5, ry + 0.32, 1.1, 0.4, "AS-SG",
     font=SANS, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
card(s, 1.7, ry, 5.5, rh, fill=CORAL_LT, border=CORAL_LT, border_w=0.5)
text(s, 1.85, ry + 0.13, 5.2, 0.35, "✗  Flooding",
     font=SANS, size=12, bold=True, color=CORAL)
text(s, 1.85, ry + 0.5, 5.2, 0.5,
     "fixed window covers most of the distribution",
     font=SANS, size=11, color=INK)
card(s, 7.4, ry, 5.5, rh, fill=CORAL_LT, border=CORAL_LT, border_w=0.5)
text(s, 7.55, ry + 0.13, 5.2, 0.35, "✗  Starvation",
     font=SANS, size=12, bold=True, color=CORAL)
text(s, 7.55, ry + 0.5, 5.2, 0.5,
     "fixed window misses the distribution",
     font=SANS, size=11, color=INK)

# Row 2: RS-SG (amber)
ry2 = ry + rh + 0.15
card(s, 0.5, ry2, 1.1, rh, fill=AMBER, border=AMBER, border_w=0)
text(s, 0.5, ry2 + 0.32, 1.1, 0.4, "RS-SG",
     font=SANS, size=12, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
card(s, 1.7, ry2, 5.5, rh, fill=AMBER_LT, border=AMBER_LT, border_w=0.5)
text(s, 1.85, ry2 + 0.13, 5.2, 0.35, "✗  Explosion",
     font=SANS, size=12, bold=True, color=CORAL)
text(s, 1.85, ry2 + 0.5, 5.2, 0.5,
     "1/V_thr blows the magnitude up",
     font=SANS, size=11, color=INK)
card(s, 7.4, ry2, 5.5, rh, fill=AMBER_LT, border=AMBER_LT, border_w=0.5)
text(s, 7.55, ry2 + 0.13, 5.2, 0.35, "✗  Vanishing",
     font=SANS, size=12, bold=True, color=CORAL)
text(s, 7.55, ry2 + 0.5, 5.2, 0.5,
     "1/V_thr shrinks the magnitude",
     font=SANS, size=11, color=INK)

# Row 3: TrSG (teal)
ry3 = ry2 + rh + 0.15
card(s, 0.5, ry3, 1.1, rh, fill=TEAL, border=TEAL, border_w=0)
text(s, 0.5, ry3 + 0.32, 1.1, 0.4, "TrSG",
     font=SANS, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
card(s, 1.7, ry3, 5.5, rh, fill=TEAL_LT, border=TEAL_LT, border_w=0.5)
text(s, 1.85, ry3 + 0.13, 5.2, 0.35, "✓  Balanced",
     font=SANS, size=12, bold=True, color=TEAL)
text(s, 1.85, ry3 + 0.5, 5.2, 0.5,
     "window scales with V_thr,  magnitude stays 1",
     font=SANS, size=11, color=INK)
card(s, 7.4, ry3, 5.5, rh, fill=TEAL_LT, border=TEAL_LT, border_w=0.5)
text(s, 7.55, ry3 + 0.13, 5.2, 0.35, "✓  Balanced",
     font=SANS, size=12, bold=True, color=TEAL)
text(s, 7.55, ry3 + 0.5, 5.2, 0.5,
     "window scales with V_thr,  magnitude stays 1",
     font=SANS, size=11, color=INK)

text(s, 0.5, 6.8, 12.3, 0.3,
     "(Geometric picture in the thesis Figure 5.1.)",
     font=SANS, size=9, italic=True, color=GRAY, align=PP_ALIGN.CENTER)

page_footer(s, 10)


# ════════════════════════════════════════════════════════════════════
# S11 — One-line fix
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Part II · TrSG")
slide_title(s, "TrSG — One-Line Fix",
            "Multiply the threshold on the forward path; the chain rule cancels the bad factor")

# Step 1
text(s, 0.5, 2.4, 12.3, 0.4,
     "Step 1.  Use the relative-scale window (RS-SG style):",
     font=SANS, size=13, bold=True, color=NAVY)
callout_box(s, 0.5, 2.9, 12.3, 0.7,
            fill=CREAM, accent=TEAL,
            body="x  =  M[t] / V_thr   −   1",
            body_color=INK, body_size=14)

# Step 2
text(s, 0.5, 3.85, 12.3, 0.4,
     "Step 2.  Multiply V_thr on the forward path  (this is the only change):",
     font=SANS, size=13, bold=True, color=NAVY)
callout_box(s, 0.5, 4.35, 12.3, 0.7,
            fill=CREAM, accent=AMBER,
            body="O[t]  =  V_thr  ·  S[t]",
            body_color=INK, body_size=14)

# Step 3 — chain rule
text(s, 0.5, 5.3, 12.3, 0.4,
     "Step 3.  The chain rule cancels the 1/V_thr factor exactly:",
     font=SANS, size=13, bold=True, color=NAVY)
callout_box(s, 0.5, 5.8, 12.3, 0.85,
            fill=NAVY, accent=AMBER,
            body="∂O / ∂M   =   V_thr  ·  ∂S/∂M   ≈   V_thr · (1/V_thr) f'(x)   =   f'(x)",
            body_color=WHITE, body_size=15)

text(s, 0.5, 6.8, 12.3, 0.3,
     "Window adapts with V_thr   ·   gradient magnitude is threshold-invariant.",
     font=SANS, size=12, bold=True, italic=True, color=TEAL, align=PP_ALIGN.CENTER)

page_footer(s, 11)


# ════════════════════════════════════════════════════════════════════
# S12 — Inference stays binary
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Part II · TrSG")
slide_title(s, "Inference Stays Binary",
            "TrSG is a training-only trick — neuromorphic hardware never sees it")

# Training row
text(s, 0.5, 2.5, 12.3, 0.4,
     "Training",
     font=SANS, size=14, bold=True, color=NAVY)
callout_box(s, 0.5, 3.0, 12.3, 0.7,
            fill=CREAM, accent=AMBER,
            body="Forward:    O[t]  =  V_thr · S[t]    (real-valued — gradient flows cleanly through V_thr)",
            body_color=INK, body_size=12)

# Inference rows
text(s, 0.5, 3.95, 12.3, 0.4,
     "At deployment, absorb V_thr into the next layer's weights",
     font=SANS, size=14, bold=True, color=NAVY)
callout_box(s, 0.5, 4.45, 12.3, 0.7,
            fill=CREAM, accent=AMBER,
            body="W'_{l→l+1}   =   V_thr  ·  W_{l→l+1}    (one-time rescaling)",
            body_color=INK, body_size=12)
callout_box(s, 0.5, 5.25, 12.3, 0.7,
            fill=TEAL_LT, accent=TEAL,
            body="Forward at test time:    O[t]  =  S[t]  ∈  {0, 1}      (standard LIF — back to binary spikes)",
            body_color=NAVY, body_size=12)

# Bottom emphasis
callout_box(s, 0.5, 6.2, 12.3, 0.7,
            fill=TEAL, accent=AMBER,
            title="Standard LIF operators at inference. Hardware compatibility preserved.",
            title_color=WHITE, title_size=13)

page_footer(s, 12)


# ════════════════════════════════════════════════════════════════════
# S13 — TrSG: Result + Mechanism
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Part II · TrSG")
slide_title(s, "TrSG Holds Up — Stress Tests + Real Trajectories")

# Left: stability metrics tables
text(s, 0.5, 2.4, 6.5, 0.4,
     "Gradient-stability metrics at extreme thresholds  (CIFAR-100, T=4, ep 100)",
     font=SANS, size=11, italic=True, color=GRAY)

# Two table cards side by side
def metric_card(slide, x, y, title, rows):
    card(slide, x, y, 3.05, 2.4, fill=CREAM, border=GRAY_LT, border_w=0.5)
    text(slide, x + 0.15, y + 0.08, 2.85, 0.3,
         title, font=SANS, size=11, bold=True, color=NAVY)
    tx = slide.shapes.add_textbox(Inches(x + 0.15), Inches(y + 0.4),
                                   Inches(2.85), Inches(2.0))
    tf = tx.text_frame; tf.word_wrap = True
    for i, (label, val, color) in enumerate(rows):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"{label:<8s} {val}"
        p.font.name = "Consolas"; p.font.size = Pt(10); p.font.color.rgb = color
        if i > 0: p.space_before = Pt(2)

metric_card(s, 0.5, 2.85,
            "V_thr = 0.1   (low)",
            [("AS-SG", "AbsStr 0.0145  · 90% active", CORAL),
             ("RS-SG", "diverged (NaN)", CORAL),
             ("TrSG",  "AbsStr 0.0153  · 21% · CV 0.87", TEAL)])
metric_card(s, 3.7, 2.85,
            "V_thr = 2.0   (high)",
            [("AS-SG", "0%  active  (zero gradient)", CORAL),
             ("RS-SG", "0.10% active · CV 2.25", CORAL),
             ("TrSG",  "AbsStr 0.0323  · 7% · CV 0.73", TEAL)])

# Right: threshold trajectory image
text(s, 7.3, 2.4, 5.6, 0.4,
     "Real trainings actually visit these regions",
     font=SANS, size=11, italic=True, color=GRAY)
img(s, ASSETS + "/threshold_training_imagenet.png", 7.3, 2.85, w=5.6, h=2.4)

# Bottom message
callout_box(s, 0.5, 5.5, 12.3, 1.4,
            fill=NAVY, accent=AMBER,
            title="Threshold-robustness is a practical necessity",
            body="Across CIFAR-100, DVS-CIFAR10, and ImageNet, trainable V_thr drifts well below 1\n"
                 "or above 2 in some layers — exactly the regimes where AS-SG and RS-SG fail.",
            title_color=AMBER, body_color=WHITE,
            title_size=14, body_size=11)

page_footer(s, 13)


# ════════════════════════════════════════════════════════════════════
# S14 — Main Results
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Experiments · 1/3")
slide_title(s, "Main Results — SOTA Across the Board")

text(s, 0.5, 2.4, 12.3, 0.4,
     "Standard LIF neurons.  No AutoAugment / Cutout.  Three trials each.",
     font=SANS, size=11, italic=True, color=GRAY)

# Bar chart area on left + key gain cards on right
# Replace bar chart with a table-like card showing all numbers
chart_card_x, chart_card_y, chart_card_w, chart_card_h = 0.5, 2.85, 8.0, 4.1
card(s, chart_card_x, chart_card_y, chart_card_w, chart_card_h,
     fill=CREAM, border=GRAY_LT, border_w=0.5)
text(s, chart_card_x + 0.25, chart_card_y + 0.2, chart_card_w - 0.5, 0.4,
     "Top-1 accuracy (%) — Ours vs. previous best",
     font=SANS, size=12, bold=True, color=TEAL)

# Result table inside
res_lines = [
    ("CIFAR-10  (R-19, T=6)",      "TAB 94.81",  "95.50",   "+0.69"),
    ("CIFAR-100 (R-19, T=6)",      "TAB 76.82",  "77.91",   "+1.09"),
    ("ImageNet  (R-34, T=6)",      "TAB 67.78",  "68.73",   "+0.95"),
    ("ImageNet  (SEW-R34, T=4)",   "MPS 69.03",  "69.67",   "+0.64"),
    ("DVS-CIFAR10 (R-19, T=10)",   "RMP 76.20",  "81.43",   "+5.23"),
    ("DVS-CIFAR10 (7-CNN, T=4)",   "TAB 76.70",  "79.27",   "+2.57"),
]
ctx = s.shapes.add_textbox(Inches(chart_card_x + 0.3), Inches(chart_card_y + 0.7),
                            Inches(chart_card_w - 0.6), Inches(chart_card_h - 0.9))
tf = ctx.text_frame; tf.word_wrap = True
for i, (dset, prev, ours, delta) in enumerate(res_lines):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = f"{dset:<26s}{prev:<14s}→  {ours:<8s}{delta:>6s}"
    p.font.name = "Consolas"; p.font.size = Pt(11); p.font.color.rgb = INK
    if i > 0: p.space_before = Pt(4)

# Right: 3 gain cards
def gain_card(slide, x, y, value, label, color):
    card(slide, x, y, 4.4, 1.25, fill=CREAM, border=color, border_w=1.5)
    text(slide, x + 0.2, y + 0.15, 2.5, 0.55, value,
         font=SERIF, size=26, bold=True, color=color)
    text(slide, x + 2.7, y + 0.27, 1.5, 0.4, "%pt",
         font=SANS, size=11, color=color)
    text(slide, x + 0.2, y + 0.75, 4.0, 0.45, label,
         font=SANS, size=10, color=INK)

gain_card(s, 8.7, 2.85, "+1.09", "CIFAR-100 over TAB", TEAL)
gain_card(s, 8.7, 4.25, "+0.64", "ImageNet (SEW-R34) over MPS", TEAL)
gain_card(s, 8.7, 5.65, "+5.23", "DVS-CIFAR10 over RMP-Loss", AMBER)

page_footer(s, 14)


# ════════════════════════════════════════════════════════════════════
# S15 — Ablation
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Experiments · 2/3")
slide_title(s, "Ablation — Independent and Complementary",
            "Each method helps alone; together they compound (especially on event-based data)")

# Two ablation tables side by side
text(s, 0.5, 2.5, 6.0, 0.4,
     "CIFAR-100   (T = 4)",
     font=SANS, size=12, bold=True, color=NAVY)
card(s, 0.5, 2.95, 6.0, 3.1, fill=CREAM, border=GRAY_LT, border_w=0.5)

cifar_rows = [
    ("✗", "✗", "75.55"),
    ("✓", "✗", "76.09"),
    ("✗", "✓", "77.15"),
    ("✓", "✓", "77.66"),
]
ctx = s.shapes.add_textbox(Inches(0.7), Inches(3.1), Inches(5.8), Inches(2.8))
tf = ctx.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "MP-Init    TrSG       Acc (%)"
p.font.name = "Consolas"; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = TEAL
p.space_after = Pt(8)
for i, (a, b, v) in enumerate(cifar_rows):
    p = tf.add_paragraph()
    bold = (i == len(cifar_rows) - 1)
    p.text = f"   {a}        {b}         {v}"
    p.font.name = "Consolas"; p.font.size = Pt(13); p.font.bold = bold
    p.font.color.rgb = TEAL if bold else INK
    p.space_before = Pt(4)

text(s, 6.85, 2.5, 6.0, 0.4,
     "DVS-CIFAR10   (T = 10)",
     font=SANS, size=12, bold=True, color=NAVY)
card(s, 6.85, 2.95, 6.0, 3.1, fill=CREAM, border=GRAY_LT, border_w=0.5)

dvs_rows = [
    ("✗", "✗", "76.60"),
    ("✓", "✗", "77.37"),
    ("✗", "✓", "80.83"),
    ("✓", "✓", "81.43"),
]
ctx = s.shapes.add_textbox(Inches(7.05), Inches(3.1), Inches(5.8), Inches(2.8))
tf = ctx.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "MP-Init    TrSG       Acc (%)"
p.font.name = "Consolas"; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = TEAL
p.space_after = Pt(8)
for i, (a, b, v) in enumerate(dvs_rows):
    p = tf.add_paragraph()
    bold = (i == len(dvs_rows) - 1)
    p.text = f"   {a}        {b}         {v}"
    p.font.name = "Consolas"; p.font.size = Pt(13); p.font.bold = bold
    p.font.color.rgb = AMBER if bold else INK
    p.space_before = Pt(4)

# Bottom message
callout_box(s, 0.5, 6.2, 12.3, 0.75,
            fill=TEAL, accent=AMBER,
            title="MP-Init alone +0.5 ~ 0.8.  TrSG alone +1.6 ~ 4.2.  Together +2.1 ~ 4.8.",
            title_color=WHITE, title_size=12)

page_footer(s, 15)


# ════════════════════════════════════════════════════════════════════
# S16 — Generalization
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Experiments · 3/3")
slide_title(s, "Beyond Classification",
            "Transformer SNNs · object detection · energy preserved")

# Three cards
def gen_card(slide, x, y, w, h, header, body):
    card(slide, x, y, w, h, fill=CREAM, border=TEAL, border_w=1.2)
    text(slide, x + 0.2, y + 0.15, w - 0.4, 0.4,
         header, font=SANS, size=12, bold=True, color=TEAL)
    text(slide, x + 0.2, y + 0.55, w - 0.4, h - 0.7, body,
         font=SANS, size=11, color=INK)

gen_card(s, 0.5, 2.5, 6.0, 1.7,
         "Transformer SNN  (QKFormer, ImageNet, T=4)",
         "Baseline 78.80    →    + ours 79.90    (+1.10 %pt)\n"
         "AS-SG    77.09    →    RS-SG  78.70    →    TrSG  79.90\n"
         "Drop-in SG swap on a state-of-the-art Transformer SNN.")

gen_card(s, 6.85, 2.5, 6.0, 1.7,
         "Object detection  (COCO-2017, EMS-ResNet10, T=3)",
         "mAP@0.5         0.291  →  0.305    (+0.014)\n"
         "mAP@0.5:0.95    0.138  →  0.147    (+0.009)\n"
         "First evidence the principles transfer beyond classification.")

gen_card(s, 0.5, 4.4, 6.0, 1.7,
         "Energy at matched firing rate  (CIFAR-100)",
         "Baseline   1.57 mJ   ·   75.58 %\n"
         "+ MP-Init + TrSG    1.67 mJ   ·   77.68 %\n"
         "+2.10 %pt at +6 % energy — gains aren't from spiking more.")

gen_card(s, 6.85, 4.4, 6.0, 1.7,
         "All four surrogate shapes",
         "Rectangular · Triangular · Arctan · Sigmoid\n"
         "TrSG remains the best across initial V_thr, initial τ,\n"
         "and surrogate shape — never tied for second.")

# Bottom emphasis
callout_box(s, 0.5, 6.25, 12.3, 0.75,
            fill=NAVY, accent=AMBER,
            title="No architectural changes, no extra losses, no per-timestep parameters — just drop in.",
            title_color=AMBER, title_size=13)

page_footer(s, 16)


# ════════════════════════════════════════════════════════════════════
# S17 — Recap
# ════════════════════════════════════════════════════════════════════
s = add_blank()
section_label(s, "Conclusion")
slide_title(s, "What This Talk Was About")

# Top: one-line story
callout_box(s, 0.5, 2.4, 12.3, 1.0,
            fill=NAVY, accent=AMBER,
            title="Two long-standing instabilities of direct SNN training, traced and fixed.",
            body="Both fixes are minimal, principled, and preserve standard inference.",
            title_color=AMBER, body_color=WHITE,
            title_size=15, body_size=11)

# Three cards: MP-Init, TrSG, Validated
def takeaway_card(slide, x, y, w, h, header, header_color, body):
    card(slide, x, y, w, h, fill=CREAM, border=header_color, border_w=1.5)
    text(slide, x + 0.2, y + 0.15, w - 0.4, 0.5,
         header, font=SERIF, size=18, bold=True, color=header_color)
    text(slide, x + 0.2, y + 0.7, w - 0.4, h - 0.8,
         body, font=SANS, size=11, color=INK)

takeaway_card(s, 0.5, 3.7, 4.0, 3.0,
              "MP-Init", TEAL,
              "TCS originates in the membrane potential — fix the initial state.\n\n"
              "Per-layer running mean of U[T].\n\n"
              "Zero overhead, standard LIF inference.")

takeaway_card(s, 4.7, 3.7, 4.0, 3.0,
              "TrSG", AMBER,
              "Threshold-induced gradient pathology — fix it with one forward-path multiplication.\n\n"
              "Window adapts, magnitude is invariant.\n\n"
              "Standard LIF after training.")

takeaway_card(s, 8.9, 3.7, 4.0, 3.0,
              "Validated", TEAL,
              "SOTA on CIFAR-10/100, ImageNet, DVS-CIFAR10.\n\n"
              "Generalizes to Transformer SNNs (QKFormer +1.10) and detection (COCO +0.014 mAP).\n\n"
              "All without architectural changes.")

page_footer(s, 17)


# ════════════════════════════════════════════════════════════════════
# S18 — Limitations + Future + Thanks
# ════════════════════════════════════════════════════════════════════
s = add_blank()
fill_bg(s, NAVY)
add_circle(s, 9.0, -1.5, 4.0, NAVY_DK)
add_circle(s, -2.0, 5.0, 5.5, NAVY_DK)

text(s, 0.7, 0.7, 10, 0.4, "CONCLUSION",
     font=SANS, size=12, bold=True, color=AMBER, spacing=400)
text(s, 0.7, 1.4, 12, 0.7,
     "Two instabilities, two principled fixes,",
     font=SERIF, size=32, bold=True, color=WHITE)
text(s, 0.7, 2.1, 12, 0.7,
     "one cohesive story.",
     font=SERIF, size=32, bold=True, color=WHITE)

# Three summary boxes
def cl_box(slide, x, y, w, h, title, body, accent_color):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    box.fill.solid(); box.fill.fore_color.rgb = NAVY
    box.line.color.rgb = accent_color; box.line.width = Pt(1.2)
    text(slide, x + 0.25, y + 0.2, w - 0.5, 0.4, title,
         font=SERIF, size=18, bold=True, color=accent_color)
    text(slide, x + 0.25, y + 0.75, w - 0.5, h - 0.95, body,
         font=SANS, size=10, color=RGBColor(0xCB, 0xD8, 0xE0))

cl_box(s, 0.5, 3.1, 4.0, 1.85,
       "MP-Init",
       "TCS originates in the membrane potential. A per-layer running mean closes the gap to the stationary distribution. No overhead, standard inference.",
       TEAL)
cl_box(s, 4.7, 3.1, 4.0, 1.85,
       "TrSG",
       "Threshold-induced gradient pathology cancelled by a one-line forward-path fix. Window adapts, magnitude is invariant. Inference reverts to binary spikes.",
       AMBER)
cl_box(s, 8.9, 3.1, 4.0, 1.85,
       "Validated",
       "SOTA on CIFAR-10/100, ImageNet, DVS-CIFAR10. Generalizes to Transformer SNNs and event-based detection without modification.",
       TEAL)

# Limitations + future
text(s, 0.5, 5.25, 11, 0.3, "LIMITATIONS  &  FUTURE  WORK",
     font=SANS, size=10, bold=True, color=AMBER, spacing=400)
text(s, 0.5, 5.6, 12.3, 1.0,
     "Per-layer constant approximates π only at the mean — richer parameterizations are open.\n"
     "Hardware with fixed neuron parameters needs quantization-aware adaptation.\n"
     "Beyond LIF: extending the principles to adaptive LIF, Izhikevich, and large foundation-scale SNNs.",
     font=SANS, size=11, color=RGBColor(0xCB, 0xD8, 0xE0))

# Amber underline + Thank you
amber_underline(s, 0.7, 6.85, w=1.0)
text(s, 0.7, 7.0, 12, 0.4,
     "Thank you.   Questions?",
     font=SERIF, size=20, bold=True, italic=True, color=AMBER)
text(s, 11.0, 7.10, 1.8, 0.3, "18 / 18",
     font=SANS, size=10, color=RGBColor(0x80, 0x95, 0xA8), align=PP_ALIGN.RIGHT)


# ─── Save ──────────────────────────────────────────────────────────
out = "Hyunho_Kook_Master_Defense.pptx"
prs.save(out)
print(f"Saved: {out}    Slides: {len(prs.slides)}")
