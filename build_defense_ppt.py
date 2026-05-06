"""
Master's thesis defense slides
Hyunho Kook -- Stabilizing Direct Training of Spiking Neural Networks
22 slides, ~20 min defense
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os

ASSETS = "ppt_assets"

# 16:9 widescreen
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Color palette
NAVY    = RGBColor(0x14, 0x36, 0x5C)   # primary dark
RED     = RGBColor(0xC0, 0x39, 0x2B)   # accent
GRAY    = RGBColor(0x55, 0x55, 0x55)
LIGHT   = RGBColor(0xF4, 0xF4, 0xF4)
GREEN   = RGBColor(0x2E, 0x7D, 0x32)
ORANGE  = RGBColor(0xE6, 0x77, 0x22)


def add_blank_slide():
    return prs.slides.add_slide(prs.slide_layouts[6])  # 6 = blank


def add_title_bar(slide, text, sub=None):
    # Top accent bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(0), Inches(0),
                                  prs.slide_width, Inches(0.07))
    bar.fill.solid(); bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()

    # Title text
    tx = slide.shapes.add_textbox(Inches(0.5), Inches(0.22),
                                   Inches(12.3), Inches(0.7))
    tf = tx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = NAVY

    if sub:
        sx = slide.shapes.add_textbox(Inches(0.5), Inches(0.85),
                                       Inches(12.3), Inches(0.4))
        sp = sx.text_frame.paragraphs[0]
        sp.text = sub
        sp.font.size = Pt(14); sp.font.color.rgb = GRAY


def add_text(slide, x, y, w, h, text, *,
             size=18, bold=False, color=None, align=None):
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    if color: p.font.color.rgb = color
    if align: p.alignment = align
    return tx


def add_bullets(slide, x, y, w, h, items, *, size=18, color=None):
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tx.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(item, tuple):
            text, lvl = item
        else:
            text, lvl = item, 0
        p.text = ('• ' if lvl == 0 else '   – ') + text
        p.font.size = Pt(size)
        p.level = lvl
        if color: p.font.color.rgb = color
    return tx


def add_image(slide, path, x, y, w=None, h=None):
    if not os.path.exists(path):
        return None
    if w is not None and h is not None:
        return slide.shapes.add_picture(path, Inches(x), Inches(y),
                                         width=Inches(w), height=Inches(h))
    elif w is not None:
        return slide.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(w))
    elif h is not None:
        return slide.shapes.add_picture(path, Inches(x), Inches(y), height=Inches(h))
    else:
        return slide.shapes.add_picture(path, Inches(x), Inches(y))


def add_footer(slide, page_num, total=22):
    tx = slide.shapes.add_textbox(Inches(11.5), Inches(7.05),
                                   Inches(1.6), Inches(0.3))
    p = tx.text_frame.paragraphs[0]
    p.text = f"{page_num} / {total}"
    p.font.size = Pt(10); p.font.color.rgb = GRAY
    p.alignment = PP_ALIGN.RIGHT


def add_box(slide, x, y, w, h, text, *,
            fill=None, outline=None, font_size=16, font_color=None, bold=True):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if outline:
        shp.line.color.rgb = outline; shp.line.width = Pt(1.5)
    else:
        shp.line.fill.background()
    tf = shp.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.1)
    tf.margin_top = tf.margin_bottom = Inches(0.05)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size); p.font.bold = bold
    p.font.color.rgb = font_color or NAVY
    p.alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return shp


# ─────────────────────────────────────────────────────────────────
# S1. Title
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
# Background bar at bottom
bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                          Inches(0), Inches(7.0),
                          prs.slide_width, Inches(0.5))
bar.fill.solid(); bar.fill.fore_color.rgb = NAVY; bar.line.fill.background()

# Top accent line
acc = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                          Inches(0), Inches(0),
                          prs.slide_width, Inches(0.15))
acc.fill.solid(); acc.fill.fore_color.rgb = NAVY; acc.line.fill.background()

add_text(s, 0.7, 1.5, 12, 1.0,
         "Master's Thesis Defense", size=20, color=GRAY)

add_text(s, 0.7, 2.1, 12, 1.6,
         "Enhancing Stability in Direct Training of",
         size=36, bold=True, color=NAVY)
add_text(s, 0.7, 2.8, 12, 1.6,
         "Spiking Neural Networks",
         size=36, bold=True, color=NAVY)
add_text(s, 0.7, 3.7, 12, 0.8,
         "Membrane Potential Initialization &",
         size=24, color=RED)
add_text(s, 0.7, 4.2, 12, 0.8,
         "Threshold-Robust Surrogate Gradient",
         size=24, color=RED)

add_text(s, 0.7, 5.5, 12, 0.4,
         "Hyunho Kook (국 현 호)", size=20, bold=True, color=NAVY)
add_text(s, 0.7, 5.95, 12, 0.4,
         "Department of Computer Science and Engineering, POSTECH",
         size=14, color=GRAY)
add_text(s, 0.7, 6.3, 12, 0.4,
         "Advisor: Prof. Eunhyeok Park", size=14, color=GRAY)

# Date placeholder bottom-right
tx = s.shapes.add_textbox(Inches(11), Inches(7.05), Inches(2.2), Inches(0.4))
p = tx.text_frame.paragraphs[0]
p.text = "January 2026"
p.font.size = Pt(12); p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
p.alignment = PP_ALIGN.RIGHT


# ─────────────────────────────────────────────────────────────────
# S2. Why SNNs?
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Why Spiking Neural Networks?",
              "Energy-efficient AI through event-driven computation")

# Left column: motivation
add_text(s, 0.5, 1.5, 6, 0.5,
         "The energy problem", size=20, bold=True, color=NAVY)
add_bullets(s, 0.5, 2.0, 6.2, 3,
            ["Modern DNNs scale to megawatt power budgets",
             "Edge inference faces tight energy constraints",
             "SNNs: spike-based, event-driven  →  efficient on neuromorphic HW"],
            size=16)

# Right column: status
add_text(s, 7.0, 1.5, 6, 0.5,
         "Current state of SNN training", size=20, bold=True, color=NAVY)
add_bullets(s, 7.0, 2.0, 6, 3,
            ["Direct training narrows the accuracy gap",
             "Yet two persistent instabilities remain"],
            size=16)

# Bottom: two boxes highlighting the two problems
add_box(s, 0.5, 5.0, 6, 1.6,
        "① Temporal Covariate Shift\n(membrane potential drift across timesteps)",
        outline=RED, font_size=18, font_color=RED)
add_box(s, 7.0, 5.0, 6, 1.6,
        "② Gradient Pathology\n(when V_thr is trainable, gradients explode/vanish)",
        outline=RED, font_size=18, font_color=RED)

add_text(s, 0.5, 6.7, 13, 0.4,
         "This thesis: precise diagnoses + minimal interventions for both.",
         size=16, bold=True, color=GRAY, align=PP_ALIGN.CENTER)
add_footer(s, 2)


# ─────────────────────────────────────────────────────────────────
# S3. LIF Neuron
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "LIF Neuron in 30 Seconds",
              "The discrete-time leaky integrate-and-fire model")

# Three equations on the left
add_text(s, 0.6, 1.4, 6.4, 0.4,
         "Three updates per timestep", size=18, bold=True, color=NAVY)

# Equation boxes
eq_y = 2.0
add_box(s, 0.6, eq_y,        6.4, 0.85,
        "Integrate:    M[t] = (1 − 1/τ)·U[t−1]  +  (1/τ)·I_in[t]",
        outline=NAVY, font_size=14, bold=False, font_color=NAVY)
add_box(s, 0.6, eq_y + 1.0,  6.4, 0.85,
        "Fire:    S[t] = H( M[t] − V_thr )       (1 if exceeds, else 0)",
        outline=NAVY, font_size=14, bold=False, font_color=NAVY)
add_box(s, 0.6, eq_y + 2.0,  6.4, 0.85,
        "Reset:   U[t] = M[t] − V_thr · S[t]    (soft reset)",
        outline=NAVY, font_size=14, bold=False, font_color=NAVY)

# Right side: cartoon description
add_text(s, 7.4, 1.4, 5.5, 0.4,
         "Intuition", size=18, bold=True, color=NAVY)
add_bullets(s, 7.4, 1.9, 5.5, 4.5,
            ["M[t]: membrane potential — the running 'sum'",
             "V_thr: firing threshold",
             "When M[t] crosses V_thr, the neuron fires (S=1) and resets",
             "Two key learnable parameters: V_thr, τ",
             "All temporal dynamics happen here"],
            size=15)

add_text(s, 0.6, 6.5, 12, 0.5,
         "Everything in this thesis revolves around the membrane potential M[t] and threshold V_thr.",
         size=14, bold=True, color=RED, align=PP_ALIGN.CENTER)
add_footer(s, 3)


# ─────────────────────────────────────────────────────────────────
# S4. Direct Training & Surrogate Gradient
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Direct Training & Surrogate Gradient",
              "Learning through the discontinuous spike function")

add_text(s, 0.6, 1.4, 12, 0.5,
         "Heaviside spike → not differentiable", size=18, bold=True, color=NAVY)

# Left: H(x) explanation
add_box(s, 0.6, 2.1, 5.8, 1.2,
        "S = H(M − V_thr)\nH'(x) = δ(x)  (Dirac delta)\nGradient blocked",
        fill=LIGHT, outline=GRAY, font_size=14, font_color=GRAY)

# Right: surrogate
add_box(s, 6.8, 2.1, 5.8, 1.2,
        "Replace H with smooth surrogate f(x)\nf'(x) ≈ H'(x) near 0\nGradient flows",
        fill=LIGHT, outline=NAVY, font_size=14, font_color=NAVY)

# Common surrogate shapes
add_text(s, 0.6, 3.6, 12, 0.4,
         "Common surrogate-derivative shapes", size=16, bold=True, color=NAVY)

add_box(s, 0.6, 4.1, 3.0, 0.8,
        "Rectangular\nf' = 1/γ · 1{|x|<γ/2}",
        outline=NAVY, font_size=12, font_color=NAVY)
add_box(s, 3.8, 4.1, 3.0, 0.8,
        "Triangular\nf' = max(0, γ−|x|)/γ²",
        outline=NAVY, font_size=12, font_color=NAVY)
add_box(s, 7.0, 4.1, 3.0, 0.8,
        "Arctan\nf' = γ / (2(1+(πγx/2)²))",
        outline=NAVY, font_size=12, font_color=NAVY)
add_box(s, 10.2, 4.1, 3.0, 0.8,
        "Sigmoid\nf' = (1/γ)·σ(x/γ)(1−σ)",
        outline=NAVY, font_size=12, font_color=NAVY)

# Foreshadow
add_text(s, 0.6, 5.6, 13, 1.5,
         "→ The choice of where x is centered (M − V_thr  vs.  M/V_thr − 1)\n"
         "    looks innocent, but becomes critical once V_thr is learnable.\n"
         "    (The other half of this talk.)",
         size=16, color=RED, bold=True)
add_footer(s, 4)


# ─────────────────────────────────────────────────────────────────
# S5. TCS — what's the problem
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Temporal Covariate Shift (TCS)",
              "Per-timestep activation distributions drift apart")

add_image(s, ASSETS + "/density_plot_membrane_potential_tdbn.png",
          0.5, 1.4, w=6.5)

add_text(s, 7.5, 1.4, 5.5, 0.5,
         "Symptom", size=18, bold=True, color=NAVY)
add_bullets(s, 7.5, 1.9, 5.5, 4,
            ["Membrane potential M[t] drifts across timesteps",
             "BN statistics estimated at one timestep no longer fit at another",
             "Spike outputs swing → accuracy degrades"],
            size=14)

add_text(s, 7.5, 4.5, 5.5, 0.5,
         "Existing fixes are partial", size=18, bold=True, color=RED)
add_bullets(s, 7.5, 5.0, 5.5, 1.5,
            ["TEBN, TAB: per-timestep BN parameters",
             "More parameters, more compute",
             "And — they don't actually fix M[t]"],
            size=14, color=RED)

add_text(s, 0.5, 6.6, 12.5, 0.4,
         "tdBN (no TCS handling): each timestep curve sits at a different center",
         size=12, color=GRAY, align=PP_ALIGN.CENTER)
add_footer(s, 5)


# ─────────────────────────────────────────────────────────────────
# S6. Trainable V_thr
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Why is Trainable V_thr Hard?",
              "The threshold sits in the surrogate's argument — small changes warp the gradient")

add_bullets(s, 0.6, 1.5, 12.2, 1.5,
            ["V_thr controls when a neuron fires — fundamental to LIF behavior",
             "Modern training methods (PLIF, LTMD, DIET-SNN) make it trainable",
             "But: existing surrogate gradients are highly sensitive to V_thr scale"],
            size=16)

# Visual: two scenarios
add_text(s, 0.6, 3.4, 12.2, 0.5,
         "What can go wrong (preview of contribution 2)", size=18, bold=True, color=RED)

add_box(s, 0.6, 4.0, 6.0, 2.0,
        "V_thr too small\n→ gradient floods the layer\nor explodes through 1/V_thr",
        fill=LIGHT, outline=RED, font_size=16, font_color=RED)
add_box(s, 6.8, 4.0, 6.0, 2.0,
        "V_thr too large\n→ gradient starves\nor vanishes through 1/V_thr",
        fill=LIGHT, outline=RED, font_size=16, font_color=RED)

add_text(s, 0.6, 6.4, 12.2, 0.5,
         "→ Learning V_thr blindly is dangerous. We diagnose & fix in Part B.",
         size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
add_footer(s, 6)


# ─────────────────────────────────────────────────────────────────
# S7. Two Open Problems
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Two Open Problems",
              "Unsolved by the existing literature")

# Big numbered boxes
add_box(s, 0.6, 1.7, 12.2, 2.3,
        "①  TCS in the membrane potential is not addressed by existing normalization.\n\n"
        "    TEBN / TAB stabilize activations but the underlying M[t] keeps drifting.",
        fill=LIGHT, outline=NAVY, font_size=18, font_color=NAVY)

add_box(s, 0.6, 4.3, 12.2, 2.3,
        "②  Gradient pathology under trainable V_thr has never been explicitly analyzed.\n\n"
        "    Prior surrogate-gradient methods are not designed to be threshold-robust.",
        fill=LIGHT, outline=RED, font_size=18, font_color=RED)

add_text(s, 0.6, 6.85, 12.2, 0.4,
         "Both problems persist despite years of SNN training research.",
         size=14, color=GRAY, align=PP_ALIGN.CENTER)
add_footer(s, 7)


# ─────────────────────────────────────────────────────────────────
# S8. Contributions
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Contributions",
              "Two methods, one principled story")

# Left: MP-Init contribution
add_box(s, 0.5, 1.5, 6.2, 4.8,
        "MP-Init",
        fill=NAVY, outline=NAVY, font_size=22, font_color=RGBColor(0xFF,0xFF,0xFF))
add_text(s, 0.5, 2.4, 6.2, 0.4,
         "Membrane Potential Initialization",
         size=14, color=GRAY, align=PP_ALIGN.CENTER)
add_bullets(s, 0.7, 3.0, 5.8, 3.2,
            ["Trace TCS to its root: membrane potential",
             "Theorem: M[t] converges geometrically to a unique stationary π (Doeblin)",
             "Algorithm: per-layer running mean of U[T] aligns initial state with π",
             "Eliminates TCS — at zero parameter overhead",
             "Inference: standard LIF, fully HW-compatible"],
            size=13)

# Right: TrSG contribution
add_box(s, 6.9, 1.5, 6.2, 4.8,
        "TrSG",
        fill=RED, outline=RED, font_size=22, font_color=RGBColor(0xFF,0xFF,0xFF))
add_text(s, 6.9, 2.4, 6.2, 0.4,
         "Threshold-Robust Surrogate Gradient",
         size=14, color=GRAY, align=PP_ALIGN.CENTER)
add_bullets(s, 7.1, 3.0, 5.8, 3.2,
            ["Classify existing SGs into AS-SG / RS-SG families",
             "Show how each fails when V_thr changes scale",
             "Fix: forward pass O = V_thr · S — cancels 1/V_thr factor exactly",
             "Gradient magnitude becomes threshold-invariant",
             "Inference: weight-rescaling restores binary spikes"],
            size=13)

add_text(s, 0.5, 6.55, 12.5, 0.5,
         "Validation:  SOTA on CIFAR-10/100, ImageNet, DVS-CIFAR10  +  generalizes to Transformer SNNs and detection.",
         size=14, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
add_footer(s, 8)


# ─────────────────────────────────────────────────────────────────
# Section divider — Part A: MP-Init
# (NOT counted as a numbered slide; we'll keep S9-S13 for Part A)
# Actually we keep numbering tight, the user said 22 slides total.
# So we put the section header tag inside each Part A slide title.
# Skip a divider slide. Continue with S9.
# ─────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────
# S9. Empirical Evidence
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part A · MP-Init  ▸  Empirical Evidence",
              "The membrane potential drifts — even under TCS-aware normalization")

# 4-panel composite
add_image(s, ASSETS + "/density_plot_membrane_potential_tdbn.png", 0.4, 1.4, w=3.1)
add_text(s, 0.4, 4.1, 3.1, 0.3, "(a) tdBN", size=12, color=GRAY, align=PP_ALIGN.CENTER)
add_image(s, ASSETS + "/density_plot_membrane_potential_tebn.png", 3.7, 1.4, w=3.1)
add_text(s, 3.7, 4.1, 3.1, 0.3, "(b) TEBN", size=12, color=GRAY, align=PP_ALIGN.CENTER)
add_image(s, ASSETS + "/density_plot_membrane_potential_tab.png", 7.0, 1.4, w=3.1)
add_text(s, 7.0, 4.1, 3.1, 0.3, "(c) TAB", size=12, color=GRAY, align=PP_ALIGN.CENTER)
add_image(s, ASSETS + "/density_plot_membrane_potential_mpinit.png", 10.3, 1.4, w=3.1)
add_text(s, 10.3, 4.1, 3.1, 0.3, "(d) MP-Init (ours)", size=12, color=GREEN, align=PP_ALIGN.CENTER, bold=True)

add_text(s, 0.5, 4.7, 12.5, 0.5,
         "Membrane potential distributions across timesteps (ResNet-19, CIFAR-100)",
         size=14, color=GRAY, align=PP_ALIGN.CENTER)

add_box(s, 0.5, 5.4, 12.3, 1.5,
        "Existing TCS-aware normalizations (TEBN, TAB) align activations\nbut leave M[t] drifting.   MP-Init is the first to align M[t] itself.",
        fill=LIGHT, outline=RED, font_size=16, font_color=RED)
add_footer(s, 9)


# ─────────────────────────────────────────────────────────────────
# S10. Markov Chain View
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part A · MP-Init  ▸  Markov-Chain View",
              "Why does M[t] always drift toward a single distribution?")

add_text(s, 0.5, 1.4, 12.5, 0.5,
         "Recurrence is a Markov chain", size=18, bold=True, color=NAVY)
add_box(s, 0.5, 1.9, 12.5, 0.7,
        "U[t+1]  =  f( U[t], I_in[t+1] )       —  i.i.d. inputs ⇒ Markov chain",
        outline=NAVY, font_size=14, font_color=NAVY, bold=False)

add_text(s, 0.5, 2.8, 12.5, 0.5,
         "Theorem (informal)", size=18, bold=True, color=NAVY)
add_box(s, 0.5, 3.3, 12.5, 1.4,
        "Under (i.i.d. inputs) + (bounded U[t]), the chain satisfies\n"
        "Doeblin's minorization condition. Hence it has a unique stationary distribution π,\n"
        "and    ‖ Pⁿ(x, ·) − π(·) ‖_TV  ≤  C · (1 − ε)ⁿ      (geometric convergence)",
        fill=LIGHT, outline=NAVY, font_size=14, font_color=NAVY, bold=False)

add_text(s, 0.5, 4.95, 12.5, 0.5,
         "Implications", size=18, bold=True, color=RED)
add_bullets(s, 0.5, 5.45, 12.5, 1.7,
            ["Initial state ≠ π  ⇒  TCS is mathematically inevitable",
             "Standard practice: U[0]=0 — but π(0) is generally far from zero",
             "→  Fix the initial state, not the normalization layer"],
            size=14)
add_footer(s, 10)


# ─────────────────────────────────────────────────────────────────
# S11. The Idea
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part A · MP-Init  ▸  The Idea",
              "A per-layer constant suffices — and the right constant is E[π]")

# Top: lemma
add_text(s, 0.5, 1.4, 12.5, 0.5,
         "Lemma: best constant is the mean", size=18, bold=True, color=NAVY)
add_box(s, 0.5, 1.9, 12.5, 0.7,
        "argmin_c  E[ (π − c)² ]   =   E[π]",
        fill=LIGHT, outline=NAVY, font_size=14, font_color=NAVY, bold=False)

# Middle: how to estimate E[π]
add_text(s, 0.5, 2.85, 12.5, 0.5,
         "How to estimate E[π] without knowing π", size=18, bold=True, color=NAVY)
add_bullets(s, 0.5, 3.35, 12.5, 1.5,
            ["By the theorem, U[T] is the timestep closest to π",
             "Maintain a running mean μ ← β·μ + (1−β)·E[U[T]]   over training",
             "Use μ as the initial state at every batch"],
            size=14)

# Bottom: visual story
add_box(s, 0.5, 5.0, 12.5, 1.7,
        "Story:    U[0] = 0   ───drift───▸   π        becomes        U[0] = E[π]   ───stable───▸   π",
        fill=LIGHT, outline=GREEN, font_size=15, font_color=GREEN)

add_text(s, 0.5, 6.85, 12.5, 0.4,
         "Cost: one running-mean buffer per layer.",
         size=12, color=GRAY, align=PP_ALIGN.CENTER)
add_footer(s, 11)


# ─────────────────────────────────────────────────────────────────
# S12. MP-Init Algorithm
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part A · MP-Init  ▸  Algorithm",
              "Six lines of pseudocode, training-only update")

# Pseudocode box
code_lines = [
    "Init  μ_l ← 0   for each layer l",
    "for each batch:",
    "   U_l[0] ← μ_l                              ← apply at start",
    "   run T LIF steps as usual",
    "   batch_mean ← mean( U_l[T] of active neurons )",
    "   μ_l ← β · μ_l  +  (1 − β) · batch_mean    ← EMA, β = 0.9",
]
add_text(s, 0.5, 1.4, 12.5, 0.5,
         "MP-Init pseudocode", size=18, bold=True, color=NAVY)

# code monospace box
shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                          Inches(0.5), Inches(2.0),
                          Inches(12.5), Inches(2.6))
shp.fill.solid(); shp.fill.fore_color.rgb = LIGHT
shp.line.color.rgb = NAVY; shp.line.width = Pt(1)
tf = shp.text_frame; tf.word_wrap = True
tf.margin_left = Inches(0.3); tf.margin_top = Inches(0.15)
for i, line in enumerate(code_lines):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = line
    p.font.name = "Consolas"
    p.font.size = Pt(14); p.font.color.rgb = NAVY

# Inference box highlighting
add_box(s, 0.5, 4.9, 12.5, 1.2,
        "Inference:    μ_l is fixed  ⇒  standard LIF dynamics, no extra ops, full HW compatibility.",
        fill=NAVY, outline=NAVY,
        font_size=16, font_color=RGBColor(0xFF,0xFF,0xFF))

add_bullets(s, 0.5, 6.3, 12.5, 1.0,
            ["O(L) extra parameters  (L = layers)",
             "No change to BN, no per-timestep parameters, no extra losses"],
            size=13, color=GRAY)
add_footer(s, 12)


# ─────────────────────────────────────────────────────────────────
# S13. MP-Init Works
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part A · MP-Init  ▸  It Works",
              "Distribution alignment + geometric convergence + consistent gains")

add_image(s, ASSETS + "/density_plot_membrane_potential_mpinit.png", 0.4, 1.4, w=4.0)
add_text(s, 0.4, 4.7, 4.0, 0.3, "Aligned distributions",
         size=11, color=GRAY, align=PP_ALIGN.CENTER)

add_image(s, ASSETS + "/tv_distance.png", 4.6, 1.4, w=4.0)
add_text(s, 4.6, 4.7, 4.0, 0.3, "Geometric TV decay",
         size=11, color=GRAY, align=PP_ALIGN.CENTER)

# Right: Table 6.1 highlights
add_text(s, 8.9, 1.4, 4.2, 0.4,
         "Consistent gains", size=16, bold=True, color=NAVY)
add_box(s, 8.9, 1.9, 4.2, 2.7,
        "CIFAR-100 (T=4)\n"
        "  tdBN          75.55  →  76.09\n"
        "  TEBN          75.96  →  76.45\n"
        "  TAB           76.25  →  77.24\n\n"
        "DVS-CIFAR10 (T=10)\n"
        "  tdBN          76.60  →  77.37\n"
        "  TEBN          75.17  →  76.70",
        fill=LIGHT, outline=NAVY, font_size=12, font_color=NAVY, bold=False)

add_box(s, 0.5, 5.4, 12.5, 1.5,
        "MP-Init helps tdBN, TEBN, and TAB alike.\n"
        "Most pronounced at small T:  T=2  →  +1.66 %pt    (latency-critical regime)",
        fill=LIGHT, outline=GREEN, font_size=15, font_color=GREEN)
add_footer(s, 13)


# ─────────────────────────────────────────────────────────────────
# Part B: TrSG
# ─────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────
# S14. Two SG Families
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part B · TrSG  ▸  Two Surrogate-Gradient Families",
              "AS-SG vs RS-SG — innocent-looking choice with consequences")

# Left box
add_box(s, 0.5, 1.5, 6.0, 0.7,
        "AS-SG: Absolute-Scale", fill=NAVY, outline=NAVY,
        font_size=18, font_color=RGBColor(0xFF,0xFF,0xFF))
add_box(s, 0.5, 2.3, 6.0, 1.0,
        "x  =  M[t]  −  V_thr",
        fill=LIGHT, outline=NAVY, font_size=14, font_color=NAVY, bold=False)
add_bullets(s, 0.6, 3.5, 5.8, 1.8,
            ["Window width fixed at γ — independent of V_thr",
             "Magnitude independent of V_thr",
             "Used by tdBN, TET, PLIF, LTMD, …"],
            size=13)

# Right box
add_box(s, 6.7, 1.5, 6.0, 0.7,
        "RS-SG: Relative-Scale", fill=NAVY, outline=NAVY,
        font_size=18, font_color=RGBColor(0xFF,0xFF,0xFF))
add_box(s, 6.7, 2.3, 6.0, 1.0,
        "x  =  M[t] / V_thr  −  1",
        fill=LIGHT, outline=NAVY, font_size=14, font_color=NAVY, bold=False)
add_bullets(s, 6.8, 3.5, 5.8, 1.8,
            ["Window width scales with V_thr ✓",
             "But magnitude scales as 1 / V_thr ✗",
             "Used by Meng 2023, DIET-SNN"],
            size=13)

add_box(s, 0.5, 5.5, 12.2, 1.4,
        "These two have been treated as essentially equivalent.\n"
        "But once V_thr is trainable, their failure modes diverge sharply.",
        fill=LIGHT, outline=RED, font_size=15, font_color=RED)
add_footer(s, 14)


# ─────────────────────────────────────────────────────────────────
# S15. The Failure Modes — Figure 5.1
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part B · TrSG  ▸  The Failure Modes",
              "Window width × gradient magnitude — both must scale correctly")

# second_main has aspect 0.87 (taller than wide).  Size by height to fit slide.
# Height 4.9 → width ≈ 4.26.  Center it.
add_image(s, ASSETS + "/second_main.png", 4.55, 1.4, h=4.9)

add_text(s, 0.5, 6.4, 12.5, 0.4,
         "Each colored box = active gradient region.   Width = where gradient flows.   Height = gradient magnitude.",
         size=12, color=GRAY, align=PP_ALIGN.CENTER)
add_text(s, 0.5, 6.8, 12.5, 0.4,
         "AS-SG ⇒ flooding / starvation   ·   RS-SG ⇒ explosion / vanishing   ·   TrSG ⇒ both balanced",
         size=14, bold=True, color=RED, align=PP_ALIGN.CENTER)
add_footer(s, 15)


# ─────────────────────────────────────────────────────────────────
# S16. TrSG — One-Line Fix
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part B · TrSG  ▸  One-Line Fix",
              "A forward-path multiplication that exactly cancels the bad factor")

# Center: the trick
add_text(s, 0.5, 1.4, 12.5, 0.5,
         "Keep relative-scale window, multiply threshold on forward path:",
         size=18, bold=True, color=NAVY)

add_box(s, 0.5, 2.0, 12.5, 1.4,
        "x  =  M[t] / V_thr  −  1                                    \n"
        "O[t]  =  V_thr  ·  S[t]                  ← only difference vs RS-SG",
        fill=LIGHT, outline=NAVY, font_size=16, font_color=NAVY, bold=False)

# Chain rule cancellation
add_text(s, 0.5, 3.6, 12.5, 0.5,
         "Chain rule:", size=18, bold=True, color=NAVY)
add_box(s, 0.5, 4.1, 12.5, 1.2,
        "∂O / ∂M  =  V_thr  ·  ∂S / ∂M  ≈  V_thr  ·  (1/V_thr) f'(x)  =  f'(x)",
        fill=LIGHT, outline=GREEN, font_size=18, font_color=GREEN, bold=False)

# Conclusion
add_box(s, 0.5, 5.6, 12.5, 1.4,
        "Window adapts with V_thr   ✓\n"
        "Magnitude is threshold-invariant   ✓",
        fill=GREEN, outline=GREEN, font_size=20, font_color=RGBColor(0xFF,0xFF,0xFF))
add_footer(s, 16)


# ─────────────────────────────────────────────────────────────────
# S17. Inference Stays Binary
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part B · TrSG  ▸  Inference Stays Binary",
              "The trick is training-only — neuromorphic HW never sees it")

add_text(s, 0.5, 1.5, 12.5, 0.5,
         "Training", size=18, bold=True, color=NAVY)
add_box(s, 0.5, 2.0, 12.5, 0.9,
        "Forward:    O[t]  =  V_thr · S[t]              (S ∈ {0,1}, output is real-valued)",
        fill=LIGHT, outline=NAVY, font_size=14, font_color=NAVY, bold=False)

add_text(s, 0.5, 3.2, 12.5, 0.5,
         "Inference", size=18, bold=True, color=NAVY)
add_box(s, 0.5, 3.7, 12.5, 0.9,
        "Absorb V_thr into the next layer:    W'_{l→l+1}  =  V_thr · W_{l→l+1}",
        fill=LIGHT, outline=NAVY, font_size=14, font_color=NAVY, bold=False)
add_box(s, 0.5, 4.7, 12.5, 0.9,
        "Forward at test time:    O[t]  =  S[t] ∈ {0,1}                   (standard LIF)",
        fill=LIGHT, outline=GREEN, font_size=14, font_color=GREEN, bold=False)

add_box(s, 0.5, 6.0, 12.5, 1.0,
        "Net effect: TrSG is purely a training-time optimization aid.\n"
        "Deployment uses identical operators as ordinary LIF SNNs.",
        fill=GREEN, outline=GREEN, font_size=15, font_color=RGBColor(0xFF,0xFF,0xFF))
add_footer(s, 17)


# ─────────────────────────────────────────────────────────────────
# S18. Stress Tests
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part B · TrSG  ▸  Stress Tests",
              "What happens at extreme thresholds during real training?")

# Tables
add_text(s, 0.5, 1.3, 6.0, 0.4,
         "V_thr = 0.1  (low)", size=14, bold=True, color=NAVY)
add_box(s, 0.5, 1.7, 6.0, 1.5,
        "                    AbsStr      RatioAG    GradCV\n"
        "AS-SG (ep100)    0.0145    90.20%      1.22 \n"
        "RS-SG (ep100)         NaN       0%          NaN \n"
        "TrSG  (ep100)    0.0153    21.45%      0.87",
        fill=LIGHT, outline=GRAY, font_size=11, font_color=NAVY, bold=False)

add_text(s, 6.7, 1.3, 6.0, 0.4,
         "V_thr = 2.0  (high)", size=14, bold=True, color=NAVY)
add_box(s, 6.7, 1.7, 6.0, 1.5,
        "                    AbsStr      RatioAG    GradCV\n"
        "AS-SG (ep100)    0.0000     0.00%       0.00 \n"
        "RS-SG (ep100)    0.0021     0.10%       2.25 \n"
        "TrSG  (ep100)    0.0323     6.90%       0.73",
        fill=LIGHT, outline=GRAY, font_size=11, font_color=NAVY, bold=False)

# Threshold trajectories
add_text(s, 0.5, 3.45, 12.5, 0.4,
         "Real trainings actually visit these regions",
         size=14, bold=True, color=RED)
add_image(s, ASSETS + "/threshold_training_cifar100.png", 0.5, 3.85, w=6.0, h=2.6)
add_image(s, ASSETS + "/threshold_training_imagenet.png", 6.8, 3.85, w=6.0, h=2.6)
add_text(s, 0.5, 6.5, 6.0, 0.3, "CIFAR-100", size=10, color=GRAY, align=PP_ALIGN.CENTER)
add_text(s, 6.8, 6.5, 6.0, 0.3, "ImageNet", size=10, color=GRAY, align=PP_ALIGN.CENTER)

add_text(s, 0.5, 6.85, 12.5, 0.3,
         "→ Threshold-robustness is a practical necessity, not a corner case.",
         size=13, bold=True, color=RED, align=PP_ALIGN.CENTER)
add_footer(s, 18)


# ─────────────────────────────────────────────────────────────────
# S19. Main Results
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part C · Experiments  ▸  Main Results",
              "State-of-the-art on every benchmark we evaluate")

# Big results table
add_box(s, 0.5, 1.5, 12.5, 0.5,
        "Static and dynamic image datasets — standard LIF, no architectural changes",
        outline=NAVY, font_size=12, font_color=NAVY, bold=False)

# Three cards
def result_card(slide, x, y, w, h, name, prev_best, ours, color):
    add_box(slide, x, y, w, 0.7, name,
            fill=color, outline=color, font_size=18,
            font_color=RGBColor(0xFF,0xFF,0xFF))
    add_text(slide, x, y+0.8, w, 0.4,
             "Previous best", size=12, color=GRAY, align=PP_ALIGN.CENTER)
    add_text(slide, x, y+1.2, w, 0.5,
             prev_best, size=20, bold=True, color=GRAY, align=PP_ALIGN.CENTER)
    add_text(slide, x, y+1.85, w, 0.4,
             "Ours", size=12, color=color, align=PP_ALIGN.CENTER)
    add_text(slide, x, y+2.25, w, 0.6,
             ours, size=24, bold=True, color=color, align=PP_ALIGN.CENTER)

result_card(s, 0.5, 2.3, 4.0, 2.9, "CIFAR-100", "76.82", "77.91", NAVY)
result_card(s, 4.7, 2.3, 4.0, 2.9, "ImageNet (SEW-R34)", "69.03", "69.67", RED)
result_card(s, 8.9, 2.3, 4.0, 2.9, "DVS-CIFAR10", "76.20", "81.43", GREEN)

add_text(s, 0.5, 5.4, 12.5, 0.4,
         "All entries: top-1 accuracy (%).   Baselines: TAB / MPS / RMP-Loss respectively.",
         size=12, color=GRAY, align=PP_ALIGN.CENTER)

add_box(s, 0.5, 6.0, 12.5, 1.0,
        "DVS-CIFAR10 +5.23 %pt — biggest gain on the most challenging (event-based) dataset.",
        fill=LIGHT, outline=GREEN, font_size=15, font_color=GREEN)
add_footer(s, 19)


# ─────────────────────────────────────────────────────────────────
# S20. Ablation & Generalization
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part C · Experiments  ▸  Ablation & Generalization",
              "Independent + complementary;  classification-only is not the full story")

# Left: ablation table
add_text(s, 0.5, 1.4, 6.0, 0.4,
         "Ablation (DVS-CIFAR10, T=10)", size=14, bold=True, color=NAVY)
add_box(s, 0.5, 1.85, 6.0, 2.5,
        "MP-Init   TrSG       Acc (%)\n"
        "   ✗         ✗            76.60\n"
        "   ✓         ✗            77.37    +0.77\n"
        "   ✗         ✓            80.83    +4.23\n"
        "   ✓         ✓            81.43    +4.83",
        fill=LIGHT, outline=NAVY, font_size=13, font_color=NAVY, bold=False)

# Right: generalization
add_text(s, 7.0, 1.4, 6.0, 0.4,
         "Beyond classification", size=14, bold=True, color=NAVY)
add_box(s, 7.0, 1.85, 6.0, 1.1,
        "Transformer SNN (QKFormer, ImageNet T=4)\n"
        "  baseline  78.80    →    + ours  79.90    (+1.10)",
        fill=LIGHT, outline=RED, font_size=12, font_color=RED, bold=False)
add_box(s, 7.0, 3.05, 6.0, 1.3,
        "Object detection (COCO-2017, EMS-ResNet10 T=3)\n"
        "  mAP@0.5       0.291  →  0.305    (+0.014)\n"
        "  mAP@0.5:0.95  0.138  →  0.147    (+0.009)",
        fill=LIGHT, outline=RED, font_size=12, font_color=RED, bold=False)

add_box(s, 0.5, 4.6, 12.5, 0.9,
        "MP-Init and TrSG are independently helpful, and together fully complementary.",
        fill=LIGHT, outline=GREEN, font_size=15, font_color=GREEN)

add_text(s, 0.5, 5.7, 12.5, 1.4,
         "Drop-in replacements: no architecture changes, no extra training pipeline,\n"
         "no new losses, no per-timestep parameters.",
         size=14, color=GRAY, align=PP_ALIGN.CENTER)
add_footer(s, 20)


# ─────────────────────────────────────────────────────────────────
# S21. Why It Works
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Part C · Experiments  ▸  Why It Works",
              "Mechanistic evidence beyond accuracy numbers")

# Left: gradient cosine sim
add_image(s, ASSETS + "/grad_collision.png", 0.5, 1.5, w=6.0)
add_text(s, 0.5, 4.6, 6.0, 0.4,
         "Gradient cosine similarity across timesteps",
         size=12, color=GRAY, align=PP_ALIGN.CENTER, bold=True)
add_bullets(s, 0.5, 5.0, 6.0, 1.6,
            ["Without MP-Init: gradients at t=1 fight later steps",
             "With MP-Init: temporal ensemble realigned",
             "Stable optimization — not just stable inference"],
            size=12)

# Right: pareto frontier
add_image(s, ASSETS + "/firing_rate_vs_accuracy.png", 7.0, 1.5, w=6.0)
add_text(s, 7.0, 4.6, 6.0, 0.4,
         "Accuracy vs firing rate",
         size=12, color=GRAY, align=PP_ALIGN.CENTER, bold=True)
add_bullets(s, 7.0, 5.0, 6.0, 1.6,
            ["MP-Init dominates baseline at every firing rate",
             "Even at 0.06 firing rate, beats all baselines",
             "Gains are not from spiking more — from spiking better"],
            size=12)

add_text(s, 0.5, 6.85, 12.5, 0.4,
         "→ Efficiency-preserving improvements, not accuracy-via-spend.",
         size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
add_footer(s, 21)


# ─────────────────────────────────────────────────────────────────
# S22. Conclusion + Q&A
# ─────────────────────────────────────────────────────────────────
s = add_blank_slide()
add_title_bar(s, "Conclusion",
              "Two diagnoses, two minimal fixes, broad gains")

# One-line summary
add_box(s, 0.5, 1.5, 12.5, 1.4,
        "TCS is a problem of the membrane potential — fixed at the initial state.\n"
        "Threshold-invariance is achieved by one forward-path multiplication.",
        fill=NAVY, outline=NAVY, font_size=18,
        font_color=RGBColor(0xFF,0xFF,0xFF))

# Three columns
add_text(s, 0.5, 3.1, 4.0, 0.4, "Contributions",
         size=16, bold=True, color=NAVY)
add_bullets(s, 0.5, 3.5, 4.0, 3.0,
            ["MP-Init: align M[0] with stationary mean",
             "TrSG: cancel 1/V_thr factor exactly",
             "SOTA on 4 datasets",
             "Generalizes to Transformer SNN, detection"],
            size=12)

add_text(s, 4.7, 3.1, 4.0, 0.4, "Limitations",
         size=16, bold=True, color=ORANGE)
add_bullets(s, 4.7, 3.5, 4.0, 3.0,
            ["Per-layer constant — richer parametrization possible",
             "Trainable LIF params may need quantization for fixed-param HW",
             "i.i.d. assumption needs untrained-network analysis"],
            size=12, color=ORANGE)

add_text(s, 8.9, 3.1, 4.0, 0.4, "Future Work",
         size=16, bold=True, color=GREEN)
add_bullets(s, 8.9, 3.5, 4.0, 3.0,
            ["Hardware-aware adaptation (Loihi, TrueNorth)",
             "Beyond LIF: adaptive LIF, Izhikevich",
             "Foundation-scale SNNs"],
            size=12, color=GREEN)

# Thanks
add_text(s, 0.5, 6.6, 12.5, 0.5,
         "Thank you.   Questions?",
         size=24, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
add_footer(s, 22)


# Save
out = "Hyunho_Kook_Master_Defense.pptx"
prs.save(out)
print(f"Saved: {out}  ({prs.slide_width}x{prs.slide_height} EMU = 16:9)")
print(f"Slides: {len(prs.slides)}")
