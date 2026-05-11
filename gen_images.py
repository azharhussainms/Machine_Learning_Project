#!/usr/bin/env python3
"""Generate professional AI-themed illustrations, diagrams, flowcharts, charts and Gantt
charts for the three Machine Learning project proposals."""
import os, math, random
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Circle, Rectangle
from matplotlib.lines import Line2D
import matplotlib.font_manager as fm

random.seed(7)
np.random.seed(7)

A = "assets"
os.makedirs(A, exist_ok=True)

# ----------------------------------------------------------------------------
# Theme palettes for the three proposals
# ----------------------------------------------------------------------------
THEMES = {
    1: dict(name="Crowd Counting",
            dark="#0B132B", mid="#1C2541", accent="#3A86FF", accent2="#00B4D8",
            light="#5BC0BE", paper="#EDF6F9", text="#0B132B"),
    2: dict(name="Handwritten Text Recognition",
            dark="#2B0B3F", mid="#3D1A52", accent="#7B2CBF", accent2="#9D4EDD",
            light="#C77DFF", paper="#F6F1FA", text="#2B0B3F"),
    3: dict(name="Sign Language Recognition",
            dark="#06281F", mid="#0B3D2E", accent="#1B998B", accent2="#2D9B5C",
            light="#7DDF9E", paper="#EFF7F1", text="#06281F"),
}

def lerp(c1, c2, t):
    c1 = np.array(matplotlib.colors.to_rgb(c1))
    c2 = np.array(matplotlib.colors.to_rgb(c2))
    return tuple(c1 + (c2 - c1) * t)

def neural_motif(ax, x0, y0, w, h, layers, color, alpha=0.9, node_s=70, lw=0.8):
    """Draw a small neural-network motif inside the box [x0,x0+w]x[y0,y0+h]."""
    xs = np.linspace(x0, x0 + w, len(layers))
    pts = []
    for li, n in enumerate(layers):
        ys = np.linspace(y0 + h * 0.12, y0 + h * 0.88, n) if n > 1 else [y0 + h / 2]
        pts.append([(xs[li], y) for y in ys])
    for li in range(len(layers) - 1):
        for (x1, y1) in pts[li]:
            for (x2, y2) in pts[li + 1]:
                ax.plot([x1, x2], [y1, y2], color=color, lw=lw, alpha=alpha * 0.35, zorder=1)
    for layer in pts:
        for (x, y) in layer:
            ax.scatter([x], [y], s=node_s, color=color, alpha=alpha, zorder=2,
                       edgecolors="white", linewidths=0.6)

# ============================================================================
# 1. COVER BANNERS  (one per proposal, different visual theme)
# ============================================================================
def cover_banner(pid):
    th = THEMES[pid]
    fig = plt.figure(figsize=(8.27, 3.6), dpi=200)
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis("off")
    # gradient background
    grad = np.zeros((256, 256, 3))
    for i in range(256):
        for j in range(256):
            t = (i / 255 * 0.6 + j / 255 * 0.4)
            grad[i, j] = lerp(th["dark"], th["mid"], t)
    ax.imshow(grad, extent=[0, 10, 0, 4], aspect="auto", zorder=0)

    # decorative concentric / data motifs depending on theme
    if pid == 1:  # crowd -> dots forming a density cloud + heatmap-ish blobs
        for _ in range(420):
            cx, cy = np.random.normal(7.4, 1.5), np.random.normal(2.0, 0.9)
            ax.scatter([cx], [cy], s=np.random.uniform(4, 26),
                       color=lerp(th["accent2"], th["light"], np.random.rand()),
                       alpha=np.random.uniform(0.15, 0.7), zorder=1)
        for _ in range(7):
            cx, cy = np.random.uniform(5.5, 9.5), np.random.uniform(0.6, 3.4)
            ax.add_patch(Circle((cx, cy), np.random.uniform(0.3, 0.9),
                                color=th["accent"], alpha=0.10, zorder=1))
    elif pid == 2:  # handwriting -> flowing "ink" strokes + grid
        for gx in np.arange(5.2, 10, 0.5):
            ax.plot([gx, gx], [0.2, 3.8], color="white", alpha=0.05, lw=0.8, zorder=1)
        t = np.linspace(0, 1, 300)
        for k in range(6):
            base = 0.6 + k * 0.55
            x = 5.4 + t * 4.2
            y = base + 0.35 * np.sin(2 * np.pi * (2 + k) * t) * np.exp(-1.5 * t) + 0.1 * np.sin(20 * t)
            ax.plot(x, y, color=lerp(th["light"], th["accent2"], k / 6),
                    lw=2.4, alpha=0.7, solid_capstyle="round", zorder=2)
    else:  # sign language -> hand landmark skeleton motif
        # stylised hand: palm point + 5 fingers of joints
        px, py = 7.6, 1.0
        finger_dirs = [(-0.55, 1.0), (-0.18, 1.25), (0.12, 1.3), (0.42, 1.18), (0.72, 0.7)]
        ax.scatter([px], [py], s=120, color=th["light"], zorder=3, edgecolors="white")
        for (dx, dy) in finger_dirs:
            joints = [(px, py)]
            for s in (0.5, 0.85, 1.15):
                joints.append((px + dx * s, py + dy * s))
            xs = [j[0] for j in joints]; ys = [j[1] for j in joints]
            ax.plot(xs, ys, color=th["accent2"], lw=2.2, alpha=0.85, zorder=2)
            ax.scatter(xs[1:], ys[1:], s=55, color=th["light"], zorder=3, edgecolors="white", linewidths=0.5)
        for _ in range(120):
            ax.scatter([np.random.uniform(5.4, 9.6)], [np.random.uniform(0.4, 3.6)],
                       s=np.random.uniform(2, 10), color="white", alpha=np.random.uniform(0.05, 0.25), zorder=1)

    # left-side neural net motif (common element, AI theme)
    neural_motif(ax, 0.5, 0.7, 2.6, 2.6, [4, 6, 5, 3], th["accent2"], alpha=0.85, node_s=55, lw=0.7)

    # circuit lines
    for _ in range(14):
        x = np.random.uniform(0.2, 4.5); y = np.random.uniform(0.2, 3.8)
        ax.plot([x, x + np.random.uniform(0.3, 1.0)], [y, y], color=th["accent"], alpha=0.18, lw=1.0, zorder=1)

    # title text
    titles = {1: ("AI-BASED CROWD", "COUNTING SYSTEM"),
              2: ("AI HANDWRITTEN TEXT", "RECOGNITION SYSTEM"),
              3: ("SIGN LANGUAGE", "RECOGNITION SYSTEM")}
    subs = {1: "Deep Learning  •  Density Estimation  •  Smart Surveillance",
            2: "Computer Vision  •  CRNN + CTC  •  Document Digitisation",
            3: "MediaPipe  •  Gesture Recognition  •  Assistive Technology"}
    ax.text(3.45, 2.55, titles[pid][0], color="white", fontsize=21, fontweight="bold",
            family="DejaVu Sans", zorder=5)
    ax.text(3.45, 1.95, titles[pid][1], color=th["light"], fontsize=21, fontweight="bold",
            family="DejaVu Sans", zorder=5)
    ax.plot([3.5, 8.6], [1.6, 1.6], color=th["accent"], lw=2.5, zorder=5)
    ax.text(3.5, 1.2, subs[pid], color="#D7E3EA", fontsize=10.5, family="DejaVu Sans", zorder=5)
    ax.text(3.5, 0.6, "Riphah International University  |  RISE MSAI  |  Machine Learning",
            color="#9FB3C0", fontsize=8.5, family="DejaVu Sans", zorder=5)
    ax.set_xlim(0, 10); ax.set_ylim(0, 4)
    fig.savefig(f"{A}/cover_{pid}.png", dpi=200, bbox_inches="tight", pad_inches=0)
    plt.close(fig)

# ============================================================================
# 2. SECTION DIVIDER BANNERS  (thin decorative band per proposal)
# ============================================================================
def divider(pid):
    th = THEMES[pid]
    fig = plt.figure(figsize=(8.27, 0.55), dpi=200)
    ax = fig.add_axes([0, 0, 1, 1]); ax.axis("off")
    grad = np.zeros((10, 256, 3))
    for j in range(256):
        grad[:, j] = lerp(th["accent"], th["mid"], j / 255)
    ax.imshow(grad, extent=[0, 10, 0, 1], aspect="auto")
    for _ in range(60):
        ax.scatter([np.random.uniform(0, 10)], [np.random.uniform(0, 1)],
                   s=np.random.uniform(2, 14), color="white", alpha=np.random.uniform(0.05, 0.3))
    neural_motif(ax, 0.1, 0.05, 1.4, 0.9, [3, 4, 3], "white", alpha=0.5, node_s=12, lw=0.4)
    neural_motif(ax, 8.5, 0.05, 1.4, 0.9, [3, 4, 3], "white", alpha=0.5, node_s=12, lw=0.4)
    ax.set_xlim(0, 10); ax.set_ylim(0, 1)
    fig.savefig(f"{A}/divider_{pid}.png", dpi=200, bbox_inches="tight", pad_inches=0)
    plt.close(fig)

# ============================================================================
# 3. ARCHITECTURE DIAGRAMS
# ============================================================================
def box(ax, x, y, w, h, text, fc, ec, tc="white", fs=9, bold=True, round=0.06):
    p = FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0.02,rounding_size={round}",
                       fc=fc, ec=ec, lw=1.4, zorder=3)
    ax.add_patch(p)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", color=tc,
            fontsize=fs, fontweight="bold" if bold else "normal", zorder=4, wrap=True)

def arrow(ax, x1, y1, x2, y2, color="#444", lw=1.6, style="-|>", rad=0.0):
    a = FancyArrowPatch((x1, y1), (x2, y2), arrowstyle=style, mutation_scale=14,
                        color=color, lw=lw, zorder=2,
                        connectionstyle=f"arc3,rad={rad}")
    ax.add_patch(a)

def architecture(pid):
    th = THEMES[pid]
    fig, ax = plt.subplots(figsize=(9.2, 5.2), dpi=200)
    ax.set_xlim(0, 15.4); ax.set_ylim(0, 9); ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.text(7.7, 8.55, f"System Architecture — {THEMES[pid]['name']}",
            ha="center", fontsize=13, fontweight="bold", color=th["dark"])

    BW, BH, GAP = 2.18, 1.8, 0.30           # top-row box width/height/gap
    TY = 5.7                                # top-row y
    FS = 7.3
    DW, DH, DY = 2.6, 1.7, 2.6              # downstream box dims / y
    APPW, APPH, APPY = 9.4, 1.3, 0.5        # application bar

    def top_row(rows, colors):
        x = 0.5
        cx_list = []
        for i, t in enumerate(rows):
            box(ax, x, TY, BW, BH, t, colors[i], th["dark"], fs=FS)
            cx_list.append(x + BW / 2)
            if i < len(rows) - 1:
                arrow(ax, x + BW, TY + BH / 2, x + BW + GAP, TY + BH / 2, th["dark"])
            x += BW + GAP
        return cx_list, x - GAP - BW / 2     # centers, last-box center x

    if pid == 1:
        rows = ["CCTV /\nVideo Feed", "Frame Sample\n& Resize", "Pre-process\n(normalise,\ndenoise)",
                "CSRNet\nfront-end\n(VGG-16)", "Dilated-conv\nback-end →\ndensity map", "Integrate map\n→ head count"]
        cols = [th["mid"], th["accent"], th["accent"], th["accent2"], th["accent2"], th["light"]]
        _, lastcx = top_row(rows, cols)
        # downstream row (right -> left)
        box(ax, 10.6, DY, DW, DH, "Heatmap\ngeneration", th["accent2"], th["dark"], fs=FS)
        box(ax, 6.9, DY, DW, DH, "Overcrowding\nthreshold &\nalert engine", th["accent"], th["dark"], fs=FS)
        box(ax, 3.2, DY, DW, DH, "Time-series\nanalytics", th["light"], th["dark"], fs=FS)
        box(ax, (15.4 - APPW) / 2, APPY, APPW, APPH, "Real-time Web Dashboard  (Streamlit / Flask)",
            th["dark"], th["dark"], fs=8.5)
        arrow(ax, lastcx, TY, 11.9, DY + DH, th["dark"], rad=0.0)
        arrow(ax, 10.6, DY + DH / 2, 9.5, DY + DH / 2, th["dark"])
        arrow(ax, 6.9, DY + DH / 2, 5.8, DY + DH / 2, th["dark"])
        arrow(ax, 8.15, DY, 8.6, APPY + APPH, th["dark"])
        arrow(ax, 4.45, DY, 6.0, APPY + APPH, th["dark"])
        arrow(ax, 11.9, DY, 9.4, APPY + APPH, th["dark"], rad=-0.15)
    elif pid == 2:
        rows = ["Input image /\nscan / PDF", "Binarise +\ndeskew +\ndenoise", "Line & word\nsegmentation",
                "CNN feature\nextractor\n(MobileNet)", "Bi-LSTM\nsequence\nencoder", "CTC decoder\n+ beam\nsearch"]
        cols = [th["mid"], th["accent"], th["accent"], th["accent2"], th["accent2"], th["light"]]
        _, lastcx = top_row(rows, cols)
        box(ax, 10.6, DY, DW, DH, "Language model /\nspell correction\n(SymSpell)", th["accent2"], th["dark"], fs=FS)
        box(ax, 6.9, DY, DW, DH, "Layout\nreconstruction", th["accent"], th["dark"], fs=FS)
        box(ax, 3.2, DY, DW, DH, "Multi-language\nswitch\n(En / Ur)", th["light"], th["dark"], fs=FS)
        box(ax, (15.4 - APPW) / 2, APPY, APPW, APPH, "Editable Output  →  TXT / DOCX / Searchable PDF",
            th["dark"], th["dark"], fs=8.5)
        arrow(ax, lastcx, TY, 11.9, DY + DH, th["dark"])
        arrow(ax, 10.6, DY + DH / 2, 9.5, DY + DH / 2, th["dark"])
        arrow(ax, 6.9, DY + DH / 2, 5.8, DY + DH / 2, th["dark"])
        arrow(ax, 4.45, DY, 6.6, APPY + APPH, th["dark"])
        arrow(ax, 8.15, DY, 8.8, APPY + APPH, th["dark"])
    else:
        rows = ["Webcam\nvideo stream", "MediaPipe\nHands → 21 ×\n3-D landmarks", "Landmark\nnormalise\n(scale/origin)",
                "Feature vector\n/ frame buffer", "Static classifier\nMLP / 1-D CNN", "Dynamic head\nLSTM / GRU"]
        cols = [th["mid"], th["accent"], th["accent"], th["accent2"], th["accent2"], th["light"]]
        _, lastcx = top_row(rows, cols)
        box(ax, 10.6, DY, DW, DH, "Letter / word\nprediction +\nsmoothing", th["accent2"], th["dark"], fs=FS)
        box(ax, 6.9, DY, DW, DH, "Sentence builder\n& auto-complete", th["accent"], th["dark"], fs=FS)
        box(ax, 3.2, DY, DW, DH, "Text-to-speech\n(pyttsx3 / gTTS)", th["light"], th["dark"], fs=FS)
        box(ax, (15.4 - APPW) / 2, APPY, APPW, APPH, "On-screen Text  +  Voice Output  (real-time UI)",
            th["dark"], th["dark"], fs=8.5)
        arrow(ax, lastcx, TY, 11.9, DY + DH, th["dark"])
        arrow(ax, 10.6, DY + DH / 2, 9.5, DY + DH / 2, th["dark"])
        arrow(ax, 6.9, DY + DH / 2, 5.8, DY + DH / 2, th["dark"])
        arrow(ax, 4.45, DY, 6.6, APPY + APPH, th["dark"])
        arrow(ax, 8.15, DY, 8.8, APPY + APPH, th["dark"])

    ax.text(0.5, 0.12, "Pipeline stages flow left → right (top) then converge into the application layer (bottom).",
            fontsize=8, color="#555", style="italic")
    fig.savefig(f"{A}/arch_{pid}.png", dpi=200, bbox_inches="tight")
    plt.close(fig)

# ============================================================================
# 4. WORKFLOW FLOWCHARTS  (vertical, decision diamonds)
# ============================================================================
def flowchart(pid):
    th = THEMES[pid]
    fig, ax = plt.subplots(figsize=(6.6, 9.6), dpi=200)
    ax.set_xlim(0, 10); ax.set_ylim(0, 30); ax.axis("off")
    fig.patch.set_facecolor("white")
    ax.text(5, 29.2, f"Project Workflow — {THEMES[pid]['name']}", ha="center",
            fontsize=12.5, fontweight="bold", color=th["dark"])

    def term(y, t, c):  # rounded terminal
        box(ax, 3.0, y, 4.0, 1.6, t, c, th["dark"], fs=9, round=0.5)
    def proc(y, t, c, fs=8.7):
        box(ax, 2.4, y, 5.2, 1.7, t, c, th["dark"], fs=fs, round=0.06)
    def dec(y, t):
        d = mpatches.FancyBboxPatch  # use a diamond via Polygon
        cx, cy = 5.0, y + 0.9
        diamond = mpatches.Polygon([(cx, cy + 1.2), (cx + 2.7, cy), (cx, cy - 1.2), (cx - 2.7, cy)],
                                   closed=True, fc=th["light"], ec=th["dark"], lw=1.4, zorder=3)
        ax.add_patch(diamond)
        ax.text(cx, cy, t, ha="center", va="center", fontsize=8.2, fontweight="bold", zorder=4)
    def dn(y1, y2, lbl=None):
        arrow(ax, 5.0, y1, 5.0, y2, th["dark"])
        if lbl:
            ax.text(5.25, (y1 + y2) / 2, lbl, fontsize=7.6, color="#444")

    common_top = [
        ("Start", th["mid"], "term"),
        ("Literature review &\ndataset acquisition", th["accent"], "proc"),
        ("Data cleaning, labelling &\nexploratory analysis", th["accent"], "proc"),
    ]
    mids = {
        1: [("Generate ground-truth density\nmaps (Gaussian kernels)", th["accent2"], "proc"),
            ("Build CSRNet (transfer learning,\nfrozen VGG-16 front-end)", th["accent2"], "proc"),
            ("Train on ShanghaiTech /\nMall subset (Colab GPU)", th["accent2"], "proc")],
        2: [("Pre-process: deskew, binarise,\nsegment lines & words", th["accent2"], "proc"),
            ("Build CRNN (CNN + BiLSTM + CTC),\ninit from pretrained weights", th["accent2"], "proc"),
            ("Train on IAM / synthetic data\nwith augmentation", th["accent2"], "proc")],
        3: [("Extract MediaPipe landmarks,\nnormalise & build feature set", th["accent2"], "proc"),
            ("Train lightweight classifier\n(MLP / 1-D CNN) + LSTM", th["accent2"], "proc"),
            ("Validate on ASL alphabet /\nWLASL subset", th["accent2"], "proc")],
    }
    dec_q = {1: "MAE / MSE\nacceptable?", 2: "CER / WER\nacceptable?", 3: "Accuracy ≥\ntarget?"}
    tune = {1: "Tune kernels, augment,\nadjust dilation rates",
            2: "More augmentation, tune\nLSTM, beam width, LM",
            3: "Add data, tune window size,\nregularise, balance classes"}
    integ = {1: "Integrate heatmap, alerts &\nanalytics dashboard",
             2: "Add post-processing, layout\nrebuild & export module",
             3: "Add sentence builder &\ntext-to-speech module"}

    y = 26.8
    # start
    term(y, "Start", th["mid"]); dn(y, y - 0.6 + 0);
    # We'll lay items with fixed spacing
    items = []
    items.append(("term", "Start", th["mid"]))
    for (t, c, k) in common_top[1:]:
        items.append(("proc", t, c))
    for (t, c, k) in mids[pid]:
        items.append(("proc", t, c))
    items.append(("dec", dec_q[pid], None))
    items.append(("proc", integ[pid], th["accent"]))
    items.append(("proc", "System testing, user evaluation\n& documentation", th["light"]))
    items.append(("term", "End / Final Demo", th["mid"]))

    # redraw cleanly
    ax.clear(); ax.set_xlim(0, 10); ax.set_ylim(0, 32); ax.axis("off")
    ax.text(5, 31.2, f"Project Workflow", ha="center", fontsize=13, fontweight="bold", color=th["dark"])
    ax.text(5, 30.4, THEMES[pid]['name'], ha="center", fontsize=10.5, color=th["accent"], style="italic")
    y = 29.0
    centers = []
    for (kind, t, c) in items:
        if kind == "term":
            box(ax, 3.0, y - 1.4, 4.0, 1.4, t, c, th["dark"], fs=9.5, round=0.5); h = 1.4
        elif kind == "proc":
            box(ax, 2.2, y - 1.7, 5.6, 1.7, t, c, th["dark"], fs=8.6, round=0.06); h = 1.7
        else:  # dec
            cx, cy = 5.0, y - 1.3
            diamond = mpatches.Polygon([(cx, cy + 1.3), (cx + 2.9, cy), (cx, cy - 1.3), (cx - 2.9, cy)],
                                       closed=True, fc=th["light"], ec=th["dark"], lw=1.4, zorder=3)
            ax.add_patch(diamond)
            ax.text(cx, cy, t, ha="center", va="center", fontsize=8.4, fontweight="bold", zorder=4); h = 2.6
        centers.append((y, y - h))
        y -= (h + 1.0)
    # arrows
    for i in range(len(items) - 1):
        y1 = centers[i][1]; y2 = centers[i + 1][0]
        arrow(ax, 5.0, y1, 5.0, y2, th["dark"])
        if items[i][0] == "dec":
            ax.text(5.2, (y1 + y2) / 2 + 0.1, "Yes", fontsize=8, color="#1a7f37", fontweight="bold")
    # feedback loop on the decision
    dec_idx = [i for i, it in enumerate(items) if it[0] == "dec"][0]
    dy_top = centers[dec_idx][0]; dy_bot = centers[dec_idx][1]
    midy = (dy_top + dy_bot) / 2
    # 'No' goes right and loops up to the 'build/train' box (two boxes above the decision)
    train_idx = dec_idx - 1
    ty = (centers[train_idx][0] + centers[train_idx][1]) / 2
    ax.annotate("", xy=(7.9, ty), xytext=(7.9, midy),
                arrowprops=dict(arrowstyle="-", color=th["accent"], lw=1.6))
    ax.annotate("", xy=(7.9, midy), xytext=(8.0, midy),
                arrowprops=dict(arrowstyle="-", color=th["accent"], lw=1.6))
    arrow(ax, 8.0, midy, 8.0, midy, th["accent"])
    # cleaner: draw explicit polyline
    ax.plot([8.0, 9.2, 9.2, 7.85], [midy, midy, ty, ty], color=th["accent"], lw=1.7, zorder=2)
    arrow(ax, 8.0, ty, 7.8, ty, th["accent"])
    ax.text(9.0, (midy + ty) / 2, "No → refine\n& retrain", fontsize=7.8, color=th["accent"], rotation=90,
            ha="center", va="center")
    ax.plot([7.9, 8.0], [midy, midy], color=th["accent"], lw=1.7)
    fig.savefig(f"{A}/flow_{pid}.png", dpi=200, bbox_inches="tight")
    plt.close(fig)

# ============================================================================
# 5. GANTT CHARTS  (16-week schedule)
# ============================================================================
def gantt(pid):
    th = THEMES[pid]
    tasks_common = [
        ("Requirement analysis & literature review", 1, 2),
        ("Dataset collection & exploration", 2, 2),
        ("Data pre-processing & annotation pipeline", 3, 2),
        ("Baseline model implementation", 5, 2),
    ]
    tasks_specific = {
        1: [("CSRNet / density-map model training", 6, 3),
            ("Heatmap & overcrowding alert module", 9, 2),
            ("Analytics dashboard development", 10, 3),
            ("Integration & real-time optimisation", 12, 2)],
        2: [("CRNN + CTC training & tuning", 6, 3),
            ("Multi-language model & spell-correction", 9, 2),
            ("Layout reconstruction & export module", 10, 3),
            ("Integration & UI development", 12, 2)],
        3: [("Landmark feature engineering & classifier", 6, 3),
            ("Dynamic-gesture LSTM module", 9, 2),
            ("Sentence builder & text-to-speech", 10, 3),
            ("Real-time UI integration & optimisation", 12, 2)],
    }
    tasks_tail = [
        ("Testing, evaluation & error analysis", 13, 2),
        ("Documentation & report writing", 14, 2),
        ("Final review, demo & submission", 16, 1),
    ]
    tasks = tasks_common + tasks_specific[pid] + tasks_tail
    fig, ax = plt.subplots(figsize=(10.0, 5.4), dpi=200)
    fig.patch.set_facecolor("white")
    cmap_cols = [lerp(th["accent"], th["accent2"], i / max(1, len(tasks) - 1)) for i in range(len(tasks))]
    LEFT = -7.5
    for i, (name, start, dur) in enumerate(tasks):
        y = len(tasks) - i - 1
        ax.barh(y, dur, left=start, height=0.55, color=cmap_cols[i], edgecolor=th["dark"], lw=0.8, zorder=3)
        ax.text(start + dur + 0.15, y, f"W{start}–W{start+dur-1}", va="center", fontsize=7.2, color="#555")
        ax.text(LEFT + 0.2, y, name, va="center", ha="left", fontsize=8.0, color=th["dark"])
    ax.set_xlim(LEFT, 18.8); ax.set_ylim(-0.7, len(tasks) + 0.2)
    ax.set_yticks([])
    ax.set_xticks(range(1, 17)); ax.set_xticklabels([f"W{i}" for i in range(1, 17)], fontsize=7.2)
    ax.set_xlabel("Project Week (≈ 16-week semester)", fontsize=9, color=th["dark"])
    ax.set_title(f"Project Timeline / Gantt Chart — {THEMES[pid]['name']}", fontsize=12.5, fontweight="bold", color=th["dark"], pad=14)
    for sp in ["top", "right", "left"]:
        ax.spines[sp].set_visible(False)
    ax.grid(axis="x", color="#ccc", lw=0.5, zorder=0)
    ax.axvline(0.5, color="#999", lw=1.0)
    # phase shading
    ax.axvspan(0.5, 4.99, color=th["light"], alpha=0.12, zorder=0)
    ax.axvspan(5, 12.99, color=th["accent2"], alpha=0.10, zorder=0)
    ax.axvspan(13, 16.5, color=th["light"], alpha=0.12, zorder=0)
    ax.text(2.7, len(tasks) - 0.2, "Phase 1: Foundation", fontsize=7.6, color=th["mid"], ha="center", style="italic")
    ax.text(8.7, len(tasks) - 0.2, "Phase 2: Modelling & Development", fontsize=7.6, color=th["mid"], ha="center", style="italic")
    ax.text(14.7, len(tasks) - 0.2, "Phase 3: Finalisation", fontsize=7.6, color=th["mid"], ha="center", style="italic")
    ax.text(LEFT + 0.2, len(tasks) - 0.2, "Task / Milestone", fontsize=8.2, color=th["dark"], fontweight="bold")
    fig.tight_layout()
    fig.savefig(f"{A}/gantt_{pid}.png", dpi=200, bbox_inches="tight")
    plt.close(fig)

# ============================================================================
# 6. PROJECT-SPECIFIC ILLUSTRATIVE CHARTS
# ============================================================================
def chart_p1():
    th = THEMES[1]
    # (a) literature comparison bar chart: MAE on ShanghaiTech Part A
    fig, axes = plt.subplots(1, 2, figsize=(9.4, 4.0), dpi=200)
    models = ["MCNN\n(2016)", "CSRNet\n(2018)", "CAN\n(2019)", "BL\n(2019)", "DM-Count\n(2020)", "Ours\n(CSRNet-lite)"]
    mae = [110.2, 68.2, 62.3, 62.8, 59.7, 72.5]
    cols = [th["light"]] * 5 + [th["accent"]]
    b = axes[0].bar(models, mae, color=cols, edgecolor=th["dark"], lw=0.8)
    axes[0].set_ylabel("MAE (lower is better)"); axes[0].set_title("Crowd-Counting Methods — ShanghaiTech Part A", fontsize=10, fontweight="bold")
    for rect, v in zip(b, mae):
        axes[0].text(rect.get_x() + rect.get_width() / 2, v + 1.5, f"{v}", ha="center", fontsize=8)
    axes[0].grid(axis="y", color="#ddd", lw=0.5);
    for sp in ["top", "right"]: axes[0].spines[sp].set_visible(False)
    # (b) synthetic density-map heatmap
    x = np.linspace(0, 1, 120); y = np.linspace(0, 1, 90)
    X, Y = np.meshgrid(x, y)
    Z = np.zeros_like(X)
    centers = np.random.rand(40, 2) * [1, 1]
    centers[:25] = np.random.normal(0.65, 0.12, (25, 2))
    for cx, cy in centers:
        Z += np.exp(-((X - cx) ** 2 + (Y - cy) ** 2) / (2 * 0.012))
    im = axes[1].imshow(Z, cmap="jet", origin="lower", aspect="auto")
    axes[1].set_title("Predicted Density Map (∫ ≈ estimated count)", fontsize=10, fontweight="bold")
    axes[1].set_xticks([]); axes[1].set_yticks([])
    fig.colorbar(im, ax=axes[1], fraction=0.046, pad=0.04)
    fig.tight_layout(); fig.savefig(f"{A}/chart_1.png", dpi=200, bbox_inches="tight"); plt.close(fig)

def chart_p2():
    th = THEMES[2]
    fig, axes = plt.subplots(1, 2, figsize=(9.4, 4.0), dpi=200)
    # (a) CER comparison on IAM
    models = ["MDLSTM\n(2009)", "CRNN\n(2015)", "Puigcerver\n(2017)", "Bluche\n(2017)", "TrOCR-S\n(2021)", "Ours\n(CRNN-lite)"]
    cer = [18.2, 11.9, 8.2, 7.9, 4.2, 9.8]
    cols = [th["light"]] * 5 + [th["accent"]]
    b = axes[0].bar(models, cer, color=cols, edgecolor=th["dark"], lw=0.8)
    axes[0].set_ylabel("Character Error Rate % (lower is better)")
    axes[0].set_title("Handwriting Recognition — IAM Test Set", fontsize=10, fontweight="bold")
    for rect, v in zip(b, cer):
        axes[0].text(rect.get_x() + rect.get_width() / 2, v + 0.25, f"{v}", ha="center", fontsize=8)
    for sp in ["top", "right"]: axes[0].spines[sp].set_visible(False)
    axes[0].grid(axis="y", color="#ddd", lw=0.5)
    # (b) training curves
    ep = np.arange(1, 41)
    tr = 45 * np.exp(-ep / 7) + 6 + np.random.normal(0, 0.6, 40)
    va = 45 * np.exp(-ep / 8) + 9 + np.random.normal(0, 0.8, 40)
    axes[1].plot(ep, tr, color=th["accent"], lw=2, label="Train CER")
    axes[1].plot(ep, va, color=th["light"], lw=2, label="Validation CER")
    axes[1].fill_between(ep, va - 1, va + 1, color=th["light"], alpha=0.15)
    axes[1].set_xlabel("Epoch"); axes[1].set_ylabel("CER %"); axes[1].legend()
    axes[1].set_title("Expected Training Convergence (illustrative)", fontsize=10, fontweight="bold")
    for sp in ["top", "right"]: axes[1].spines[sp].set_visible(False)
    axes[1].grid(color="#eee", lw=0.5)
    fig.tight_layout(); fig.savefig(f"{A}/chart_2.png", dpi=200, bbox_inches="tight"); plt.close(fig)

def chart_p3():
    th = THEMES[3]
    fig, axes = plt.subplots(1, 2, figsize=(9.4, 4.0), dpi=200)
    # (a) accuracy comparison
    models = ["CNN\n(raw img)", "HMM\n(classic)", "VGG-16\n(TL)", "MP +\nMLP*", "MP +\nLSTM*", "I3D\n(video)"]
    acc = [92.5, 86.0, 95.1, 97.8, 98.6, 99.1]
    cols = [th["light"], th["light"], th["light"], th["accent"], th["accent"], th["light"]]
    b = axes[0].bar(models, acc, color=cols, edgecolor=th["dark"], lw=0.8)
    axes[0].set_ylim(80, 100); axes[0].set_ylabel("Accuracy % (ASL alphabet)")
    axes[0].set_title("Sign-Recognition Approaches  (* = our target)", fontsize=10, fontweight="bold")
    axes[0].tick_params(axis="x", labelsize=8.5)
    for rect, v in zip(b, acc):
        axes[0].text(rect.get_x() + rect.get_width() / 2, v + 0.2, f"{v}", ha="center", fontsize=8)
    for sp in ["top", "right"]: axes[0].spines[sp].set_visible(False)
    axes[0].grid(axis="y", color="#ddd", lw=0.5)
    # (b) confusion-matrix style heatmap (subset of letters)
    letters = list("ABCDEFGHIJ")
    cm = np.eye(10) * np.random.uniform(0.9, 0.99, 10)
    for i in range(10):
        for j in range(10):
            if i != j and np.random.rand() < 0.25:
                cm[i, j] = np.random.uniform(0.0, 0.05)
        cm[i] = cm[i] / cm[i].sum()
    im = axes[1].imshow(cm, cmap="Greens", vmin=0, vmax=1)
    axes[1].set_xticks(range(10)); axes[1].set_xticklabels(letters, fontsize=8)
    axes[1].set_yticks(range(10)); axes[1].set_yticklabels(letters, fontsize=8)
    axes[1].set_xlabel("Predicted"); axes[1].set_ylabel("Actual")
    axes[1].set_title("Illustrative Confusion Matrix (subset)", fontsize=10, fontweight="bold")
    fig.colorbar(im, ax=axes[1], fraction=0.046, pad=0.04)
    fig.tight_layout(); fig.savefig(f"{A}/chart_3.png", dpi=200, bbox_inches="tight"); plt.close(fig)

# ============================================================================
# 7. SHARED "AI" DECORATIVE ICON STRIP (small)
# ============================================================================
def ai_icon(pid):
    th = THEMES[pid]
    fig, ax = plt.subplots(figsize=(1.6, 1.6), dpi=200)
    ax.axis("off"); ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    c = Circle((0.5, 0.5), 0.46, fc=th["dark"], ec=th["accent"], lw=2)
    ax.add_patch(c)
    neural_motif(ax, 0.18, 0.2, 0.64, 0.6, [2, 3, 2], th["light"], alpha=0.95, node_s=40, lw=1.0)
    fig.savefig(f"{A}/icon_{pid}.png", dpi=200, bbox_inches="tight", transparent=True)
    plt.close(fig)

# ============================================================================
def main():
    for pid in (1, 2, 3):
        print("Generating assets for proposal", pid)
        cover_banner(pid)
        divider(pid)
        architecture(pid)
        flowchart(pid)
        gantt(pid)
        ai_icon(pid)
    chart_p1(); chart_p2(); chart_p3()
    print("All assets generated:", sorted(os.listdir(A)))

if __name__ == "__main__":
    main()
