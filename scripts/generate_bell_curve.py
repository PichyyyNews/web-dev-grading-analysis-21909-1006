# -*- coding: utf-8 -*-
import os
import sys
from pathlib import Path
import numpy as np
import scipy.stats as stats
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# Ensure UTF-8 stdout
sys.stdout.reconfigure(encoding='utf-8')

# 1. Setup Thai Fonts (TH SarabunPSK)
font_regular = Path(r"C:\Users\Newsk\.gemini\config\skills\mathplot\fonts\THSarabun.ttf")
font_bold = Path(r"C:\Users\Newsk\.gemini\config\skills\mathplot\fonts\THSarabun Bold.ttf")

fm.fontManager.addfont(str(font_regular))
fm.fontManager.addfont(str(font_bold))

prop_regular = fm.FontProperties(fname=str(font_regular))
prop_bold = fm.FontProperties(fname=str(font_bold), weight='bold')

font_family = prop_regular.get_name()

# 2. Configure Matplotlib RC Params (Academic Document Theme)
mpl.rcParams.update({
    "figure.constrained_layout.use": True,
    "figure.facecolor": "#FFFFFF",
    "axes.facecolor": "#FFFFFF",
    "axes.edgecolor": "#CBD5E1",
    "axes.linewidth": 1.0,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.spines.left": True,
    "axes.spines.bottom": True,
    "axes.grid": True,
    "axes.grid.axis": "y",             # Grid lines on Y-axis only (horizontal)
    "grid.color": "#E2E8F0",
    "grid.linestyle": "--",
    "grid.linewidth": 0.7,
    "grid.alpha": 0.85,
    "axes.axisbelow": True,
    "font.family": font_family,
    "font.size": 18,
    "axes.titlesize": 22,
    "axes.titleweight": "bold",
    "axes.labelsize": 20,
    "axes.labelweight": "bold",
    "xtick.labelsize": 18,
    "ytick.labelsize": 18,
    "xtick.color": "#111827",
    "ytick.color": "#111827",
    "axes.labelcolor": "#111827",
    "legend.frameon": False,
    "legend.fontsize": 18,
    "axes.unicode_minus": False,
})

# 3. Data Preparation
scores_room1 = np.array([14.5, 31.5, 50.4, 51.2, 57.3, 61.3, 64.9, 66.7, 69.2, 70.2, 73.7, 74.7, 75.7, 79.0, 80.4, 82.1, 86.7, 87.6])
scores_room2 = np.array([24.5, 32.3, 34.2, 45.7, 49.4, 59.6, 60.1, 62.6, 62.8, 65.3, 67.0, 70.2, 74.6, 75.6, 76.2, 79.3])

mean1, std1 = np.mean(scores_room1), np.std(scores_room1, ddof=1)
mean2, std2 = np.mean(scores_room2), np.std(scores_room2, ddof=1)

# Generate Normal curves
x = np.linspace(0, 100, 500)
pdf1 = stats.norm.pdf(x, mean1, std1)
pdf2 = stats.norm.pdf(x, mean2, std2)

# 4. Create Plot
fig, ax = plt.subplots(figsize=(14.8, 7.5), dpi=300)

color1 = "#1D4ED8"       # Navy Blue (Room 1)
color2 = "#0D9488"       # Teal (Room 2)
color_thresh = "#DC2626" # Crimson Red

# Fill under curves with subtle transparency
ax.fill_between(x, pdf1, color=color1, alpha=0.14)
ax.fill_between(x, pdf2, color=color2, alpha=0.14)

# Plot Bell Curves
line1, = ax.plot(x, pdf1, color=color1, linewidth=2.8, label=f"ห้อง 1 : ค่าเฉลี่ย = {mean1:.2f}, S.D. = {std1:.2f}")
line2, = ax.plot(x, pdf2, color=color2, linewidth=2.8, label=f"ห้อง 2 : ค่าเฉลี่ย = {mean2:.2f}, S.D. = {std2:.2f}")

# Mean vertical lines (drop from peak down to baseline)
peak1 = stats.norm.pdf(mean1, mean1, std1)
peak2 = stats.norm.pdf(mean2, mean2, std2)

ax.vlines(mean1, 0, peak1, color=color1, linestyle=":", linewidth=1.8, alpha=0.9, zorder=3)
ax.vlines(mean2, 0, peak2, color=color2, linestyle=":", linewidth=1.8, alpha=0.9, zorder=3)

# Mean Annotations above peaks
ax.text(
    mean1, peak1 + 0.0009,
    f"Mean = {mean1:.1f}",
    ha="center", va="bottom",
    color=color1, fontproperties=prop_bold, fontsize=17
)

ax.text(
    mean2, peak2 + 0.0009,
    f"Mean = {mean2:.1f}",
    ha="center", va="bottom",
    color=color2, fontproperties=prop_bold, fontsize=17
)

# Threshold Line at 50 (เกณฑ์ผ่าน 50 คะแนน)
# Position text to the LEFT of the line with ha='right' to prevent any collision with Room 2 peak
thresh_y_top = 0.0265
ax.vlines(50, 0, thresh_y_top, color=color_thresh, linestyle="--", linewidth=1.6, zorder=4)
ax.text(
    49.0, 0.0250,
    "เกณฑ์ผ่าน 50 คะแนน",
    ha="right", va="center",
    color=color_thresh, fontproperties=prop_bold, fontsize=17
)

# Individual Student Score Points (inside plot area, elegant white-edged markers)
y_pts1 = np.full_like(scores_room1, 0.0006)
y_pts2 = np.full_like(scores_room2, 0.0013)

ax.scatter(
    scores_room1, y_pts1,
    s=36, color=color1, edgecolors="white", linewidths=0.6,
    zorder=5, label=f"คะแนนจริง ห้อง 1 ({len(scores_room1)} คน)"
)
ax.scatter(
    scores_room2, y_pts2,
    s=36, color=color2, edgecolors="white", linewidths=0.6,
    zorder=5, label=f"คะแนนจริง ห้อง 2 ({len(scores_room2)} คน)"
)

# 5. Title and Labels
ax.set_title(
    "การแจกแจงคะแนนรวมแบบโค้งปกติ (Normal Distribution Curve / กราฟระฆังคว่ำ)",
    fontproperties=prop_bold,
    fontsize=22,
    pad=22,
    color="#111827"
)

ax.set_xlabel(
    "คะแนนรวมสุทธิ (เต็ม 100 คะแนน ตามสัดส่วน 20 : 50 : 10 : 20)",
    fontproperties=prop_bold,
    fontsize=20,
    labelpad=12,
    color="#111827"
)

ax.set_ylabel(
    "ความหนาแน่นของความน่าจะเป็น (Probability Density)",
    fontproperties=prop_bold,
    fontsize=20,
    labelpad=14,
    color="#111827",
    rotation=90
)

ax.set_xlim(0, 100)
ax.set_xticks(np.arange(0, 101, 10))
ax.set_xticklabels([str(val) for val in np.arange(0, 101, 10)], fontproperties=prop_regular, fontsize=18)

y_max = 0.0295
ax.set_ylim(0, y_max)
y_ticks = np.linspace(0, 0.025, 6)
ax.set_yticks(y_ticks)
ax.set_yticklabels([f"{y:.3f}" for y in y_ticks], fontproperties=prop_regular, fontsize=17)

# 6. Legend Placement (Whitespace analysis: upper left quadrant x=3 to 35 is wide open)
ax.legend(
    loc="upper left",
    bbox_to_anchor=(0.03, 0.97),
    frameon=False,
    prop=prop_bold,
    fontsize=17.5,
    handlelength=1.2,
    handletextpad=0.5,
    labelspacing=0.5
)

# 7. Save Outputs
output_dir = Path(r"c:\Users\Newsk\Downloads\1ชทค\charts")
output_dir.mkdir(parents=True, exist_ok=True)

png_path = output_dir / "bell_curve_score_distribution.png"
svg_path = output_dir / "bell_curve_score_distribution.svg"

plt.savefig(png_path, dpi=300, bbox_inches="tight")
plt.savefig(svg_path, bbox_inches="tight")
plt.close(fig)

print(f"Chart successfully saved:")
print(f"PNG: {png_path}")
print(f"SVG: {svg_path}")
