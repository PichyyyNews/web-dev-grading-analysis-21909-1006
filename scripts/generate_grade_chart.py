# -*- coding: utf-8 -*-
import os
import sys
from pathlib import Path
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# Ensure UTF-8 stdout
sys.stdout.reconfigure(encoding='utf-8')

# 1. Setup Thai Fonts (TH SarabunPSK)
font_regular = Path(r"C:\Users\Newsk\.gemini\config\skills\mathplot\fonts\THSarabun.ttf")
font_bold = Path(r"C:\Users\Newsk\.gemini\config\skills\mathplot\fonts\THSarabun Bold.ttf")

if not font_regular.exists():
    raise FileNotFoundError(f"Regular font missing: {font_regular}")
if not font_bold.exists():
    raise FileNotFoundError(f"Bold font missing: {font_bold}")

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
grades = ['4.0', '3.5', '3.0', '2.5', '2.0', '1.5', '1.0', '0.0']
grade_ticks = [
    '4.0\n(80-100)',
    '3.5\n(75-79)',
    '3.0\n(70-74)',
    '2.5\n(65-69)',
    '2.0\n(60-64)',
    '1.5\n(55-59)',
    '1.0\n(50-54)',
    '0.0\n(<50)'
]

room1_counts = [4, 2, 3, 2, 2, 1, 2, 2]   # 18 students
room2_counts = [0, 3, 2, 2, 3, 1, 0, 5]   # 16 students

room1_total = sum(room1_counts)
room2_total = sum(room2_counts)

group_step = 1.25
x = np.arange(len(grades)) * group_step
bar_width = 0.40
bar_gap = 0.06
offset = bar_width / 2 + bar_gap / 2  # 0.23

# 4. Create Plot
fig, ax = plt.subplots(figsize=(14.8, 7.5), dpi=300)

# Academic Document Palette
color_room1 = "#1D4ED8"  # Navy Blue
color_room2 = "#0D9488"  # Teal

bars1 = ax.bar(
    x - offset,
    room1_counts,
    width=bar_width,
    label=f"ห้อง 1 ({room1_total} คน)",
    color=color_room1,
    zorder=3
)

bars2 = ax.bar(
    x + offset,
    room2_counts,
    width=bar_width,
    label=f"ห้อง 2 ({room2_total} คน)",
    color=color_room2,
    zorder=3
)

# 5. Add Direct Data Labels with Zero Overlap
for i in range(len(grades)):
    # Room 1 bar
    c1 = room1_counts[i]
    pct1 = (c1 / room1_total) * 100
    lbl1 = f"{c1} คน\n({pct1:.1f}%)"
    ax.text(
        x[i] - offset,
        c1 + 0.12,
        lbl1,
        ha='center',
        va='bottom',
        color="#111827",
        fontproperties=prop_regular,
        fontsize=15.5,
        linespacing=0.88
    )
    
    # Room 2 bar
    c2 = room2_counts[i]
    pct2 = (c2 / room2_total) * 100
    lbl2 = f"{c2} คน\n({pct2:.1f}%)"
    if c2 == 0:
        # Floor baseline marker and clean label for zero count
        ax.plot(
            [x[i] + offset - bar_width/2 + 0.05, x[i] + offset + bar_width/2 - 0.05],
            [0.03, 0.03],
            color="#94A3B8",
            linewidth=1.8,
            zorder=3
        )
        ax.text(
            x[i] + offset,
            0.15,
            lbl2,
            ha='center',
            va='bottom',
            color="#64748B",
            fontproperties=prop_regular,
            fontsize=15.0,
            linespacing=0.88
        )
    else:
        ax.text(
            x[i] + offset,
            c2 + 0.12,
            lbl2,
            ha='center',
            va='bottom',
            color="#111827",
            fontproperties=prop_regular,
            fontsize=15.5,
            linespacing=0.88
        )

# 6. Title and Axis Setup
ax.set_title(
    "การแจกแจงระดับผลการเรียน 8 ระดับ วิชาพื้นฐานการสร้างเว็บไซต์ (สัดส่วน 20 : 50 : 10 : 20)",
    fontproperties=prop_bold,
    fontsize=22,
    pad=22,
    color="#111827"
)

ax.set_xlabel(
    "ระดับผลการเรียน (เกรดตามเกณฑ์มาตรฐาน สอศ.)",
    fontproperties=prop_bold,
    fontsize=20,
    labelpad=12,
    color="#111827"
)

ax.set_ylabel(
    "จำนวนนักเรียน (คน)",
    fontproperties=prop_bold,
    fontsize=20,
    labelpad=14,
    color="#111827",
    rotation=90
)

ax.set_xticks(x)
ax.set_xticklabels(grade_ticks, fontproperties=prop_regular, fontsize=17.5)

ax.set_xlim(-0.65, (len(grades)-1)*group_step + 0.65)
ax.set_ylim(0, 6.5)
ax.set_yticks(range(0, 7))
ax.set_yticklabels([str(y) for y in range(0, 7)], fontproperties=prop_regular, fontsize=18)

# 7. Legend Placement (Whitespace analysis: upper center is wide open)
legend = ax.legend(
    loc='upper center',
    bbox_to_anchor=(0.50, 0.98),
    ncols=2,
    frameon=False,
    prop=prop_bold,
    fontsize=18,
    handlelength=1.1,
    handletextpad=0.5,
    columnspacing=2.2
)

# 8. Save Outputs (PNG 300 DPI and Vector SVG)
output_dir = Path(r"c:\Users\Newsk\Downloads\1ชทค\charts")
output_dir.mkdir(parents=True, exist_ok=True)

png_path = output_dir / "grade_distribution_grouped_bar.png"
svg_path = output_dir / "grade_distribution_grouped_bar.svg"

plt.savefig(png_path, dpi=300, bbox_inches='tight')
plt.savefig(svg_path, bbox_inches='tight')
plt.close(fig)

print(f"Chart successfully saved:")
print(f"PNG: {png_path}")
print(f"SVG: {svg_path}")
