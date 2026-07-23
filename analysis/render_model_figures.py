"""
Render the most important Excel-model tabs as clean, embed-ready PNG figures,
reading the computed values straight from the workbook. These are appendix
exhibits for the proposal (Figure C1 and C2).

Output: deliverables/figures/Fig_C1_refill_by_notice.png
        deliverables/figures/Fig_C2_lastminute_correction.png
"""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
WB = ROOT / "deliverables" / "Clipboard_Analysis_Models.xlsx"
FIGDIR = ROOT / "deliverables" / "figures"
FIGDIR.mkdir(exist_ok=True)

ACCENT = "#8a3033"
INK = "#1a1a2b"
HEAD = "#e9e5ee"
STRIPE = "#f6f4f8"
GRID = "#c9c9d1"
TABCHIP = "#217346"  # Excel green

wv = load_workbook(WB, data_only=True)


def pct(x):
    return f"{x*100:.1f}%"


def money(x):
    return f"${x:,.0f}"


def draw_table(ax, title, subtitle, columns, rows, colweights, y0=0.90, row_h=0.075,
               bold_last=False, x0=0.02, x1=0.98):
    """Draw a styled Excel-like table onto ax starting at y0."""
    n = len(columns)
    edges = [x0]
    span = x1 - x0
    tot = sum(colweights)
    for w in colweights:
        edges.append(edges[-1] + span * w / tot)

    ax.text(x0, y0 + 0.055, title, fontsize=12.5, fontweight="bold", color=ACCENT, va="bottom")
    if subtitle:
        ax.text(x0, y0 + 0.018, subtitle, fontsize=8.2, style="italic", color="#555555", va="bottom")

    y = y0
    # header
    for j in range(n):
        ax.add_patch(plt.Rectangle((edges[j], y - row_h), edges[j + 1] - edges[j], row_h,
                                   facecolor=HEAD, edgecolor=GRID, lw=0.8, zorder=1))
        ha = "left" if j == 0 else "center"
        tx = edges[j] + 0.008 if j == 0 else (edges[j] + edges[j + 1]) / 2
        ax.text(tx, y - row_h / 2, columns[j], fontsize=8.6, fontweight="bold", color=INK,
                va="center", ha=ha, zorder=2)
    y -= row_h
    # body
    for ri, r in enumerate(rows):
        fill = STRIPE if ri % 2 else "white"
        for j in range(n):
            ax.add_patch(plt.Rectangle((edges[j], y - row_h), edges[j + 1] - edges[j], row_h,
                                       facecolor=fill, edgecolor=GRID, lw=0.8, zorder=1))
            ha = "left" if j == 0 else "center"
            tx = edges[j] + 0.008 if j == 0 else (edges[j] + edges[j + 1]) / 2
            bold = (bold_last and ri == len(rows) - 1) or j == 0
            ax.text(tx, y - row_h / 2, r[j], fontsize=8.5, color=INK, va="center", ha=ha,
                    fontweight="bold" if bold else "normal", zorder=2)
        y -= row_h
    return y


def sheet_chrome(fig, tabname):
    """Add a faux Excel sheet-tab chip at the bottom to signal the source tab."""
    ax = fig.add_axes([0, 0, 1, 0.045], zorder=5)
    ax.axis("off")
    ax.add_patch(plt.Rectangle((0.02, 0.15), 0.30, 0.75, facecolor="white",
                               edgecolor=TABCHIP, lw=1.4))
    ax.add_patch(plt.Rectangle((0.02, 0.05), 0.30, 0.06, facecolor=TABCHIP, edgecolor="none"))
    ax.text(0.035, 0.52, tabname, fontsize=8.6, fontweight="bold", color=TABCHIP, va="center")
    ax.text(0.98, 0.5, "Clipboard_Analysis_Models.xlsx", fontsize=7.6, color="#888888",
            va="center", ha="right", style="italic")
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)


# ---------------------------------------------------------------- Figure C1
so = wv["Calc_ShiftOutcomes"]
# refill-by-type block: rows 3,4,5 -> B shifts, C worked, D empty, E facdel, F refill%, G empty%
def rowvals(r):
    return so[f"B{r}"].value, so[f"C{r}"].value, so[f"D{r}"].value, so[f"E{r}"].value, so[f"F{r}"].value, so[f"G{r}"].value

r_labels = ["Early cancel (24h+ notice)", "Late cancel (under 24h)", "No-call-no-show"]
refill_rows = []
for i, r in enumerate([3, 4, 5]):
    b, c, d, e, f, g = rowvals(r)
    refill_rows.append([r_labels[i], f"{b:,}", pct(f), pct(g), pct(e / b)])

# empty sources block: find the rows (Late/NCNS/Early/TOTAL) by scanning col A
src_rows = []
for r in range(6, 20):
    a = so[f"A{r}"].value
    if a in ("Late cancel", "NCNS", "Early cancel", "TOTAL"):
        src_rows.append([a if a != "NCNS" else "No-show", f"{int(so[f'B{r}'].value):,}", pct(so[f"C{r}"].value)])

fig = plt.figure(figsize=(8.6, 6.4), dpi=200)
ax = fig.add_axes([0, 0.05, 1, 0.93]); ax.axis("off"); ax.set_xlim(0, 1); ax.set_ylim(0, 1)
y = draw_table(
    ax,
    "Refill outcomes by final cancel event",
    "Each shift counted once, classified by its final cancel event. Recovery collapses as notice shrinks.",
    ["Final cancel type", "Shifts", "Refilled\n& worked", "Died\nempty", "Deleted by\nfacility"],
    refill_rows, [3.1, 1.0, 1.2, 1.0, 1.2], y0=0.90, row_h=0.085)
y = draw_table(
    ax,
    "Where the cancel-driven empty shifts come from",
    "Late cancels and no-shows are 89% of the damage.",
    ["Source", "Empty shifts", "% of cancel-driven empties"],
    src_rows, [2.2, 1.3, 2.2], y0=y - 0.10, row_h=0.085, bold_last=True)
sheet_chrome(fig, "Calc_ShiftOutcomes")
fig.savefig(FIGDIR / "Fig_C1_refill_by_notice.png", dpi=200, bbox_inches="tight", facecolor="white")
plt.close(fig)

# ---------------------------------------------------------------- Figure C2
lm = wv["Calc_LastMin_Correction"]
corr_rows = [
    ["Shift-level fail (naive; counts failures that predate the claim)",
     pct(lm["B3"].value), pct(lm["C3"].value), pct(lm["D3"].value)],
    ["Worker-level: this claimant late-cancels / NCNSes after booking",
     pct(lm["B4"].value), pct(lm["C4"].value), pct(lm["D4"].value)],
    ["Worker-level: any cancellation after booking",
     pct(lm["B5"].value), pct(lm["C5"].value), pct(lm["D5"].value)],
]
# rescue evidence block rows (scan)
eb = []
for r in range(6, 20):
    a = lm[f"A{r}"].value
    if not a:
        continue
    v = lm[f"B{r}"].value
    if isinstance(v, float) and v < 1:
        vs = pct(v)
    elif isinstance(v, (int, float)):
        vs = f"{int(v):,}"
    else:
        vs = str(v)
    if any(k in str(a) for k in ["Last-minute claims (", "RESCUES", "Rescue share", "held", "WORKED"]):
        eb.append([str(a), vs])

fig = plt.figure(figsize=(8.8, 7.4), dpi=200)
ax = fig.add_axes([0, 0.055, 1, 0.9]); ax.axis("off"); ax.set_xlim(0, 1); ax.set_ylim(0, 1)
y = draw_table(
    ax,
    "The last-minute claim correction",
    "The naive shift-level view brands last-minute claimers 2x risky. Scored at the worker level, they are the most reliable segment.",
    ["Definition of 'failure'", "Base", "Last-minute\nclaims", "Booked\nahead"],
    corr_rows, [4.2, 1.0, 1.3, 1.1], y0=0.92, row_h=0.068)
y = draw_table(
    ax,
    "Why: the claim is the rescue, not the risk",
    "27% of last-minute claims land on shifts someone else already cancelled.",
    ["Rescue evidence", "Value"],
    eb, [4.6, 1.4], y0=y - 0.105, row_h=0.06)
sheet_chrome(fig, "Calc_LastMin_Correction")
fig.savefig(FIGDIR / "Fig_C2_lastminute_correction.png", dpi=200, bbox_inches="tight", facecolor="white")
plt.close(fig)

print("wrote:")
for p in sorted(FIGDIR.glob("*.png")):
    print("  ", p, p.stat().st_size, "bytes")
