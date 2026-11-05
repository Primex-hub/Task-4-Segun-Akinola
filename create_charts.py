"""Project 4 - Data Visualisation
Generates all charts for the e-commerce orders dataset using matplotlib.

Usage:
    python create_charts.py "C:\\path\\to\\Dataset for Data Analytics (1).xlsx"
If no path is supplied, the default local path is used.
"""

import csv
import os
import sys
from collections import defaultdict
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.ticker import FuncFormatter
from PIL import Image, ImageDraw, ImageFont

BASE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SRC = os.path.join(BASE, "data", "Dataset for Data Analytics (1).xlsx")
if not os.path.exists(DEFAULT_SRC):
    DEFAULT_SRC = r"C:\Users\Prime\Downloads\Dataset for Data Analytics (1).xlsx"
SRC = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SRC
CHART_DIR = os.path.join(BASE, "charts")
os.makedirs(CHART_DIR, exist_ok=True)

# ----------------------------------------------------------------------
# Load and aggregate the workbook
# ----------------------------------------------------------------------
wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
ws = wb["Sheet1"]
rows = list(ws.iter_rows(values_only=True))
headers = [str(c).strip() for c in rows[0]]
data = [
    dict(zip(headers, r))
    for r in rows[1:]
    if any(v is not None and str(v).strip() != "" for v in r)
]
wb.close()

N = len(data)
TOTAL_REV = sum(r["TotalPrice"] for r in data)
TOTAL_QTY = sum(r["Quantity"] for r in data)
AOV = TOTAL_REV / N


def group(col):
    g = defaultdict(list)
    for r in data:
        g[r[col]].append(r)
    out = {}
    for key, rs in g.items():
        rev = sum(x["TotalPrice"] for x in rs)
        out[key] = {
            "orders": len(rs),
            "units": sum(x["Quantity"] for x in rs),
            "revenue": rev,
            "aov": rev / len(rs),
            "share": rev / TOTAL_REV * 100,
        }
    return out


products = group("Product")
payments = group("PaymentMethod")
statuses = group("OrderStatus")
referrals = group("ReferralSource")

yearly = defaultdict(list)
monthly = defaultdict(list)
customer_orders = defaultdict(int)
for r in data:
    d = r["Date"]
    yearly[d.year].append(r)
    monthly[(d.year, d.month)].append(r)
    customer_orders[r["CustomerID"]] += 1

year_stats = {}
for y, rs in yearly.items():
    rev = sum(x["TotalPrice"] for x in rs)
    year_stats[y] = {
        "orders": len(rs),
        "units": sum(x["Quantity"] for x in rs),
        "revenue": rev,
        "aov": rev / len(rs),
        "share": rev / TOTAL_REV * 100,
    }

month_keys = sorted(monthly)
month_dates = [datetime(y, m, 1) for y, m in month_keys]
month_rev = [sum(x["TotalPrice"] for x in monthly[k]) for k in month_keys]
month_orders = [len(monthly[k]) for k in month_keys]
moving_avg = [
    np.nan if i < 2 else float(np.mean(month_rev[i - 2:i + 1]))
    for i in range(len(month_rev))
]

unique_customers = len(customer_orders)
repeat_customers = sum(1 for v in customer_orders.values() if v > 1)
repeat_orders = sum(v for v in customer_orders.values() if v > 1)
one_time_customers = unique_customers - repeat_customers

delivered = statuses["Delivered"]["revenue"]
delivered_orders = statuses["Delivered"]["orders"]
not_delivered = TOTAL_REV - delivered

order_values = defaultdict(list)
for r in data:
    order_values[r["Product"]].append(r["TotalPrice"])

# ----------------------------------------------------------------------
# Visual style
# ----------------------------------------------------------------------
INK = "#111827"
MUTED = "#6B7280"
GRID = "#E5E7EB"
TEAL = "#0F766E"
BLUE = "#2563EB"
ORANGE = "#F59E0B"
RED = "#DC2626"
GRAY = "#9CA3AF"
PALE = "#F3F4F6"
LIGHT_TEAL = "#99F6E4"
LIGHT_ORANGE = "#FDE68A"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 11,
    "axes.titlesize": 16,
    "axes.titleweight": "bold",
    "axes.labelsize": 11,
    "axes.edgecolor": "#D1D5DB",
    "axes.labelcolor": INK,
    "xtick.color": MUTED,
    "ytick.color": MUTED,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
})

usd = FuncFormatter(lambda x, p: f"${x:,.0f}")


def clean_ax(ax, grid_axis="x"):
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    ax.spines["left"].set_color("#D1D5DB")
    ax.spines["bottom"].set_color("#D1D5DB")
    ax.grid(axis=grid_axis, color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(axis="both", length=0)


def titles(fig, title, subtitle):
    fig.text(0.06, 0.955, title, fontsize=18, fontweight="bold", color=INK, va="top")
    fig.text(0.06, 0.905, subtitle, fontsize=11, color=MUTED, va="top")


def footer(fig, text):
    fig.text(0.06, 0.025, text, fontsize=8, color=MUTED, va="bottom")


def save(fig, name):
    path = os.path.join(CHART_DIR, name)
    fig.savefig(path, dpi=200, facecolor="white")
    plt.close(fig)
    print(f"saved {name}")


SOURCE = "Source: Dataset for Data Analytics (1).xlsx  |  1,200 orders  |  1 Jan 2023 - 30 Jun 2025"

# ----------------------------------------------------------------------
# 01 - Revenue by product
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7))
fig.subplots_adjust(left=0.18, right=0.94, top=0.82, bottom=0.14)
prod_order = sorted(products, key=lambda p: products[p]["revenue"])
prod_names = prod_order
prod_revs = [products[p]["revenue"] for p in prod_order]
prod_colors = [TEAL if p == "Chair" else ORANGE if p == "Phone" else BLUE for p in prod_order]

bars = ax.barh(prod_names, prod_revs, color=prod_colors, height=0.66, alpha=0.92)
for bar, p in zip(bars, prod_order):
    v = products[p]["revenue"]
    ax.text(v + TOTAL_REV * 0.008, bar.get_y() + bar.get_height() / 2,
            f"${v:,.0f}", va="center", ha="left", fontsize=10, color=INK, fontweight="bold")
ax.set_xlim(0, max(prod_revs) * 1.22)
ax.xaxis.set_major_formatter(usd)
ax.set_xlabel("Recorded revenue (USD)")
ax.set_yticks(range(len(prod_names)), prod_names, fontsize=11)
clean_ax(ax, "x")
ax.annotate("Chair leads Printer by just $7.50",
            xy=(products["Chair"]["revenue"], prod_names.index("Chair")),
            xytext=(-10, 34), textcoords="offset points",
            fontsize=9, color=TEAL, fontweight="bold",
            arrowprops=dict(arrowstyle="-", color=TEAL, lw=1))
titles(fig, "Chair and Printer Lead Recorded Revenue",
       "Recorded order value by product  ·  all 1,200 orders")
footer(fig, SOURCE)
save(fig, "01_revenue_by_product.png")

# ----------------------------------------------------------------------
# 02 - Monthly revenue trend
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7))
fig.subplots_adjust(left=0.10, right=0.95, top=0.82, bottom=0.15)
ax.fill_between(month_dates, month_rev, color=TEAL, alpha=0.10)
ax.plot(month_dates, month_rev, color=TEAL, linewidth=2.2, marker="o", markersize=3.2,
        label="Monthly revenue")
ax.plot(month_dates, moving_avg, color=ORANGE, linewidth=2.0, linestyle="--",
        label="3-month moving average")
ax.axvspan(datetime(2025, 1, 1), datetime(2025, 7, 1), color=GRAY, alpha=0.10, linewidth=0)
ax.text(datetime(2025, 3, 20), max(month_rev) * 0.90, "2025 partial\nyear",
        ha="center", fontsize=9, color=MUTED)

peak_i = int(np.argmax(month_rev))
low_i = int(np.argmin(month_rev))
ax.scatter([month_dates[peak_i]], [month_rev[peak_i]], s=70, color=ORANGE, zorder=5)
ax.annotate(f"Peak: ${month_rev[peak_i]:,.0f}\nJun 2024",
            xy=(month_dates[peak_i], month_rev[peak_i]),
            xytext=(-95, 35), textcoords="offset points",
            fontsize=9, color=INK, fontweight="bold",
            arrowprops=dict(arrowstyle="-", color=ORANGE, lw=1))
ax.scatter([month_dates[low_i]], [month_rev[low_i]], s=55, color=RED, zorder=5)
ax.annotate(f"Low: ${month_rev[low_i]:,.0f}\nApr 2023",
            xy=(month_dates[low_i], month_rev[low_i]),
            xytext=(18, -48), textcoords="offset points",
            fontsize=9, color=INK,
            arrowprops=dict(arrowstyle="-", color=RED, lw=1))

ax.set_ylim(0, max(month_rev) * 1.20)
ax.set_xlim(datetime(2022, 12, 15), datetime(2025, 7, 15))
ax.yaxis.set_major_formatter(usd)
ax.set_ylabel("Recorded revenue (USD)")
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))
ax.tick_params(axis="x", labelsize=9)
ax.legend(frameon=False, loc="upper left", fontsize=9)
clean_ax(ax, "y")
titles(fig, "Monthly Revenue Peaked in June 2024",
       "Monthly recorded revenue, Jan 2023 - Jun 2025  ·  3-month moving average")
footer(fig, SOURCE)
save(fig, "02_monthly_revenue_trend.png")

# ----------------------------------------------------------------------
# 03 - Payment method revenue share (donut)
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7))
fig.subplots_adjust(left=0.03, right=0.58, top=0.82, bottom=0.12)
pay_order = sorted(payments, key=lambda p: payments[p]["revenue"], reverse=True)
pay_vals = [payments[p]["revenue"] for p in pay_order]
pay_colors = [TEAL, BLUE, ORANGE, "#7C3AED", GRAY]
wedges, _, autotexts = ax.pie(
    pay_vals,
    colors=pay_colors,
    startangle=90,
    counterclock=False,
    autopct=lambda p: f"{p:.1f}%",
    pctdistance=0.78,
    wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2),
    textprops=dict(color="white", fontsize=9, fontweight="bold"),
)
ax.text(0, 0, f"${TOTAL_REV/1_000_000:.2f}M\nrecorded", ha="center", va="center",
        fontsize=15, fontweight="bold", color=INK)
ax.legend(
    wedges,
    [f"{p}  ·  ${payments[p]['revenue']:,.0f}  ({payments[p]['orders']} orders)" for p in pay_order],
    loc="center left",
    bbox_to_anchor=(1.02, 0.5),
    frameon=False,
    fontsize=10,
)
titles(fig, "Revenue Is Evenly Spread Across Payment Methods",
       "Share of recorded revenue by payment method")
footer(fig, SOURCE)
save(fig, "03_payment_method_mix.png")

# ----------------------------------------------------------------------
# 04 - Orders by status
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7))
fig.subplots_adjust(left=0.16, right=0.94, top=0.82, bottom=0.14)
status_order = sorted(statuses, key=lambda s: statuses[s]["orders"])
status_names = status_order
status_counts = [statuses[s]["orders"] for s in status_order]
status_colors = [
    RED if s == "Cancelled" else ORANGE if s == "Returned" else TEAL if s == "Delivered" else BLUE
    for s in status_order
]
bars = ax.barh(status_names, status_counts, color=status_colors, height=0.66, alpha=0.92)
for bar, s in zip(bars, status_order):
    v = statuses[s]["orders"]
    pct = v / N * 100
    ax.text(v + 3, bar.get_y() + bar.get_height() / 2,
            f"{v}  ({pct:.1f}%)", va="center", ha="left", fontsize=10, color=INK, fontweight="bold")
ax.set_xlim(0, max(status_counts) * 1.25)
ax.set_xlabel("Number of orders")
ax.set_yticks(range(len(status_names)), status_names, fontsize=11)
clean_ax(ax, "x")
titles(fig, "Cancelled Is the Most Common Order Status",
       f"Order count by status  ·  only {delivered_orders} orders "
       f"({delivered_orders/N*100:.2f}%) reached Delivered")
footer(fig, SOURCE)
save(fig, "04_orders_by_status.png")

# ----------------------------------------------------------------------
# 05 - Referral source: orders and revenue
# ----------------------------------------------------------------------
fig, axes = plt.subplots(1, 2, figsize=(13, 7), sharey=True)
fig.subplots_adjust(left=0.17, right=0.95, top=0.80, bottom=0.14, wspace=0.28)
ref_order = sorted(referrals, key=lambda r: referrals[r]["orders"])
ref_names = ref_order
ref_counts = [referrals[r]["orders"] for r in ref_order]
ref_revs = [referrals[r]["revenue"] for r in ref_order]
ref_colors = [TEAL if r == "Instagram" else BLUE for r in ref_order]

ax = axes[0]
bars = ax.barh(ref_names, ref_counts, color=ref_colors, height=0.62, alpha=0.92)
for bar, r in zip(bars, ref_order):
    v = referrals[r]["orders"]
    ax.text(v + 3, bar.get_y() + bar.get_height() / 2, f"{v}", va="center",
            fontsize=10, color=INK, fontweight="bold")
ax.set_xlim(0, max(ref_counts) * 1.22)
ax.set_xlabel("Orders")
ax.set_title("Orders", fontsize=12, color=INK, loc="left", pad=8)
ax.set_yticks(range(len(ref_names)), ref_names, fontsize=10)
clean_ax(ax, "x")

ax = axes[1]
bars = ax.barh(ref_names, ref_revs, color=ref_colors, height=0.62, alpha=0.92)
for bar, r in zip(bars, ref_order):
    v = referrals[r]["revenue"]
    ax.text(v + max(ref_revs) * 0.02, bar.get_y() + bar.get_height() / 2,
            f"${v:,.0f}", va="center", fontsize=9, color=INK, fontweight="bold")
ax.set_xlim(0, max(ref_revs) * 1.18)
ax.set_xticks([0, 100000, 200000, 300000])
ax.xaxis.set_major_formatter(FuncFormatter(lambda x, p: "$0" if x == 0 else f"${x/1000:.0f}k"))
ax.set_xlabel("Recorded revenue (USD)")
ax.set_title("Recorded revenue", fontsize=12, color=INK, loc="left", pad=8)
clean_ax(ax, "x")

titles(fig, "Instagram Is the Strongest Acquisition Source",
       "Orders and recorded revenue by referral source")
fig.text(0.17, 0.055,
         "Facebook has the highest average order value ($1,098.29) despite lower order volume.",
         fontsize=8.5, color=MUTED)
footer(fig, SOURCE)
save(fig, "05_referral_source.png")

# ----------------------------------------------------------------------
# 06 - Yearly revenue
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7))
fig.subplots_adjust(left=0.10, right=0.94, top=0.80, bottom=0.16)
years = sorted(year_stats)
year_revs = [year_stats[y]["revenue"] for y in years]
year_colors = [TEAL, BLUE, GRAY]
bars = ax.bar([str(y) for y in years], year_revs, color=year_colors, width=0.55, alpha=0.92)
bars[2].set_hatch("//")
bars[2].set_edgecolor("white")
for bar, y in zip(bars, years):
    v = year_stats[y]["revenue"]
    ax.text(bar.get_x() + bar.get_width() / 2, v + max(year_revs) * 0.025,
            f"${v:,.0f}", ha="center", va="bottom", fontsize=11, color=INK, fontweight="bold")
    if y == 2025:
        ax.text(bar.get_x() + bar.get_width() / 2, v / 2,
                f"{year_stats[y]['orders']} orders", ha="center", va="center",
                fontsize=9, color=INK, fontweight="bold",
                bbox=dict(facecolor="white", alpha=0.85, edgecolor="none", pad=2))
    else:
        ax.text(bar.get_x() + bar.get_width() / 2, v / 2,
                f"{year_stats[y]['orders']} orders", ha="center", va="center",
                fontsize=9, color="white", fontweight="bold")
ax.set_xticks(range(len(years)), [f"{y}\n{'partial year' if y == 2025 else ''}" for y in years])
ax.set_ylim(0, max(year_revs) * 1.18)
ax.yaxis.set_major_formatter(usd)
ax.set_ylabel("Recorded revenue (USD)")
clean_ax(ax, "y")
ax.text(0.98, 0.95,
        "2023 to 2024\nRevenue: -13.1%\nOrders: -10.0%",
        transform=ax.transAxes, ha="right", va="top", fontsize=10, color=INK,
        bbox=dict(boxstyle="round,pad=0.6", facecolor=PALE, edgecolor=GRID))
titles(fig, "Full-Year Revenue Fell 13.1% from 2023 to 2024",
       "Recorded revenue by year  ·  2025 covers Jan-Jun only")
footer(fig, SOURCE)
save(fig, "06_yearly_revenue.png")

# ----------------------------------------------------------------------
# 07 - Customer loyalty
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7))
fig.subplots_adjust(left=0.05, right=0.62, top=0.82, bottom=0.12)
loyalty_vals = [one_time_customers, repeat_customers]
loyalty_colors = [BLUE, ORANGE]
wedges, _, autotexts = ax.pie(
    loyalty_vals,
    colors=loyalty_colors,
    startangle=90,
    counterclock=False,
    explode=(0, 0.10),
    autopct=lambda p: f"{p:.2f}%" if p > 50 else "",
    pctdistance=0.74,
    wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2),
    textprops=dict(color="white", fontsize=10, fontweight="bold"),
)
ax.text(0, 0, f"{unique_customers:,}\ncustomers", ha="center", va="center",
        fontsize=15, fontweight="bold", color=INK)
ax.legend(
    wedges,
    [f"1 order  ·  {one_time_customers:,} customers ({one_time_customers/unique_customers*100:.2f}%)",
     f"2 orders  ·  {repeat_customers} customers ({repeat_customers/unique_customers*100:.2f}%)"],
    loc="center left",
    bbox_to_anchor=(1.02, 0.5),
    frameon=False,
    fontsize=10,
)
titles(fig, "Repeat Purchasing Is Almost Absent",
       "Customers by number of orders")
footer(fig, SOURCE)
fig.text(0.06, 0.055,
         f"The {repeat_customers} repeat customers generated {repeat_orders} orders "
         f"({repeat_orders/N*100:.2f}% of all orders).",
         fontsize=8.5, color=MUTED)
save(fig, "07_customer_loyalty.png")

# ----------------------------------------------------------------------
# 08 - Order value distribution by product
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7))
fig.subplots_adjust(left=0.16, right=0.94, top=0.82, bottom=0.14)
box_order = sorted(order_values, key=lambda p: np.median(order_values[p]))
box_data = [order_values[p] for p in box_order]
box = ax.boxplot(
    box_data,
    orientation="horizontal",
    patch_artist=True,
    widths=0.6,
    showfliers=True,
    medianprops=dict(color=INK, linewidth=1.8),
    whiskerprops=dict(color=GRAY, linewidth=1),
    capprops=dict(color=GRAY, linewidth=1),
    flierprops=dict(marker="o", markersize=3, markerfacecolor=GRAY,
                    markeredgecolor="white", alpha=0.5),
)
box_colors = [TEAL if p == "Chair" else ORANGE if p == "Phone" else BLUE for p in box_order]
for patch, color in zip(box["boxes"], box_colors):
    patch.set_facecolor(color)
    patch.set_alpha(0.55)
    patch.set_edgecolor(color)
for i, p in enumerate(box_order, start=1):
    med = float(np.median(order_values[p]))
    ax.text(med, i + 0.30, f"${med:,.0f}", ha="center", va="bottom",
            fontsize=8.5, color=INK, fontweight="bold")
ax.axvline(AOV, color=ORANGE, linestyle="--", linewidth=1.5)
ax.text(AOV, len(box_order) + 0.72, f"Overall AOV  ${AOV:,.0f}", color=ORANGE,
        fontsize=9, ha="center", va="bottom", fontweight="bold")
ax.set_xlim(0, max(max(v) for v in order_values.values()) * 1.05)
ax.xaxis.set_major_formatter(usd)
ax.set_xlabel("Recorded order value (USD)")
ax.set_yticks(range(1, len(box_order) + 1), box_order, fontsize=10)
clean_ax(ax, "x")
titles(fig, "Order Values Overlap Heavily Across Products",
       "Distribution of recorded order value by product  ·  box = IQR, dots = outliers")
footer(fig, SOURCE)
save(fig, "08_order_value_distribution.png")

# ----------------------------------------------------------------------
# 09 - Recorded revenue to delivered revenue (waterfall)
# ----------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7))
fig.subplots_adjust(left=0.10, right=0.94, top=0.80, bottom=0.18)
wf_labels = ["Recorded\nrevenue", "Cancelled", "Returned", "Pending", "Shipped", "Delivered\nrevenue"]
wf_values = [
    TOTAL_REV,
    -statuses["Cancelled"]["revenue"],
    -statuses["Returned"]["revenue"],
    -statuses["Pending"]["revenue"],
    -statuses["Shipped"]["revenue"],
    delivered,
]
bottoms, heights, colors = [], [], []
run = 0.0
for i, v in enumerate(wf_values):
    if i == 0:
        bottoms.append(0)
        heights.append(TOTAL_REV)
        colors.append(GRAY)
        run = TOTAL_REV
    elif i == len(wf_values) - 1:
        bottoms.append(0)
        heights.append(delivered)
        colors.append(TEAL)
    else:
        bottoms.append(run + v)
        heights.append(-v)
        colors.append(RED if wf_labels[i] in ("Cancelled", "Returned") else ORANGE)
        run += v

x = np.arange(len(wf_labels))
bars = ax.bar(x, heights, bottom=bottoms, color=colors, width=0.62, alpha=0.92)
for i, (bar, bot, h, v) in enumerate(zip(bars, bottoms, heights, wf_values)):
    if i == 0:
        label = f"${v:,.0f}"
        y = bot + h
        ax.text(bar.get_x() + bar.get_width() / 2, y + TOTAL_REV * 0.018, label,
                ha="center", va="bottom", fontsize=10, fontweight="bold", color=INK)
    elif i == len(wf_values) - 1:
        ax.text(bar.get_x() + bar.get_width() / 2, h + TOTAL_REV * 0.018, f"${v:,.0f}",
                ha="center", va="bottom", fontsize=10, fontweight="bold", color=TEAL)
        ax.text(bar.get_x() + bar.get_width() / 2, h / 2, f"{v/TOTAL_REV*100:.1f}%\nof recorded",
                ha="center", va="center", fontsize=9, fontweight="bold", color="white")
    else:
        ax.text(bar.get_x() + bar.get_width() / 2, bot + h + TOTAL_REV * 0.018,
                f"-${-v:,.0f}", ha="center", va="bottom", fontsize=9, fontweight="bold", color=INK)

for i in range(len(wf_labels) - 1):
    top = bottoms[i] + heights[i] if i > 0 else TOTAL_REV
    ax.plot([x[i] + 0.31, x[i + 1] - 0.31], [top, top], color=GRAY,
            linestyle="--", linewidth=0.8, alpha=0.8)

ax.set_xticks(x, wf_labels, fontsize=9.5)
ax.set_ylim(0, TOTAL_REV * 1.12)
ax.yaxis.set_major_formatter(usd)
ax.set_ylabel("Recorded revenue (USD)")
clean_ax(ax, "y")
ax.text(0.98, 0.95,
        f"${not_delivered:,.0f}\n({not_delivered/TOTAL_REV*100:.1f}%) not delivered",
        transform=ax.transAxes, ha="right", va="top", fontsize=10, color=INK,
        bbox=dict(boxstyle="round,pad=0.6", facecolor="#FEF2F2", edgecolor="#FECACA"))
titles(fig, "Only 19.2% of Recorded Revenue Was Delivered",
       "From recorded order value to delivered revenue  ·  waterfall of order statuses")
footer(fig, SOURCE)
save(fig, "09_delivered_revenue_waterfall.png")

# ----------------------------------------------------------------------
# 10 - Dashboard (composed with PIL from four charts)
# ----------------------------------------------------------------------
def font(size, bold=False):
    candidates = [
        r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def fit_image(img, w, h):
    ratio = max(w / img.width, h / img.height)
    img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.LANCZOS)
    left = (img.width - w) // 2
    top = (img.height - h) // 2
    return img.crop((left, top, left + w, top + h))


W, H = 2400, 1750
header_h, footer_h, margin, gap = 230, 70, 40, 30
cell_w = (W - 2 * margin - gap) // 2
cell_h = (H - header_h - footer_h - gap - margin) // 2

canvas = Image.new("RGB", (W, H), "white")
draw = ImageDraw.Draw(canvas)

draw.rectangle([0, 0, W, 10], fill=TEAL)
draw.text((margin, 38), "E-commerce Data Visualisation", font=font(54, True), fill=INK)
draw.text((margin + 2, 108),
          "Project 4  ·  Segun Akinola  ·  1,200 orders  ·  Jan 2023 - Jun 2025",
          font=font(26), fill=MUTED)

kpis = [
    "1,200 orders",
    f"${TOTAL_REV/1_000_000:.2f}M recorded",
    f"${AOV:,.2f} AOV",
    f"{delivered_orders/N*100:.2f}% delivered",
]
chip_x = margin
chip_y = 162
chip_h = 46
chip_font = font(22, True)
for text in kpis:
    bbox = draw.textbbox((0, 0), text, font=chip_font)
    tw = bbox[2] - bbox[0]
    cw = tw + 36
    draw.rounded_rectangle([chip_x, chip_y, chip_x + cw, chip_y + chip_h],
                           radius=12, fill=PALE, outline=GRID, width=1)
    draw.text((chip_x + 18, chip_y + 10), text, font=chip_font, fill=INK)
    chip_x += cw + 16

panels = [
    ("01_revenue_by_product.png", margin, header_h),
    ("02_monthly_revenue_trend.png", margin + cell_w + gap, header_h),
    ("04_orders_by_status.png", margin, header_h + cell_h + gap),
    ("09_delivered_revenue_waterfall.png", margin + cell_w + gap, header_h + cell_h + gap),
]
for name, px, py in panels:
    img = Image.open(os.path.join(CHART_DIR, name)).convert("RGB")
    canvas.paste(fit_image(img, cell_w, cell_h), (px, py))

draw.line([margin, H - footer_h + 10, W - margin, H - footer_h + 10], fill=GRID, width=2)
draw.text((margin, H - footer_h + 24),
          "Source: Dataset for Data Analytics (1).xlsx  |  Charts generated with Python (matplotlib)",
          font=font(20), fill=MUTED)

dashboard_path = os.path.join(CHART_DIR, "00_dashboard.png")
canvas.save(dashboard_path, quality=95)
print("saved 00_dashboard.png")

# ----------------------------------------------------------------------
# Chart data appendix
# ----------------------------------------------------------------------
csv_path = os.path.join(BASE, "chart_data.csv")
with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["chart", "category", "orders", "units", "revenue",
                "avg_order_value", "share_pct"])

    for p in sorted(products, key=lambda k: products[k]["revenue"], reverse=True):
        d = products[p]
        w.writerow(["revenue_by_product", p, d["orders"], d["units"],
                    f"{d['revenue']:.2f}", f"{d['aov']:.2f}", f"{d['share']:.2f}"])

    for k in month_keys:
        rs = monthly[k]
        rev = sum(x["TotalPrice"] for x in rs)
        w.writerow(["monthly_revenue", f"{k[0]}-{k[1]:02d}", len(rs), "",
                    f"{rev:.2f}", "", ""])

    for p in sorted(payments, key=lambda k: payments[k]["revenue"], reverse=True):
        d = payments[p]
        w.writerow(["payment_mix", p, d["orders"], d["units"],
                    f"{d['revenue']:.2f}", f"{d['aov']:.2f}", f"{d['share']:.2f}"])

    for s in sorted(statuses, key=lambda k: statuses[k]["orders"], reverse=True):
        d = statuses[s]
        w.writerow(["order_status", s, d["orders"], d["units"],
                    f"{d['revenue']:.2f}", f"{d['aov']:.2f}", f"{d['share']:.2f}"])

    for r in sorted(referrals, key=lambda k: referrals[k]["orders"], reverse=True):
        d = referrals[r]
        w.writerow(["referral_source", r, d["orders"], d["units"],
                    f"{d['revenue']:.2f}", f"{d['aov']:.2f}", f"{d['share']:.2f}"])

    for y in years:
        d = year_stats[y]
        w.writerow(["yearly_revenue", str(y), d["orders"], d["units"],
                    f"{d['revenue']:.2f}", f"{d['aov']:.2f}", f"{d['share']:.2f}"])

print(f"wrote {csv_path}")

# ----------------------------------------------------------------------
# Console verification summary
# ----------------------------------------------------------------------
print("\nVERIFIED SUMMARY")
print(f"records={N} total_revenue={TOTAL_REV:.2f} aov={AOV:.2f}")
print(f"unique_customers={unique_customers} repeat_customers={repeat_customers} "
      f"one_time_customers={one_time_customers}")
print(f"delivered_revenue={delivered:.2f} delivered_pct={delivered/TOTAL_REV*100:.2f}")
print(f"top_product={max(products, key=lambda p: products[p]['revenue'])}")
print(f"top_referral={max(referrals, key=lambda r: referrals[r]['orders'])}")
