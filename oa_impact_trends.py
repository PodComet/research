"""
OA flavor citation impact trends over time (2018–2024).
Queries average citations per work by OA flavor × year, for:
  - All US academic institutions (education type)
  - 8 international organizations
Generates: 3 charts + 1 spreadsheet
"""
import requests, socket, sys, os, time
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import urllib3.util.connection as u
u.allowed_gai_family = lambda: socket.AF_INET
sys.stdout.reconfigure(encoding='utf-8')

API_KEY = "eN9RGcYK6NQuQJHOJ2XBDX"
BASE = "https://api.openalex.org"
OUTPUT_DIR = "C:/Users/talay/Documents/openalex_project/charts_2018"
XLSX_PATH = "C:/Users/talay/Documents/openalex_project/OpenAlex_OA_Impact_Trends.xlsx"
os.makedirs(OUTPUT_DIR, exist_ok=True)

YEARS = list(range(2018, 2025))  # 2018–2024 (2025+ too fresh for citations)
OA_FLAVORS = ["gold", "green", "hybrid", "bronze", "diamond", "closed"]
OA_COLORS = {
    "gold": "#F59E0B", "green": "#10B981", "hybrid": "#8B5CF6",
    "bronze": "#F97316", "diamond": "#06B6D4", "closed": "#9CA3AF",
}
OA_LABELS = {
    "gold": "Gold", "green": "Green", "hybrid": "Hybrid",
    "bronze": "Bronze", "diamond": "Diamond", "closed": "Closed",
}
OA_MARKERS = {
    "gold": "o", "green": "s", "hybrid": "D",
    "bronze": "^", "diamond": "P", "closed": "X",
}

plt.rcParams.update({
    'figure.facecolor': '#f8f9fa', 'axes.facecolor': '#ffffff',
    'axes.grid': True, 'grid.alpha': 0.3, 'font.size': 11,
    'axes.titlesize': 14, 'axes.titleweight': 'bold',
    'figure.titlesize': 16, 'figure.titleweight': 'bold',
})

H_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=11)
H_FILL = PatternFill('solid', fgColor='2563EB')
D_FONT = Font(name='Arial', size=10)
B_FONT = Font(name='Arial', bold=True, size=10)
L_FONT = Font(name='Arial', size=10, color='2563EB', underline='single')
NUM = '#,##0'; DEC = '0.0'
BDR = Border(bottom=Side(style='thin', color='D1D5DB'), right=Side(style='thin', color='D1D5DB'))
CTR = Alignment(horizontal='center', vertical='center')
LFT = Alignment(horizontal='left', vertical='center')
WRP = Alignment(horizontal='center', vertical='center', wrap_text=True)


def api(endpoint, params=None, retries=3):
    if params is None: params = {}
    params["api_key"] = API_KEY
    for attempt in range(retries):
        try:
            r = requests.get(f"{BASE}/{endpoint}", params=params, timeout=30)
            r.raise_for_status(); return r.json()
        except Exception as e:
            if attempt < retries - 1: time.sleep(1)
            else: print(f"      [!] {e}"); return {}


def save_fig(fig, fn):
    p = f"{OUTPUT_DIR}/{fn}"
    fig.savefig(p, dpi=150, bbox_inches='tight', facecolor=fig.get_facecolor())
    plt.close(fig); print(f"    Saved: {p}")


def style_header(ws, row, mc):
    for c in range(1, mc+1):
        cell = ws.cell(row=row, column=c)
        cell.font = H_FONT; cell.fill = H_FILL; cell.alignment = WRP; cell.border = BDR


def auto_width(ws, mn=10, mx=38):
    for col_cells in ws.columns:
        cl = get_column_letter(col_cells[0].column)
        w = max((len(str(c.value or '')) for c in col_cells), default=mn)
        ws.column_dimensions[cl].width = min(max(w+2, mn), mx)


def sample_impact(filt, pages=2):
    """Fetch up to pages×100 works and return (avg_citations, sample_size, total_count)."""
    all_cites = []
    total_count = 0
    for pg in range(1, pages + 1):
        data = api("works", {
            "filter": filt,
            "per_page": 100,
            "page": pg,
            "select": "id,cited_by_count",
        })
        total_count = data.get("meta", {}).get("count", 0)
        for w in data.get("results", []):
            all_cites.append(w.get("cited_by_count", 0))
        if len(data.get("results", [])) < 100:
            break
    avg = round(sum(all_cites) / len(all_cites), 2) if all_cites else 0
    return avg, len(all_cites), total_count


# ==================================================================
print("\n" + "=" * 65)
print("  OA IMPACT TRENDS BY FLAVOR — 2018–2024")
print("=" * 65)

# ── 1. Get intl org IDs ──
INTL_ORGS = [
    "United Nations", "World Bank", "International Monetary Fund",
    "Food and Agriculture Organization", "International Labour Organization",
    "International Telecommunication Union", "World Intellectual Property Organization",
    "United Nations Development Programme",
]

print("\n  [1/4] Resolving intl org IDs...")
intl_ids = []
for org_name in INTL_ORGS:
    data = api("institutions", {"search": org_name, "per_page": 5})
    inst = data.get("results", [None])[0]
    for r in data.get("results", []):
        if r.get("type") in ("government", "nonprofit", "facility", "other"):
            inst = r; break
    if inst:
        iid = inst["id"].split("/")[-1]
        intl_ids.append(iid)
        print(f"    ✓ {inst['display_name']} → {iid}")
intl_id_filter = "|".join(intl_ids)

# ── 2. Fetch year × flavor for US academic ──
total_calls = len(YEARS) * len(OA_FLAVORS) * 2  # US + intl
print(f"\n  [2/4] Fetching US academic impact by year × flavor ({len(YEARS)}×{len(OA_FLAVORS)} = {len(YEARS)*len(OA_FLAVORS)} queries)...")

us_trends = {}  # {flavor: {year: avg_cite}}
for flavor in OA_FLAVORS:
    us_trends[flavor] = {}
    for year in YEARS:
        filt = (f"authorships.institutions.type:education,"
                f"authorships.institutions.country_code:US,"
                f"open_access.oa_status:{flavor},"
                f"publication_year:{year}")
        avg, n, total = sample_impact(filt, pages=2)
        us_trends[flavor][year] = {"avg": avg, "n": n, "total": total}
        print(f"    US {flavor:8s} {year}: avg={avg:8.1f}  (sample={n}, total={total:,})")
        time.sleep(0.02)

# ── 3. Fetch year × flavor for intl orgs ──
print(f"\n  [3/4] Fetching intl org impact by year × flavor...")

intl_trends = {}
for flavor in OA_FLAVORS:
    intl_trends[flavor] = {}
    for year in YEARS:
        filt = (f"authorships.institutions.id:{intl_id_filter},"
                f"open_access.oa_status:{flavor},"
                f"publication_year:{year}")
        avg, n, total = sample_impact(filt, pages=2)
        intl_trends[flavor][year] = {"avg": avg, "n": n, "total": total}
        print(f"    Intl {flavor:8s} {year}: avg={avg:8.1f}  (sample={n}, total={total:,})")
        time.sleep(0.02)

# ── 4. Charts & Spreadsheet ──
print(f"\n  [4/4] Generating charts & spreadsheet...")

# --- Chart A: US Academic — Impact trends by flavor ---
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
fig.suptitle("Citation Impact by OA Flavor Over Time — 2018–2024")

ax1.set_title("US Academic Institutions")
for flavor in OA_FLAVORS:
    vals = [us_trends[flavor][y]["avg"] for y in YEARS]
    ax1.plot(YEARS, vals, color=OA_COLORS[flavor], marker=OA_MARKERS[flavor],
             linewidth=2.5, markersize=8, label=OA_LABELS[flavor], zorder=3)
ax1.set_xlabel("Publication Year"); ax1.set_ylabel("Avg Citations per Work (sampled)")
ax1.legend(fontsize=10, loc='upper left')
ax1.set_xticks(YEARS)
ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))

ax2.set_title("International Organizations")
for flavor in OA_FLAVORS:
    vals = [intl_trends[flavor][y]["avg"] for y in YEARS]
    ax2.plot(YEARS, vals, color=OA_COLORS[flavor], marker=OA_MARKERS[flavor],
             linewidth=2.5, markersize=8, label=OA_LABELS[flavor], zorder=3)
ax2.set_xlabel("Publication Year"); ax2.set_ylabel("Avg Citations per Work (sampled)")
ax2.legend(fontsize=10, loc='upper left')
ax2.set_xticks(YEARS)

fig.tight_layout()
save_fig(fig, "16_oa_impact_trends_sidebyside.png")

# --- Chart B: Normalized index (2018=100) for US ---
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
fig.suptitle("OA Flavor Citation Impact — Indexed to 2018 (2018 = 100)")

ax1.set_title("US Academic Institutions")
for flavor in OA_FLAVORS:
    base = us_trends[flavor][2018]["avg"]
    if base > 0:
        vals = [100 * us_trends[flavor][y]["avg"] / base for y in YEARS]
        ax1.plot(YEARS, vals, color=OA_COLORS[flavor], marker=OA_MARKERS[flavor],
                 linewidth=2.5, markersize=8, label=OA_LABELS[flavor], zorder=3)
ax1.axhline(y=100, color='#64748b', linewidth=1, linestyle='--', alpha=0.5)
ax1.set_xlabel("Publication Year"); ax1.set_ylabel("Index (2018 = 100)")
ax1.legend(fontsize=10, loc='upper right')
ax1.set_xticks(YEARS)

ax2.set_title("International Organizations")
for flavor in OA_FLAVORS:
    base = intl_trends[flavor][2018]["avg"]
    if base > 0:
        vals = [100 * intl_trends[flavor][y]["avg"] / base for y in YEARS]
        ax2.plot(YEARS, vals, color=OA_COLORS[flavor], marker=OA_MARKERS[flavor],
                 linewidth=2.5, markersize=8, label=OA_LABELS[flavor], zorder=3)
ax2.axhline(y=100, color='#64748b', linewidth=1, linestyle='--', alpha=0.5)
ax2.set_xlabel("Publication Year"); ax2.set_ylabel("Index (2018 = 100)")
ax2.legend(fontsize=10, loc='upper right')
ax2.set_xticks(YEARS)

fig.tight_layout()
save_fig(fig, "17_oa_impact_indexed_trends.png")

# --- Chart C: Stacked area — publication volume by flavor over time (US) ---
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
fig.suptitle("Publication Volume by OA Flavor Over Time — 2018–2024")

ax1.set_title("US Academic — Works per Year by OA Flavor")
ys = np.array(YEARS)
stacks_us = np.array([[us_trends[fl][y]["total"] for y in YEARS] for fl in OA_FLAVORS])
ax1.stackplot(ys, stacks_us, labels=[OA_LABELS[fl] for fl in OA_FLAVORS],
              colors=[OA_COLORS[fl] for fl in OA_FLAVORS], alpha=0.85)
ax1.set_xlabel("Publication Year"); ax1.set_ylabel("Number of Works")
ax1.legend(fontsize=9, loc='upper left')
ax1.set_xticks(YEARS)
ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x/1e6:.1f}M' if x >= 1e6 else f'{x/1e3:.0f}K'))

ax2.set_title("Intl Organizations — Works per Year by OA Flavor")
stacks_intl = np.array([[intl_trends[fl][y]["total"] for y in YEARS] for fl in OA_FLAVORS])
ax2.stackplot(ys, stacks_intl, labels=[OA_LABELS[fl] for fl in OA_FLAVORS],
              colors=[OA_COLORS[fl] for fl in OA_FLAVORS], alpha=0.85)
ax2.set_xlabel("Publication Year"); ax2.set_ylabel("Number of Works")
ax2.legend(fontsize=9, loc='upper left')
ax2.set_xticks(YEARS)
ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x/1e3:.1f}K'))

fig.tight_layout()
save_fig(fig, "18_oa_volume_stacked_trends.png")

# --- Spreadsheet ---
print("\n    Building spreadsheet...")
wb = Workbook()

# Sheet 1: US Academic detail
ws1 = wb.active
ws1.title = "US Academic Trends"
headers1 = ["OA Flavor", "Year", "Avg Citations/Work", "Sample Size", "Total Works"]
ws1.append(headers1); style_header(ws1, 1, len(headers1))
for fl in OA_FLAVORS:
    for y in YEARS:
        r = ws1.max_row + 1
        d = us_trends[fl][y]
        ws1.cell(r, 1, OA_LABELS[fl]); ws1.cell(r, 1).font = D_FONT
        ws1.cell(r, 2, y)
        ws1.cell(r, 3, d["avg"]); ws1.cell(r, 3).number_format = DEC
        ws1.cell(r, 4, d["n"]); ws1.cell(r, 4).number_format = NUM
        ws1.cell(r, 5, d["total"]); ws1.cell(r, 5).number_format = NUM
        for c in range(1, 6):
            ws1.cell(r, c).border = BDR
            ws1.cell(r, c).alignment = CTR if c > 1 else LFT
auto_width(ws1)

# Sheet 2: Intl Orgs detail
ws2 = wb.create_sheet("Intl Org Trends")
ws2.append(headers1); style_header(ws2, 1, len(headers1))
for fl in OA_FLAVORS:
    for y in YEARS:
        r = ws2.max_row + 1
        d = intl_trends[fl][y]
        ws2.cell(r, 1, OA_LABELS[fl]); ws2.cell(r, 1).font = D_FONT
        ws2.cell(r, 2, y)
        ws2.cell(r, 3, d["avg"]); ws2.cell(r, 3).number_format = DEC
        ws2.cell(r, 4, d["n"]); ws2.cell(r, 4).number_format = NUM
        ws2.cell(r, 5, d["total"]); ws2.cell(r, 5).number_format = NUM
        for c in range(1, 6):
            ws2.cell(r, c).border = BDR
            ws2.cell(r, c).alignment = CTR if c > 1 else LFT
auto_width(ws2)

# Sheet 3: Summary pivot — avg citations by flavor × year
ws3 = wb.create_sheet("Summary Pivot")
headers3 = ["OA Flavor", "Group"] + [str(y) for y in YEARS]
ws3.append(headers3); style_header(ws3, 1, len(headers3))
for fl in OA_FLAVORS:
    # US row
    r = ws3.max_row + 1
    ws3.cell(r, 1, OA_LABELS[fl]); ws3.cell(r, 1).font = B_FONT
    ws3.cell(r, 2, "US Academic")
    for ci, y in enumerate(YEARS):
        ws3.cell(r, 3+ci, us_trends[fl][y]["avg"]); ws3.cell(r, 3+ci).number_format = DEC
    for c in range(1, len(headers3)+1):
        ws3.cell(r, c).border = BDR; ws3.cell(r, c).alignment = CTR if c > 1 else LFT
    # Intl row
    r = ws3.max_row + 1
    ws3.cell(r, 1, OA_LABELS[fl]); ws3.cell(r, 1).font = D_FONT
    ws3.cell(r, 2, "Intl Orgs")
    for ci, y in enumerate(YEARS):
        ws3.cell(r, 3+ci, intl_trends[fl][y]["avg"]); ws3.cell(r, 3+ci).number_format = DEC
    for c in range(1, len(headers3)+1):
        ws3.cell(r, c).border = BDR; ws3.cell(r, c).alignment = CTR if c > 1 else LFT
auto_width(ws3)

wb.save(XLSX_PATH)
print(f"    Saved: {XLSX_PATH}")

# Print summary table
print("\n    ┌──────────────────────────────────────────────────────────────────────────────────┐")
print("    │   US Academic — Avg Citations/Work by OA Flavor × Year                          │")
print("    ├──────────┬──────────┬──────────┬──────────┬──────────┬──────────┬──────────┬─────┤")
hdr = "    │ Flavor   │" + "│".join(f"  {y}  " for y in YEARS) + "│"
print(hdr)
print("    ├──────────┼──────────┼──────────┼──────────┼──────────┼──────────┼──────────┼─────┤")
for fl in OA_FLAVORS:
    vals = "│".join(f" {us_trends[fl][y]['avg']:7.1f} " for y in YEARS)
    print(f"    │ {OA_LABELS[fl]:8s} │{vals}│")
print("    └──────────┴──────────┴──────────┴──────────┴──────────┴──────────┴──────────┴─────┘")

print("\n" + "=" * 65)
print("  DONE — 3 charts + spreadsheet")
print(f"  Charts: {OUTPUT_DIR}")
print(f"  Spreadsheet: {XLSX_PATH}")
print("=" * 65)
