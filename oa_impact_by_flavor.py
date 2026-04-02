"""
Average citation impact per OA flavor for all 58 institutions (8 intl + 50 US).
Scope: 2018–2026.  Method: sample 50 works per institution×flavor, compute mean cited_by_count.
Generates: 3 charts + 1 spreadsheet
"""
import requests, socket, sys, os, time
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import urllib3.util.connection as u
u.allowed_gai_family = lambda: socket.AF_INET
sys.stdout.reconfigure(encoding='utf-8')

API_KEY = "eN9RGcYK6NQuQJHOJ2XBDX"
BASE = "https://api.openalex.org"
YEAR_FILTER = "from_publication_date:2018-01-01"
OUTPUT_DIR = "C:/Users/talay/Documents/openalex_project/charts_2018"
XLSX_PATH = "C:/Users/talay/Documents/openalex_project/OpenAlex_OA_Impact_by_Flavor.xlsx"
os.makedirs(OUTPUT_DIR, exist_ok=True)

OA_FLAVORS = ["gold", "green", "hybrid", "bronze", "diamond", "closed"]
OA_COLORS = {
    "gold": "#F59E0B", "green": "#10B981", "hybrid": "#8B5CF6",
    "bronze": "#F97316", "diamond": "#06B6D4", "closed": "#9CA3AF",
}
OA_LABELS = {
    "gold": "Gold", "green": "Green", "hybrid": "Hybrid",
    "bronze": "Bronze", "diamond": "Diamond", "closed": "Closed",
}

plt.rcParams.update({
    'figure.facecolor': '#f8f9fa', 'axes.facecolor': '#ffffff',
    'axes.grid': True, 'grid.alpha': 0.3, 'font.size': 11,
    'axes.titlesize': 13, 'axes.titleweight': 'bold',
    'figure.titlesize': 15, 'figure.titleweight': 'bold',
})

H_FONT = Font(name='Arial', bold=True, color='FFFFFF', size=11)
H_FILL = PatternFill('solid', fgColor='2563EB')
D_FONT = Font(name='Arial', size=10)
B_FONT = Font(name='Arial', bold=True, size=10)
L_FONT = Font(name='Arial', size=10, color='2563EB', underline='single')
NUM = '#,##0'; DEC = '0.0'; PCT = '0.0%'
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


def shorten(name, mx=28):
    reps = [
        ("University of California, ", "UC "), ("University of California ", "UC "),
        ("The University of Texas at ", "UT "), ("University of North Carolina at ", "UNC "),
        ("University of Illinois Urbana-Champaign", "UIUC"),
        ("University of Illinois Chicago", "UIC"),
        ("Rutgers, The State University of New Jersey", "Rutgers"),
        ("Purdue University West Lafayette", "Purdue"),
        ("Indiana University Bloomington", "Indiana U"),
        ("Indiana University – Purdue University Indianapolis", "IUPUI"),
        ("Washington University in St. Louis", "WashU St. Louis"),
        ("Massachusetts Institute of Technology", "MIT"),
        ("Massachusetts General Hospital", "Mass General"),
        ("The Ohio State University", "Ohio State"),
        ("Pennsylvania State University", "Penn State"),
        ("Michigan State University", "Michigan State"),
        ("George Washington University", "George Washington U"),
        ("Texas A&M University", "Texas A&M"),
        ("International Monetary Fund", "IMF"),
        ("Food and Agriculture Organization of the United Nations", "FAO"),
        ("Food and Agriculture Organization", "FAO"),
        ("International Labour Organization", "ILO"),
        ("International Telecommunication Union", "ITU"),
        ("World Intellectual Property Organization", "WIPO"),
        ("United Nations Development Programme", "UNDP"),
        ("United Nations", "UN"), ("World Bank", "World Bank"),
        ("University of ", "U "), ("University", "U"),
    ]
    for old, new in reps:
        if old in name: name = name.replace(old, new); break
    return name[:mx-2] + ".." if len(name) > mx else name


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


# ==================================================================
print("\n" + "=" * 65)
print("  OA IMPACT BY FLAVOR — ALL INSTITUTIONS — 2018+")
print("=" * 65)

# ── 1. Get institution lists ──
INTL_ORGS = [
    "United Nations", "World Bank", "International Monetary Fund",
    "Food and Agriculture Organization", "International Labour Organization",
    "International Telecommunication Union", "World Intellectual Property Organization",
    "United Nations Development Programme",
]

print("\n  [1/3] Building institution list...")
intl = []
for org_name in INTL_ORGS:
    data = api("institutions", {"search": org_name, "per_page": 5})
    inst = data.get("results", [None])[0]
    for r in data.get("results", []):
        if r.get("type") in ("government", "nonprofit", "facility", "other"):
            inst = r; break
    if not inst: continue
    intl.append({"id": inst["id"].split("/")[-1], "name": inst["display_name"], "group": "intl"})
    print(f"    ✓ {inst['display_name']}")

ss_raw = api("works", {
    "filter": f"concepts.id:C36289849,authorships.institutions.country_code:US,"
              f"authorships.institutions.type:education,{YEAR_FILTER}",
    "group_by": "authorships.institutions.id", "per_page": 50,
})
us = []
for g in ss_raw.get("group_by", []):
    us.append({"id": g["key"].split("/")[-1], "name": g["key_display_name"], "group": "us"})
print(f"    ✓ {len(us)} US institutions loaded")

all_insts = intl + us

# ── 2. Fetch impact per OA flavor ──
print(f"\n  [2/3] Fetching impact by OA flavor for {len(all_insts)} institutions...")
print(f"        ({len(all_insts)} × {len(OA_FLAVORS)} = {len(all_insts)*len(OA_FLAVORS)} API calls)\n")

results = []
for i, inst in enumerate(all_insts, 1):
    row = {"id": inst["id"], "name": inst["name"], "group": inst["group"]}
    short = shorten(inst["name"], 30)
    flavors_done = []
    for flavor in OA_FLAVORS:
        filt = (f"authorships.institutions.id:{inst['id']},"
                f"open_access.oa_status:{flavor},{YEAR_FILTER},"
                f"to_publication_date:2024-12-31")
        # Get total count first
        data0 = api("works", {"filter": filt, "per_page": 1})
        total_count = data0.get("meta", {}).get("count", 0)
        # Sample works sorted by cited_by_count (gives citation-representative view)
        # We take 200 works with default sort (relevance) for a balanced sample
        all_cites = []
        for pg in [1, 2]:
            data = api("works", {
                "filter": filt,
                "per_page": 100,
                "page": pg,
                "select": "id,cited_by_count",
            })
            for w in data.get("results", []):
                all_cites.append(w.get("cited_by_count", 0))
            if len(data.get("results", [])) < 100:
                break
        if all_cites:
            avg_cite = sum(all_cites) / len(all_cites)
        else:
            avg_cite = 0
        row[f"{flavor}_count"] = total_count
        row[f"{flavor}_impact"] = round(avg_cite, 2)
        row[f"{flavor}_sample_n"] = len(all_cites)
        flavors_done.append(f"{flavor[:2]}={avg_cite:.1f}")
        time.sleep(0.02)

    print(f"    [{i:2d}/{len(all_insts)}] {short:30s}  {' | '.join(flavors_done)}")
    results.append(row)

# ── 3. Compute aggregates ──
print(f"\n  [3/3] Computing aggregates & generating outputs...")

# Global averages (weighted by sample, across all institutions)
global_avg = {}
intl_avg = {}
us_avg = {}
for flavor in OA_FLAVORS:
    # All institutions
    impacts = [r[f"{flavor}_impact"] for r in results if r[f"{flavor}_count"] > 0]
    global_avg[flavor] = round(np.mean(impacts), 2) if impacts else 0
    # Intl only
    impacts_i = [r[f"{flavor}_impact"] for r in results if r["group"] == "intl" and r[f"{flavor}_count"] > 0]
    intl_avg[flavor] = round(np.mean(impacts_i), 2) if impacts_i else 0
    # US only
    impacts_u = [r[f"{flavor}_impact"] for r in results if r["group"] == "us" and r[f"{flavor}_count"] > 0]
    us_avg[flavor] = round(np.mean(impacts_u), 2) if impacts_u else 0

print("\n    ┌─────────────────────────────────────────────────────────┐")
print("    │   Average Citations/Work by OA Flavor (2018+)          │")
print("    ├───────────┬──────────┬──────────┬──────────────────────┤")
print("    │  Flavor   │   All    │ Intl Org │   US Academic        │")
print("    ├───────────┼──────────┼──────────┼──────────────────────┤")
for fl in OA_FLAVORS:
    print(f"    │  {OA_LABELS[fl]:8s} │  {global_avg[fl]:6.2f}  │  {intl_avg[fl]:6.2f}  │  {us_avg[fl]:6.2f}               │")
print("    └───────────┴──────────┴──────────┴──────────────────────┘")

# --- Chart A: Global average impact by flavor (grouped bar: All / Intl / US) ---
fig, ax = plt.subplots(figsize=(12, 6))
fig.suptitle("Average Citation Impact by OA Flavor — 2018–2026")
x = np.arange(len(OA_FLAVORS))
w = 0.25
bars1 = ax.bar(x - w, [global_avg[f] for f in OA_FLAVORS], w, label='All Institutions',
               color='#1e3a5f', edgecolor='white', linewidth=0.5)
bars2 = ax.bar(x,     [intl_avg[f] for f in OA_FLAVORS], w, label='Intl Organizations',
               color='#2563eb', edgecolor='white', linewidth=0.5)
bars3 = ax.bar(x + w, [us_avg[f] for f in OA_FLAVORS], w, label='US Academic',
               color='#16a34a', edgecolor='white', linewidth=0.5)
ax.set_xticks(x)
ax.set_xticklabels([OA_LABELS[f] for f in OA_FLAVORS], fontsize=12)
ax.set_ylabel("Avg Citations per Work (sampled)", fontsize=11)
# Value labels
for bars in [bars1, bars2, bars3]:
    for bar in bars:
        h = bar.get_height()
        if h > 0:
            ax.text(bar.get_x() + bar.get_width()/2, h + 0.3,
                    f"{h:.1f}", ha='center', va='bottom', fontsize=9, fontweight='bold')
ax.legend(fontsize=11, loc='upper right')
ax.set_ylim(0, max(max(global_avg.values()), max(intl_avg.values()), max(us_avg.values())) * 1.2)
fig.tight_layout()
save_fig(fig, "13_oa_flavor_impact_global.png")

# --- Chart B: Heatmap — impact per flavor per institution (top 20 US + 8 intl) ---
subset = [r for r in results if r["group"] == "intl"] + \
         [r for r in results if r["group"] == "us"][:20]
names = [shorten(r["name"], 26) for r in subset]
data_matrix = np.array([[r[f"{fl}_impact"] for fl in OA_FLAVORS] for r in subset])

fig, ax = plt.subplots(figsize=(12, 14))
fig.suptitle("Citation Impact by OA Flavor — Top 28 Institutions (2018–2026)", y=0.995)
im = ax.imshow(data_matrix, cmap='YlOrRd', aspect='auto', interpolation='nearest')
ax.set_xticks(np.arange(len(OA_FLAVORS)))
ax.set_xticklabels([OA_LABELS[f] for f in OA_FLAVORS], fontsize=11, fontweight='bold')
ax.set_yticks(np.arange(len(names)))
ax.set_yticklabels(names, fontsize=9)
# Annotate cells
for i in range(len(subset)):
    for j in range(len(OA_FLAVORS)):
        v = data_matrix[i, j]
        if v > 0:
            color = 'white' if v > data_matrix.max() * 0.65 else '#1a1a1a'
            ax.text(j, i, f"{v:.1f}", ha='center', va='center', fontsize=8, color=color)
# Divider line between intl and US
n_intl = len([r for r in subset if r["group"] == "intl"])
ax.axhline(y=n_intl - 0.5, color='#2563eb', linewidth=2, linestyle='--')
ax.text(len(OA_FLAVORS)-1, n_intl - 0.75, "— Intl Orgs above | US Academic below —",
        ha='right', va='bottom', fontsize=8, color='#2563eb', fontstyle='italic')
cbar = fig.colorbar(im, ax=ax, fraction=0.02, pad=0.02)
cbar.set_label("Avg Citations per Work", fontsize=10)
fig.tight_layout()
save_fig(fig, "14_oa_flavor_impact_heatmap.png")

# --- Chart C: Dot plot — each institution's impact by flavor (top 10 US + intl) ---
subset2 = [r for r in results if r["group"] == "intl"] + \
          [r for r in results if r["group"] == "us"][:10]
fig, ax = plt.subplots(figsize=(14, 10))
fig.suptitle("Citation Impact by OA Flavor — Intl Orgs + Top 10 US (2018–2026)")
names2 = [shorten(r["name"], 24) for r in subset2]
y_pos = np.arange(len(names2))
markers = {'gold': 'o', 'green': 's', 'hybrid': 'D', 'bronze': '^', 'diamond': 'P', 'closed': 'X'}
for fl in OA_FLAVORS:
    vals = [r[f"{fl}_impact"] for r in subset2]
    ax.scatter(vals, y_pos, c=OA_COLORS[fl], marker=markers[fl],
               s=120, label=OA_LABELS[fl], edgecolors='white', linewidth=0.5, zorder=3)
ax.set_yticks(y_pos)
ax.set_yticklabels(names2, fontsize=10)
ax.set_xlabel("Avg Citations per Work (sampled)", fontsize=11)
ax.legend(fontsize=10, loc='lower right', ncol=2, framealpha=0.9)
ax.axhline(y=n_intl - 0.5, color='#94a3b8', linewidth=1, linestyle='--')
ax.invert_yaxis()
fig.tight_layout()
save_fig(fig, "15_oa_flavor_impact_dotplot.png")

# --- Spreadsheet ---
print("\n    Building spreadsheet...")
wb = Workbook()

# Sheet 1: Per-institution detail
ws1 = wb.active
ws1.title = "Impact by Flavor"
headers = ["Institution", "Type",
           "Gold Impact", "Green Impact", "Hybrid Impact",
           "Bronze Impact", "Diamond Impact", "Closed Impact",
           "Gold Works", "Green Works", "Hybrid Works",
           "Bronze Works", "Diamond Works", "Closed Works",
           "OpenAlex Link"]
ws1.append(headers); style_header(ws1, 1, len(headers))
for d in results:
    r = ws1.max_row + 1
    ws1.cell(r, 1, d["name"])
    ws1.cell(r, 2, "Intl Org" if d["group"] == "intl" else "US Academic")
    for ci, fl in enumerate(OA_FLAVORS):
        ws1.cell(r, 3+ci, d[f"{fl}_impact"]); ws1.cell(r, 3+ci).number_format = DEC
        ws1.cell(r, 9+ci, d[f"{fl}_count"]); ws1.cell(r, 9+ci).number_format = NUM
    link = f"https://api.openalex.org/works?filter=authorships.institutions.id:{d['id']},from_publication_date:2018-01-01&group_by=open_access.oa_status&api_key={API_KEY}"
    ws1.cell(r, 15, "API Link"); ws1.cell(r, 15).hyperlink = link
    ws1.cell(r, 15).font = L_FONT
    for c in range(1, len(headers)+1):
        ws1.cell(r, c).border = BDR
        ws1.cell(r, c).alignment = CTR if c > 1 else LFT
auto_width(ws1)

# Sheet 2: Summary averages
ws2 = wb.create_sheet("Summary Averages")
headers2 = ["OA Flavor", "Avg Impact (All)", "Avg Impact (Intl Orgs)", "Avg Impact (US Academic)"]
ws2.append(headers2); style_header(ws2, 1, len(headers2))
for fl in OA_FLAVORS:
    r = ws2.max_row + 1
    ws2.cell(r, 1, OA_LABELS[fl]); ws2.cell(r, 1).font = B_FONT
    ws2.cell(r, 2, global_avg[fl]); ws2.cell(r, 2).number_format = DEC
    ws2.cell(r, 3, intl_avg[fl]); ws2.cell(r, 3).number_format = DEC
    ws2.cell(r, 4, us_avg[fl]); ws2.cell(r, 4).number_format = DEC
    for c in range(1, 5):
        ws2.cell(r, c).border = BDR
        ws2.cell(r, c).alignment = CTR if c > 1 else LFT
auto_width(ws2)

wb.save(XLSX_PATH)
print(f"    Saved: {XLSX_PATH}")

print("\n" + "=" * 65)
print("  DONE — 3 charts + spreadsheet generated")
print(f"  Charts: {OUTPUT_DIR}")
print(f"  Spreadsheet: {XLSX_PATH}")
print("=" * 65)
