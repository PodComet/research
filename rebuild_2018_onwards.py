"""
Rebuild all data, spreadsheet, and visualizations for 2018–2026 only.
Sources: OpenAlex (institutions, OA, topics), ICPSR (via OpenAlex), Globus MDF.
"""
import requests, socket, sys, os, time, json
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
GLOBUS = "https://search.api.globus.org"
MDF_INDEX = "1a57bbe5-5272-477f-9d31-343b8258b7a5"
YEAR_FILTER = "from_publication_date:2018-01-01"
YEARS = list(range(2018, 2027))
OUTPUT_DIR = "C:/Users/talay/Documents/openalex_project/charts_2018"
XLSX_PATH = "C:/Users/talay/Documents/openalex_project/OpenAlex_2018_Onwards.xlsx"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Styles ──
plt.rcParams.update({
    'figure.facecolor': '#f8f9fa', 'axes.facecolor': '#ffffff',
    'axes.grid': True, 'grid.alpha': 0.3, 'font.size': 10,
    'axes.titlesize': 13, 'axes.titleweight': 'bold',
    'figure.titlesize': 15, 'figure.titleweight': 'bold',
})
C10 = ['#2563eb','#dc2626','#16a34a','#ea580c','#7c3aed',
       '#0891b2','#be185d','#4f46e5','#059669','#d97706']
C20 = C10 + ['#6366f1','#ef4444','#14b8a6','#f97316','#8b5cf6',
             '#06b6d4','#e11d48','#7c3aed','#10b981','#eab308']

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
            else: print(f"    [!] {e}"); return {}


def shorten(name, mx=28):
    reps = [
        ("University of California, ", "UC "), ("University of California ", "UC "),
        ("The University of Texas at ", "UT "), ("University of North Carolina at ", "UNC "),
        ("University of Illinois Urbana-Champaign", "UIUC"),
        ("University of Illinois Chicago", "UIC"),
        ("Rutgers, The State University of New Jersey", "Rutgers"),
        ("Purdue University West Lafayette", "Purdue"),
        ("Indiana University Bloomington", "Indiana U"),
        ("Washington University in St. Louis", "WashU St. Louis"),
        ("Massachusetts Institute of Technology", "MIT"),
        ("The Ohio State University", "Ohio State"),
        ("Pennsylvania State University", "Penn State"),
        ("Michigan State University", "Michigan State"),
        ("George Washington University", "George Washington U"),
        ("Texas A&M University", "Texas A&M"),
        ("International Monetary Fund", "IMF"),
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


def sh(ws, r, c, fmt=None, bold=False, link=False):
    cell = ws.cell(row=r, column=c)
    cell.font = L_FONT if link else (B_FONT if bold else D_FONT)
    cell.alignment = CTR if c > 1 else LFT
    cell.border = BDR
    if fmt: cell.number_format = fmt


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
print("  REBUILDING ALL DATA — 2018 ONWARDS")
print("=" * 65)

# ── 1. Intl Orgs ──
INTL_ORGS = [
    "United Nations", "World Bank", "International Monetary Fund",
    "Food and Agriculture Organization", "International Labour Organization",
    "International Telecommunication Union", "World Intellectual Property Organization",
    "United Nations Development Programme",
]

print("\n  [1/8] Intl organizations (2018+ works & citations)...")
intl = []
for org_name in INTL_ORGS:
    print(f"    -> {org_name}")
    data = api("institutions", {"search": org_name, "per_page": 5})
    inst = data.get("results", [None])[0]
    for r in data.get("results", []):
        if r.get("type") in ("government","nonprofit","facility","other"):
            inst = r; break
    if not inst: continue
    cby = inst.get("counts_by_year", [])
    yd = {e["year"]: e for e in cby}
    works_2018 = sum(yd.get(y,{}).get("works_count",0) for y in YEARS)
    cites_2018 = sum(yd.get(y,{}).get("cited_by_count",0) for y in YEARS)
    intl.append({
        "id": inst["id"].split("/")[-1], "name": inst["display_name"],
        "country": inst.get("country_code",""), "type": inst.get("type",""),
        "works": works_2018, "citations": cites_2018,
        "impact": round(cites_2018/works_2018,2) if works_2018 else 0,
        "h_index": inst.get("summary_stats",{}).get("h_index",0) or 0,
        "cby": cby,
    })

# ── 2. Top 50 US SS ──
print("\n  [2/8] Top 50 US social science institutions (2018+ SS works)...")
ss_raw = api("works", {
    "filter": f"concepts.id:C36289849,authorships.institutions.country_code:US,authorships.institutions.type:education,{YEAR_FILTER}",
    "group_by": "authorships.institutions.id", "per_page": 50,
})
ss_rankings = [(g["key_display_name"], g["count"]) for g in ss_raw.get("group_by",[])]

print(f"\n  [3/8] Fetching full profiles for {len(ss_rankings)} institutions...")
us = []
for name, ss_count in ss_rankings:
    print(f"    -> {name} ({ss_count:,})")
    data = api("institutions", {"search": name, "per_page": 1})
    inst = data.get("results",[None])[0]
    if not inst: continue
    cby = inst.get("counts_by_year",[])
    yd = {e["year"]: e for e in cby}
    works_2018 = sum(yd.get(y,{}).get("works_count",0) for y in YEARS)
    cites_2018 = sum(yd.get(y,{}).get("cited_by_count",0) for y in YEARS)
    us.append({
        "id": inst["id"].split("/")[-1], "name": inst["display_name"],
        "works": works_2018, "citations": cites_2018,
        "impact": round(cites_2018/works_2018,2) if works_2018 else 0,
        "h_index": inst.get("summary_stats",{}).get("h_index",0) or 0,
        "i10_index": inst.get("summary_stats",{}).get("i10_index",0) or 0,
        "ss_works": ss_count, "cby": cby,
    })
us_sorted = sorted(us, key=lambda x: x["ss_works"], reverse=True)
top20 = us_sorted[:20]

# ── 3. OA (top 20) ──
print("\n  [4/8] OA breakdowns (top 20, 2018+)...")
oa_map = {}
for p in top20:
    print(f"    -> {p['name']}")
    data = api("works", {
        "filter": f"authorships.institutions.id:{p['id']},{YEAR_FILTER}",
        "group_by": "open_access.oa_status"
    })
    oa_map[p["id"]] = {g["key_display_name"]: g["count"] for g in data.get("group_by",[])}

# ── 4. Topics (top 20) ──
print("\n  [5/8] Research topics (top 20, 2018+)...")
topic_map = {}
for p in top20:
    print(f"    -> {p['name']}")
    data = api("works", {
        "filter": f"authorships.institutions.id:{p['id']},{YEAR_FILTER}",
        "group_by": "topics.id", "per_page": 8
    })
    topic_map[p["id"]] = [(g["key_display_name"], g["count"]) for g in data.get("group_by",[])[:8]]

# ── 5. ICPSR (2018+) ──
print("\n  [6/8] ICPSR deposits & citations (2018+)...")
icpsr_deposits = api("works", {
    "filter": f"primary_location.source.id:S4363604084,{YEAR_FILTER}",
    "group_by": "authorships.institutions.id", "per_page": 200,
})
dep_map = {g["key_display_name"]: g["count"] for g in icpsr_deposits.get("group_by",[])}

icpsr_mentions = api("works", {
    "search": "ICPSR",
    "filter": YEAR_FILTER,
    "group_by": "authorships.institutions.id", "per_page": 200,
})
ment_map = {g["key_display_name"]: g["count"] for g in icpsr_mentions.get("group_by",[])}

# Fuzzy match helper
def fuzzy_get(m, name):
    v = m.get(name, 0)
    if v == 0:
        for k, val in m.items():
            if name.split(",")[0] in k or name.split("-")[0].strip() in k:
                v = max(v, val)
    return v

for p in us_sorted:
    p["icpsr_deposits"] = fuzzy_get(dep_map, p["name"])
    p["icpsr_cites"] = fuzzy_get(ment_map, p["name"])

# ICPSR top cited datasets 2018+
print("    -> Top cited ICPSR datasets 2018+...")
icpsr_top = api("works", {
    "filter": f"primary_location.source.id:S4363604084,{YEAR_FILTER}",
    "sort": "cited_by_count:desc", "per_page": 15,
})
icpsr_top_works = icpsr_top.get("results",[])

# ── 6. Globus MDF (2018+) — already scoped to 2018+ ──
print("\n  [7/8] Globus MDF datasets with institutional affiliations...")
mdf_r = requests.post(f"{GLOBUS}/v1/index/{MDF_INDEX}/search", json={
    "q": "*",
    "filters": [{"type": "match_any", "field_name": "mdf.resource_type", "values": ["dataset"]}],
    "limit": 50
}, timeout=30)
mdf_data = mdf_r.json()
mdf_datasets = []
for entry in mdf_data.get("gmeta",[]):
    content = entry.get("entries",[{}])[0].get("content",{})
    dc = content.get("dc",{}); mdf_meta = content.get("mdf",{})
    titles = dc.get("titles",[])
    title = titles[0].get("title","N/A") if titles else "N/A"
    dois = [r.get("relatedIdentifier","") for r in dc.get("relatedIdentifiers",[]) if "10." in r.get("relatedIdentifier","")]
    affs = set()
    for cr in dc.get("creators",[]):
        for a in cr.get("affiliations",[]): affs.add(a)
    mdf_datasets.append({"title": title[:80], "dois": dois[:2], "affiliations": list(affs)[:5],
                          "ingest": mdf_meta.get("ingest_date","")[:10], "source": mdf_meta.get("source_name","")})

# ==================================================================
# BUILD CHARTS
# ==================================================================
print("\n  [8/8] Generating charts & spreadsheet...\n")

# ── Chart 1: Intl Orgs Comparison (2018+) ──
names = [shorten(p["name"],18) for p in intl]
x = np.arange(len(names)); w = 0.6
fig, axes = plt.subplots(1, 3, figsize=(20,7))
fig.suptitle("International Organizations — 2018–2026 Publishing Metrics", y=1.02)
for ax, key, label, fmt_fn in [
    (axes[0], "works", "Publication Output (2018+)", lambda v,_: f'{v/1000:.0f}K'),
    (axes[1], "impact", "Citations per Work (2018+)", None),
    (axes[2], "h_index", "h-index (all-time)", lambda v,_: f'{v:,.0f}'),
]:
    vals = [p[key] for p in intl]
    bars = ax.barh(x, vals, color=C10[:len(intl)], height=w)
    ax.set_yticks(x); ax.set_yticklabels(names); ax.set_title(label); ax.invert_yaxis()
    if fmt_fn: ax.xaxis.set_major_formatter(ticker.FuncFormatter(fmt_fn))
    for bar, val in zip(bars, vals):
        txt = f'{val:,.1f}' if isinstance(val, float) else f'{val:,}'
        ax.text(bar.get_width()+max(vals)*0.01, bar.get_y()+bar.get_height()/2, txt, va='center', fontsize=8)
fig.tight_layout(); save_fig(fig, "01_intl_orgs_2018.png")

# ── Chart 2: Intl Org Trends (2018–2025) ──
fig, (ax1,ax2) = plt.subplots(1,2,figsize=(16,7))
fig.suptitle("International Organizations — Trends (2018–2025)", y=1.02)
for i, p in enumerate(intl):
    yd = {e["year"]: e for e in p["cby"]}
    ax1.plot(YEARS, [yd.get(y,{}).get("works_count",0) for y in YEARS], marker='o', ms=4, lw=2, color=C10[i%10], label=shorten(p["name"],14))
    ax2.plot(YEARS, [yd.get(y,{}).get("cited_by_count",0) for y in YEARS], marker='s', ms=4, lw=2, color=C10[i%10], label=shorten(p["name"],14))
ax1.set_title("Annual Publications"); ax1.set_xlabel("Year"); ax1.legend(fontsize=7)
ax2.set_title("Annual Citations"); ax2.set_xlabel("Year"); ax2.legend(fontsize=7)
ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v,_: f'{v/1000:.0f}K'))
fig.tight_layout(); save_fig(fig, "02_intl_trends_2018.png")

# ── Chart 3: Top 50 SS Works (2018+) ──
fig, ax = plt.subplots(figsize=(12,18))
fig.suptitle("Top 50 US Social Science Institutions — Works 2018–2026", y=0.995)
colors = plt.cm.viridis(np.linspace(0.2,0.9,50))
vals = [p["ss_works"] for p in us_sorted]
bars = ax.barh(range(50), vals, color=colors, height=0.7)
ax.set_yticks(range(50)); ax.set_yticklabels([f"{i+1}. {shorten(p['name'],26)}" for i,p in enumerate(us_sorted)], fontsize=8)
ax.set_xlabel("Social Science Works (2018+)"); ax.invert_yaxis()
for bar, val in zip(bars, vals):
    ax.text(bar.get_width()+max(vals)*0.005, bar.get_y()+bar.get_height()/2, f'{val:,}', va='center', fontsize=7)
fig.tight_layout(); save_fig(fig, "03_top50_ss_2018.png")

# ── Chart 4: Top 50 Citations/Work (2018+) ──
sp = sorted(us_sorted, key=lambda p: p["impact"], reverse=True)
fig, ax = plt.subplots(figsize=(12,18))
fig.suptitle("Top 50 US Social Science — Citations per Work (2018–2026)", y=0.995)
colors = plt.cm.plasma(np.linspace(0.15,0.85,50))
vals = [p["impact"] for p in sp]
bars = ax.barh(range(50), vals, color=colors, height=0.7)
ax.set_yticks(range(50)); ax.set_yticklabels([f"{i+1}. {shorten(p['name'],26)}" for i,p in enumerate(sp)], fontsize=8)
ax.set_xlabel("Citations per Work (2018+)"); ax.invert_yaxis()
for bar, val in zip(bars, vals):
    ax.text(bar.get_width()+max(vals)*0.005, bar.get_y()+bar.get_height()/2, f'{val:.1f}', va='center', fontsize=7)
fig.tight_layout(); save_fig(fig, "04_top50_impact_2018.png")

# ── Chart 5: Top 20 Trends (2018–2025) ──
fig, (ax1,ax2) = plt.subplots(1,2,figsize=(18,8))
fig.suptitle("Top 20 US Social Science — Trends (2018–2025)", y=1.02)
for i, p in enumerate(top20):
    yd = {e["year"]: e for e in p["cby"]}
    lbl = shorten(p["name"],16)
    ax1.plot(YEARS, [yd.get(y,{}).get("works_count",0) for y in YEARS], marker='o', ms=3, lw=1.5, color=C20[i], label=lbl, alpha=0.85)
    ax2.plot(YEARS, [yd.get(y,{}).get("cited_by_count",0) for y in YEARS], marker='s', ms=3, lw=1.5, color=C20[i], label=lbl, alpha=0.85)
ax1.set_title("Annual Publications"); ax1.set_xlabel("Year"); ax1.set_ylabel("Works")
ax1.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v,_: f'{v/1000:.0f}K')); ax1.legend(fontsize=6, ncol=2)
ax2.set_title("Annual Citations"); ax2.set_xlabel("Year"); ax2.set_ylabel("Citations")
ax2.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v,_: f'{v/1000:.0f}K')); ax2.legend(fontsize=6, ncol=2)
fig.tight_layout(); save_fig(fig, "05_top20_trends_2018.png")

# ── Chart 6: OA Breakdown (2018+) ──
oa_cats = ["gold","green","hybrid","bronze","diamond","closed"]
oa_colors = {"gold":"#f59e0b","green":"#22c55e","hybrid":"#8b5cf6","bronze":"#d97706","diamond":"#06b6d4","closed":"#9ca3af"}
fig, ax = plt.subplots(figsize=(14,10))
fig.suptitle("Top 20 US Social Science — Open Access 2018–2026", y=1.02)
y = np.arange(len(top20)); left = np.zeros(len(top20))
for cat in oa_cats:
    vals = []
    for p in top20:
        oa = oa_map.get(p["id"],{}); total = sum(oa.values()) or 1
        vals.append(oa.get(cat,0)/total*100)
    ax.barh(y, vals, left=left, height=0.6, color=oa_colors[cat], label=cat.title())
    left += np.array(vals)
ax.set_yticks(y); ax.set_yticklabels([shorten(p["name"],24) for p in top20], fontsize=9)
ax.set_xlabel("% of Publications"); ax.set_xlim(0,100)
ax.legend(loc='lower right', ncol=3); ax.invert_yaxis()
fig.tight_layout(); save_fig(fig, "06_oa_breakdown_2018.png")

# ── Chart 7: ICPSR Activity (2018+) ──
icpsr_sorted = sorted([p for p in us_sorted if p["icpsr_deposits"]>0 or p["icpsr_cites"]>0],
                       key=lambda p: p["icpsr_deposits"]+p["icpsr_cites"], reverse=True)[:25]
fig, ax = plt.subplots(figsize=(14,10))
fig.suptitle("ICPSR Activity by Institution — 2018–2026", y=1.02)
y = np.arange(len(icpsr_sorted)); w = 0.35
bars1 = ax.barh(y-w/2, [p["icpsr_deposits"] for p in icpsr_sorted], height=w, color='#2563eb', label='ICPSR Deposits')
bars2 = ax.barh(y+w/2, [p["icpsr_cites"] for p in icpsr_sorted], height=w, color='#dc2626', label='Works Citing ICPSR')
ax.set_yticks(y); ax.set_yticklabels([shorten(p["name"],26) for p in icpsr_sorted], fontsize=8)
ax.set_xlabel("Count"); ax.legend(); ax.invert_yaxis()
fig.tight_layout(); save_fig(fig, "07_icpsr_activity_2018.png")

# ── Chart 8: Topic Heatmap (2018+) ──
all_topics = {}
for p in top20:
    for tname, tc in topic_map.get(p["id"],[]):
        st = tname[:35] if len(tname)>35 else tname
        all_topics[st] = all_topics.get(st,0) + tc
top_t = sorted(all_topics, key=lambda t: all_topics[t], reverse=True)[:15]
matrix = np.zeros((len(top20), len(top_t)))
for i, p in enumerate(top20):
    td = {(t[:35] if len(t)>35 else t): c for t,c in topic_map.get(p["id"],[])}
    for j, t in enumerate(top_t): matrix[i,j] = td.get(t,0)
rs = matrix.sum(axis=1, keepdims=True); rs[rs==0]=1; matrix_pct = matrix/rs*100
fig, ax = plt.subplots(figsize=(18,12))
fig.suptitle("Top 20 US Social Science — Research Topics 2018–2026", y=1.01)
im = ax.imshow(matrix_pct, cmap='YlOrRd', aspect='auto')
ax.set_xticks(range(len(top_t))); ax.set_xticklabels(top_t, rotation=45, ha='right', fontsize=7)
ax.set_yticks(range(len(top20))); ax.set_yticklabels([shorten(p["name"],22) for p in top20], fontsize=8)
for i in range(len(top20)):
    for j in range(len(top_t)):
        v = matrix_pct[i,j]
        if v>0: ax.text(j,i,f'{v:.0f}',ha='center',va='center',fontsize=6,color='white' if v>15 else 'black')
plt.colorbar(im, ax=ax, shrink=0.6, label='% of Top Topics')
fig.tight_layout(); save_fig(fig, "08_topic_heatmap_2018.png")

# ── Chart 9: Combined Scatter (2018+) ──
fig, ax = plt.subplots(figsize=(14,9))
fig.suptitle("Intl Orgs vs Top 10 US Institutions — 2018–2026 Output", y=0.98)
for i, p in enumerate(top20[:10]):
    sz = max(p["h_index"]/3,80)
    ax.scatter(p["works"], p["citations"], s=sz, c=C10[i], alpha=0.7, edgecolors='white', lw=1.5, zorder=3)
    ax.annotate(shorten(p["name"],16), (p["works"],p["citations"]), textcoords="offset points", xytext=(8,6), fontsize=7, fontweight='bold', color=C10[i])
mkrs = ['D','^','v','s','P','X','h','*']
for i, p in enumerate(intl):
    ax.scatter(p["works"], p["citations"], s=150, c='black', alpha=0.8, edgecolors='red', lw=2, zorder=4, marker=mkrs[i%8])
    ax.annotate(shorten(p["name"],14), (p["works"],p["citations"]), textcoords="offset points", xytext=(8,-10), fontsize=7, fontweight='bold', color='darkred')
ax.set_xlabel("Works (2018+)"); ax.set_ylabel("Citations (2018+)")
ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda v,_: f'{v/1000:.0f}K'))
ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda v,_: f'{v/1e6:.1f}M'))
ax.set_title("Circles = US universities | Shapes = Intl orgs", fontsize=9, fontstyle='italic', pad=5)
fig.tight_layout(); save_fig(fig, "09_combined_scatter_2018.png")

# ==================================================================
# BUILD SPREADSHEET
# ==================================================================
print("\n    Building spreadsheet...")
wb = Workbook()

# Sheet 1: Intl Orgs
ws1 = wb.active; ws1.title = "Intl Orgs (2018+)"
h1 = ["Organization","Country","Type","Works 2018+","Citations 2018+","Cit/Work","h-index","OpenAlex Link"]
for c,h in enumerate(h1,1): ws1.cell(row=1,column=c,value=h)
style_header(ws1,1,len(h1))
for r,p in enumerate(intl,2):
    ws1.cell(row=r,column=1,value=p["name"]); sh(ws1,r,1,bold=True)
    ws1.cell(row=r,column=2,value=p["country"]); sh(ws1,r,2)
    ws1.cell(row=r,column=3,value=p["type"]); sh(ws1,r,3)
    ws1.cell(row=r,column=4,value=p["works"]); sh(ws1,r,4,fmt=NUM)
    ws1.cell(row=r,column=5,value=p["citations"]); sh(ws1,r,5,fmt=NUM)
    ws1.cell(row=r,column=6,value=f'=E{r}/D{r}'); sh(ws1,r,6,fmt=DEC)
    ws1.cell(row=r,column=7,value=p["h_index"]); sh(ws1,r,7,fmt=NUM)
    link = f"https://api.openalex.org/institutions/{p['id']}?api_key={API_KEY}"
    cell = ws1.cell(row=r,column=8,value=link); cell.hyperlink = link; sh(ws1,r,8,link=True)
ws1.freeze_panes = 'A2'; auto_width(ws1)

# Sheet 2: Top 50 US
ws2 = wb.create_sheet("Top 50 US SS (2018+)")
h2 = ["Rank","Institution","SS Works 2018+","Total Works 2018+","Citations 2018+","Cit/Work",
      "h-index","ICPSR Deposits","ICPSR Citations","OpenAlex Link"]
for c,h in enumerate(h2,1): ws2.cell(row=1,column=c,value=h)
style_header(ws2,1,len(h2))
for r,p in enumerate(us_sorted,2):
    ws2.cell(row=r,column=1,value=r-1); sh(ws2,r,1,fmt='0',bold=True)
    ws2.cell(row=r,column=2,value=p["name"]); sh(ws2,r,2,bold=True)
    ws2.cell(row=r,column=3,value=p["ss_works"]); sh(ws2,r,3,fmt=NUM)
    ws2.cell(row=r,column=4,value=p["works"]); sh(ws2,r,4,fmt=NUM)
    ws2.cell(row=r,column=5,value=p["citations"]); sh(ws2,r,5,fmt=NUM)
    ws2.cell(row=r,column=6,value=f'=E{r}/D{r}'); sh(ws2,r,6,fmt=DEC)
    ws2.cell(row=r,column=7,value=p["h_index"]); sh(ws2,r,7,fmt=NUM)
    ws2.cell(row=r,column=8,value=p["icpsr_deposits"]); sh(ws2,r,8,fmt=NUM)
    ws2.cell(row=r,column=9,value=p["icpsr_cites"]); sh(ws2,r,9,fmt=NUM)
    link = f"https://api.openalex.org/institutions/{p['id']}?api_key={API_KEY}"
    cell = ws2.cell(row=r,column=10,value=link); cell.hyperlink = link; sh(ws2,r,10,link=True)
ws2.freeze_panes = 'C2'; auto_width(ws2)

# Sheet 3: OA Breakdown
ws3 = wb.create_sheet("OA Breakdown (2018+)")
h3 = ["Institution","Gold","Green","Hybrid","Bronze","Diamond","Closed","Total","OA Rate"]
for c,h in enumerate(h3,1): ws3.cell(row=1,column=c,value=h)
style_header(ws3,1,len(h3))
oa_type_colors = {2:'F59E0B',3:'22C55E',4:'8B5CF6',5:'D97706',6:'06B6D4',7:'9CA3AF'}
for col,color in oa_type_colors.items():
    ws3.cell(row=1,column=col).fill = PatternFill('solid',fgColor=color)
for r,p in enumerate(top20,2):
    oa = oa_map.get(p["id"],{})
    ws3.cell(row=r,column=1,value=p["name"]); sh(ws3,r,1,bold=True)
    for ci,cat in enumerate(["gold","green","hybrid","bronze","diamond","closed"],2):
        ws3.cell(row=r,column=ci,value=oa.get(cat,0)); sh(ws3,r,ci,fmt=NUM)
    ws3.cell(row=r,column=8,value=f'=SUM(B{r}:G{r})'); sh(ws3,r,8,fmt=NUM)
    ws3.cell(row=r,column=9,value=f'=(H{r}-G{r})/H{r}'); sh(ws3,r,9,fmt=PCT)
ws3.freeze_panes = 'B2'; auto_width(ws3)

# Sheet 4: ICPSR Top Datasets (2018+)
ws4 = wb.create_sheet("ICPSR Top Datasets (2018+)")
h4 = ["Rank","Title","Year","Citations","DOI","Institutions","OpenAlex Link"]
for c,h in enumerate(h4,1): ws4.cell(row=1,column=c,value=h)
style_header(ws4,1,len(h4))
for r,w in enumerate(icpsr_top_works[:15],2):
    ws4.cell(row=r,column=1,value=r-1); sh(ws4,r,1,fmt='0',bold=True)
    ws4.cell(row=r,column=2,value=(w.get("title","") or "")[:80]); sh(ws4,r,2,bold=True)
    ws4.cell(row=r,column=3,value=w.get("publication_year","")); sh(ws4,r,3)
    ws4.cell(row=r,column=4,value=w.get("cited_by_count",0)); sh(ws4,r,4,fmt=NUM)
    ws4.cell(row=r,column=5,value=(w.get("doi","") or "").replace("https://doi.org/","")); sh(ws4,r,5)
    insts = set()
    for a in w.get("authorships",[]):
        for inst in a.get("institutions",[]): insts.add(inst.get("display_name",""))
    ws4.cell(row=r,column=6,value=", ".join(list(insts)[:3])); sh(ws4,r,6)
    link = w.get("id",""); cell = ws4.cell(row=r,column=7,value=link); cell.hyperlink = link; sh(ws4,r,7,link=True)
ws4.freeze_panes = 'B2'; auto_width(ws4, mx=50)

# Sheet 5: Research Topics
ws5 = wb.create_sheet("Research Topics (2018+)")
h5 = ["Institution","Topic 1","Count","Topic 2","Count","Topic 3","Count","Topic 4","Count","Topic 5","Count"]
for c,h in enumerate(h5,1): ws5.cell(row=1,column=c,value=h)
style_header(ws5,1,len(h5))
for r,p in enumerate(top20,2):
    ws5.cell(row=r,column=1,value=p["name"]); sh(ws5,r,1,bold=True)
    for i,(tn,tc) in enumerate(topic_map.get(p["id"],[])[:5]):
        ws5.cell(row=r,column=2+i*2,value=tn); sh(ws5,r,2+i*2)
        ws5.cell(row=r,column=3+i*2,value=tc); sh(ws5,r,3+i*2,fmt=NUM)
ws5.freeze_panes = 'B2'; auto_width(ws5, mx=42)

# Sheet 6: Globus MDF Datasets
ws6 = wb.create_sheet("Globus MDF Datasets")
h6 = ["Title","Source","Ingest Date","DOI","Affiliations"]
for c,h in enumerate(h6,1): ws6.cell(row=1,column=c,value=h)
style_header(ws6,1,len(h6))
for r,d in enumerate(mdf_datasets[:50],2):
    ws6.cell(row=r,column=1,value=d["title"]); sh(ws6,r,1,bold=True)
    ws6.cell(row=r,column=2,value=d["source"]); sh(ws6,r,2)
    ws6.cell(row=r,column=3,value=d["ingest"]); sh(ws6,r,3)
    ws6.cell(row=r,column=4,value=d["dois"][0] if d["dois"] else ""); sh(ws6,r,4)
    ws6.cell(row=r,column=5,value=", ".join(d["affiliations"])); sh(ws6,r,5)
ws6.freeze_panes = 'B2'; auto_width(ws6, mx=50)

wb.save(XLSX_PATH)
print(f"    Saved: {XLSX_PATH}")

print(f"\n{'='*65}")
print(f"  All data rebuilt for 2018–2026")
print(f"  Charts: {OUTPUT_DIR}")
print(f"  Spreadsheet: {XLSX_PATH}")
print(f"{'='*65}\n")
