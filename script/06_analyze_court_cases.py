"""
Analysis of court cases concerning seized homes in occupied Ukraine (Donetsk oblast).

Input : court_cases_and_opinions.xlsx  (one row = one court case)
Output: PNG charts in ./charts/ + a per-capita summary CSV

Charts produced
  1. combined_cases_over_time.png          – all courts together, monthly
  2. combined_cases_by_solution.png        – all courts, one line per "solution"
  3. cases_by_court.png                    – small multiples, one panel per court
                                             (solution lines + bold "all cases" line),
                                             ordered by total case count (most first)
  4. cases_per_capita.png                  – cases per 10,000 residents of the
                                             court's jurisdiction (last available
                                             Ukrainian population estimates)
  5. heatmap_court_by_month.png            – calendar heat map (court x month), red scale
  6. heatmap_geographic.png                – map of Donetsk oblast, localities shaded
                                             darker red the more cases they have

All time series use `date_of_receipt` (the date the case was received), as requested.
"""

import json
import os
import urllib.request

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import pandas as pd
from matplotlib.colors import Normalize
from matplotlib.cm import ScalarMappable

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
INPUT_FILE = "../output/court_cases_and_opinions.xlsx"
OUT_DIR = "charts"
DATE_COL = "date_of_receipt"      # per instructions, use date received
COURT_COL = "court_name"
SOLUTION_COL = "solution"

# ----------------------------------------------------------------------------
# Population table
# ----------------------------------------------------------------------------
# The courts in this dataset are (renamed) district / city / raion courts in
# occupied Donetsk oblast. Each entry maps a court to its jurisdiction, an
# APPROXIMATE population figure, and coordinates of the locality.
#
# Population figures are the last available Ukrainian (Derzhstat) estimates —
# mostly the 1 Jan 2022 "present population" series for cities, and the last
# pre-2020-reform raion estimates for raions. City-district figures (Mariupol,
# Donetsk, Makiivka) are approximate splits of the city totals, since Ukraine
# has not published district-level data for these cities since 2014.
# >>> These numbers are estimates — verify/replace them before publication. <<<
POPULATION = {
    # court_name          (jurisdiction,                          population, lat,    lon,    locality for map)
    "Zhovtnevyi":          ("Mariupol – Zhovtnevyi/Tsentralnyi district", 130_000, 47.097, 37.543, "Mariupol"),
    "Primorskii":          ("Mariupol – Prymorskyi district",             90_000,  47.097, 37.543, "Mariupol"),
    "Ordzhonikidzevskii":  ("Mariupol – Ordzhonikidzevskyi/Livoberezhnyi district", 110_000, 47.097, 37.543, "Mariupol"),
    "Ilichevskii":         ("Mariupol – Illichivskyi/Kalmiuskyi district", 95_000, 47.097, 37.543, "Mariupol"),
    "Voroshilovskii":      ("Donetsk – Voroshylovskyi district",          90_000,  48.003, 37.805, "Donetsk"),
    "Kirovskii":           ("Donetsk – Kirovskyi district",               170_000, 48.003, 37.805, "Donetsk"),
    "Budennovskii":        ("Donetsk – Budonnivskyi district",            100_000, 48.003, 37.805, "Donetsk"),
    "Gornyatskii":         ("Makiivka – Hirnytskyi district",             85_000,  48.048, 37.926, "Makiivka"),
    "Tsentralno-Gorodskoi":("Makiivka – Tsentralno-Miskyi district",      85_000,  48.048, 37.926, "Makiivka"),
    "Khartsyzskii":        ("Khartsyzk (city council area)",              92_000,  48.033, 38.147, "Khartsyzk"),
    "Enakievskii":         ("Yenakiieve (city council area)",             103_000, 48.232, 38.205, "Yenakiieve"),
    "Gorlovskii":          ("Horlivka (city)",                            240_000, 48.334, 38.093, "Horlivka"),
    "Yasinovatskii":       ("Yasynuvata (city + raion)",                  57_000,  48.128, 37.855, "Yasynuvata"),
    "Volnovakhskii":       ("Volnovakha raion",                           80_000,  47.601, 37.494, "Volnovakha"),
    "Amvrosievskii":       ("Amvrosiivka raion",                          44_000,  47.795, 38.478, "Amvrosiivka"),
    "Telmanovskii":        ("Telmanove/Boikivske raion",                  25_000,  47.410, 38.098, "Telmanove"),
    "Starobeshevskii":     ("Starobesheve raion",                         49_000,  47.752, 38.028, "Starobesheve"),
    "Novoazovskii":        ("Novoazovsk raion",                           33_000,  47.111, 38.083, "Novoazovsk"),
    "Volodarskii":         ("Volodarske/Nikolske raion",                  27_000,  47.187, 37.205, "Nikolske"),
    "Debaltsevskii":       ("Debaltseve (city)",                          25_000,  48.339, 38.406, "Debaltseve"),
    "Marinskii":           ("Marinka raion",                              84_000,  47.940, 37.503, "Marinka"),
    "Pershotravnevyi":     ("Pershotravnevyi/Manhush raion",              25_000,  47.056, 37.303, "Manhush"),
}

# Donetsk oblast boundary (single-feature GeoJSON, EugeneBorshch/ukraine_geojson)
GEOJSON_FILE = "donetska.geojson"
GEOJSON_URL = ("https://raw.githubusercontent.com/EugeneBorshch/"
               "ukraine_geojson/master/UA_14_Donetska.geojson")


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def load_data(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df["date"] = pd.to_datetime(df[DATE_COL], format="%d.%m.%Y", errors="coerce")
    n_bad = df["date"].isna().sum()
    if n_bad:
        print(f"Warning: {n_bad} rows had unparseable '{DATE_COL}' and are dropped from time charts.")
    df[SOLUTION_COL] = df[SOLUTION_COL].fillna("(no solution recorded)")
    df["month"] = df["date"].dt.to_period("M").dt.to_timestamp()
    return df


def shorten_solution(s: str) -> str:
    """Compact labels for legends."""
    mapping = {
        "Claim (application, complaint) SATISFIED": "Satisfied",
        "Claim (application, complaint) LEFT WITHOUT CONSIDERATION": "Left without consideration",
        "Proceedings in the case have been TERMINATED": "Proceedings terminated",
        "Application RETURNED to applicant": "Returned to applicant",
        "REFUSED to satisfy the claim (application, complaint)": "Refused",
        "Case joined to another case": "Joined to another case",
        "ACCEPTANCE OF APPLICATION IS DENIED": "Acceptance denied",
        "Claim (application, complaint) PARTIALLY SATISFIED": "Partially satisfied",
        "Transferred according to jurisdiction, jurisdiction": "Transferred (jurisdiction)",
        "(no solution recorded)": "No solution recorded",
    }
    return mapping.get(s, s)


def monthly_index(df: pd.DataFrame) -> pd.DatetimeIndex:
    return pd.date_range(df["month"].min(), df["month"].max(), freq="MS")


def solution_palette(solutions):
    cmap = plt.get_cmap("tab10")
    return {s: cmap(i % 10) for i, s in enumerate(solutions)}


# ----------------------------------------------------------------------------
# Charts
# ----------------------------------------------------------------------------
def chart_combined(df, idx):
    counts = df.groupby("month").size().reindex(idx, fill_value=0)
    fig, ax = plt.subplots(figsize=(11, 5.5))
    ax.plot(counts.index, counts.values, color="#8B0000", lw=2.2, marker="o", ms=4)
    ax.fill_between(counts.index, counts.values, color="#8B0000", alpha=0.12)
    ax.set_title("Seized-home court cases in occupied Donetsk oblast — all courts combined\n"
                 f"(monthly, by date of receipt; n = {len(df):,})", fontsize=12)
    ax.set_ylabel("Cases received per month")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.grid(alpha=0.3)
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(f"{OUT_DIR}/1_combined_cases_over_time.png", dpi=200)
    plt.close(fig)


def chart_combined_by_solution(df, idx, sol_order, palette):
    fig, ax = plt.subplots(figsize=(11.5, 6))
    for sol in sol_order:
        sub = df[df[SOLUTION_COL] == sol]
        counts = sub.groupby("month").size().reindex(idx, fill_value=0)
        ax.plot(counts.index, counts.values, lw=1.8,
                label=f"{shorten_solution(sol)} ({len(sub):,})", color=palette[sol])
    ax.set_title("Cases over time by outcome ('solution') — all courts combined\n"
                 "(monthly, by date of receipt)", fontsize=12)
    ax.set_ylabel("Cases received per month")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.grid(alpha=0.3)
    ax.legend(fontsize=8, title="Solution (total cases)", loc="upper left")
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(f"{OUT_DIR}/2_combined_cases_by_solution.png", dpi=200)
    plt.close(fig)


def chart_per_court(df, idx, sol_order, palette, court_order):
    n = len(court_order)
    ncols = 3
    nrows = int(np.ceil(n / ncols))
    fig, axes = plt.subplots(nrows, ncols, figsize=(15, 3.1 * nrows),
                             sharex=True)
    axes = np.atleast_2d(axes)
    for k, court in enumerate(court_order):
        ax = axes[k // ncols, k % ncols]
        sub = df[df[COURT_COL] == court]
        total = sub.groupby("month").size().reindex(idx, fill_value=0)
        ax.plot(total.index, total.values, color="black", lw=2.2,
                label="All cases", zorder=5)
        for sol in sol_order:
            s2 = sub[sub[SOLUTION_COL] == sol]
            if s2.empty:
                continue
            c = s2.groupby("month").size().reindex(idx, fill_value=0)
            ax.plot(c.index, c.values, lw=1.1, alpha=0.85, color=palette[sol],
                    label=shorten_solution(sol))
        ax.set_title(f"{court} (n = {len(sub):,})", fontsize=10)
        ax.grid(alpha=0.3)
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=6))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %y"))
        ax.tick_params(labelsize=8)
    # hide unused panels
    for k in range(n, nrows * ncols):
        axes[k // ncols, k % ncols].set_visible(False)
    # one shared legend
    handles = [plt.Line2D([], [], color="black", lw=2.2, label="All cases")] + \
              [plt.Line2D([], [], color=palette[s], lw=1.5,
                          label=shorten_solution(s)) for s in sol_order]
    fig.legend(handles=handles, loc="lower center", ncol=4, fontsize=9,
               frameon=False, bbox_to_anchor=(0.5, -0.005))
    fig.suptitle("Cases over time by court (ordered by total case count, most first)\n"
                 "black = all cases; coloured = by solution; monthly, by date of receipt",
                 fontsize=13, y=0.995)
    fig.tight_layout(rect=[0, 0.035, 1, 0.97])
    fig.savefig(f"{OUT_DIR}/3_cases_by_court.png", dpi=200)
    plt.close(fig)


def chart_per_capita(df, court_order):
    rows = []
    for court in court_order:
        n = (df[COURT_COL] == court).sum()
        if court in POPULATION:
            juris, pop, lat, lon, loc = POPULATION[court]
            rows.append((court, juris, n, pop, 10_000 * n / pop))
        else:
            rows.append((court, "UNKNOWN — add to POPULATION dict", n, np.nan, np.nan))
    per_cap = pd.DataFrame(rows, columns=["court", "jurisdiction", "cases",
                                          "population_est", "cases_per_10k"])
    per_cap.to_csv(f"{OUT_DIR}/cases_per_capita.csv", index=False)

    pc = per_cap.dropna(subset=["cases_per_10k"]).sort_values("cases_per_10k")
    fig, ax = plt.subplots(figsize=(10, 0.42 * len(pc) + 2))
    norm = Normalize(vmin=0, vmax=pc["cases_per_10k"].max())
    colors = plt.get_cmap("Reds")(0.25 + 0.75 * norm(pc["cases_per_10k"].values))
    ax.barh(pc["court"] + "\n(" + pc["jurisdiction"] + ")",
        pc["cases_per_10k"], color=colors, edgecolor="grey", lw=0.4)
    for y, (v, n) in enumerate(zip(pc["cases_per_10k"], pc["cases"])):
        ax.text(v + 0.5, y, f"{v:.1f}  ({n:,} cases)", va="center", fontsize=8)
    ax.set_xlabel("Cases per 10,000 residents of the court's jurisdiction")
    ax.set_title("Case intensity relative to population\n"
                 "(last available Ukrainian population estimates — approximate; "
                 "district-level splits estimated)", fontsize=11)
    ax.set_xlim(0, pc["cases_per_10k"].max() * 1.22)
    ax.grid(axis="x", alpha=0.3)
    ax.tick_params(axis="y", labelsize=8)
    fig.tight_layout()
    fig.savefig(f"{OUT_DIR}/4_cases_per_capita.png", dpi=200)
    plt.close(fig)
    return per_cap


def chart_heatmap_court_month(df, idx, court_order):
    mat = np.zeros((len(court_order), len(idx)))
    for i, court in enumerate(court_order):
        c = df[df[COURT_COL] == court].groupby("month").size().reindex(idx, fill_value=0)
        mat[i] = c.values
    fig, ax = plt.subplots(figsize=(13, 0.42 * len(court_order) + 2.5))
    im = ax.imshow(mat, aspect="auto", cmap="Reds")
    ax.set_yticks(range(len(court_order)))
    ax.set_yticklabels([f"{c}  ({(df[COURT_COL]==c).sum():,})" for c in court_order],
                       fontsize=8)
    step = max(1, len(idx) // 15)
    ax.set_xticks(range(0, len(idx), step))
    ax.set_xticklabels([d.strftime("%b %y") for d in idx[::step]],
                       rotation=45, ha="right", fontsize=8)
    ax.set_title("Heat map: cases received per court per month "
                 "(darker red = more cases)\ncourts ordered by total case count",
                 fontsize=12)
    fig.colorbar(im, ax=ax, label="Cases per month", shrink=0.8)
    fig.tight_layout()
    fig.savefig(f"{OUT_DIR}/5_heatmap_court_by_month.png", dpi=200)
    plt.close(fig)


def _iter_polygons(geom):
    if geom["type"] == "Polygon":
        yield geom["coordinates"]
    elif geom["type"] == "MultiPolygon":
        yield from geom["coordinates"]


def chart_geographic_heatmap(df, per_cap):
    # aggregate cases by locality (Mariupol's 4 district courts -> one point, etc.)
    agg = {}
    for court, (juris, pop, lat, lon, loc) in POPULATION.items():
        n = (df[COURT_COL] == court).sum()
        if loc not in agg:
            agg[loc] = {"lat": lat, "lon": lon, "cases": 0, "pop": 0}
        agg[loc]["cases"] += n
        agg[loc]["pop"] += pop

    fig, ax = plt.subplots(figsize=(11, 9))

    # oblast boundary (optional – downloaded on first run)
    boundary_ok = False
    if not os.path.exists(GEOJSON_FILE):
        try:
            urllib.request.urlretrieve(GEOJSON_URL, GEOJSON_FILE)
        except Exception as e:
            print(f"Could not download oblast boundary ({e}); map drawn without it.")
    if os.path.exists(GEOJSON_FILE):
        try:
            gj = json.load(open(GEOJSON_FILE))
            geom = gj["geometry"] if gj.get("type") == "Feature" else \
                   gj["features"][0]["geometry"]
            for poly in _iter_polygons(geom):
                for ring in poly:
                    xs, ys = zip(*ring)
                    ax.plot(xs, ys, color="grey", lw=1)
                    ax.fill(xs, ys, color="#f5f5f5", zorder=0)
            boundary_ok = True
        except Exception as e:
            print(f"Boundary file unreadable ({e}); map drawn without it.")

    vals = np.array([v["cases"] for v in agg.values()], dtype=float)
    norm = Normalize(vmin=0, vmax=vals.max())
    cmap = plt.get_cmap("Reds")
    for loc, v in sorted(agg.items(), key=lambda kv: -kv[1]["cases"]):
        color = cmap(0.15 + 0.85 * norm(v["cases"]))
        size = 120 + 2600 * norm(v["cases"])
        ax.scatter(v["lon"], v["lat"], s=size, color=color, edgecolor="#5a0000",
                   lw=0.8, zorder=3)
        ax.annotate(f"{loc}\n{v['cases']:,}", (v["lon"], v["lat"]),
                    textcoords="offset points", xytext=(0, 12 + 0.004 * size),
                    ha="center", fontsize=8, zorder=4)

    sm = ScalarMappable(norm=norm, cmap=cmap)
    fig.colorbar(sm, ax=ax, shrink=0.7, label="Total cases")
    ax.set_title("Geographic heat map of seized-home cases, Donetsk oblast\n"
                 "(darker/larger red = more cases; localities = court jurisdictions, "
                 "city district courts aggregated)", fontsize=12)
    ax.set_xlabel("Longitude"); ax.set_ylabel("Latitude")
    ax.set_aspect(1 / np.cos(np.deg2rad(47.8)))  # rough lat correction
    if boundary_ok:
        ax.set_xlim(36.4, 39.2); ax.set_ylim(46.7, 49.4)
    ax.grid(alpha=0.2)
    fig.tight_layout()
    fig.savefig(f"{OUT_DIR}/6_heatmap_geographic.png", dpi=200)
    plt.close(fig)

def chart_avg_processing_time(df, idx):
    d = df.copy()
    d["decision_date_parsed"] = pd.to_datetime(d["decision_date"], format="%d.%m.%Y", errors="coerce")
    d["processing_days"] = (d["decision_date_parsed"] - d["date"]).dt.days

    # drop rows with missing/negative durations (bad parses or data errors)
    d = d[d["processing_days"].notna() & (d["processing_days"] >= 0)]

    avg_by_month = d.groupby("month")["processing_days"].mean().reindex(idx)

    fig, ax = plt.subplots(figsize=(11, 5.5))
    ax.plot(avg_by_month.index, avg_by_month.values, color="#2c5f8a", lw=2.2, marker="o", ms=4)
    ax.set_title("Average time from receipt to decision, over time\n"
                 f"(monthly average, by date of receipt; n = {len(d):,} cases with a decision date)",
                 fontsize=12)
    ax.set_ylabel("Average days to decision")
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.grid(alpha=0.3)
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(f"{OUT_DIR}/7_avg_processing_time.png", dpi=200)
    plt.close(fig)

def chart_geographic_heatmap_per_capita(df, per_cap):
    # aggregate cases AND population by locality (Mariupol's 4 district courts -> one point, etc.)
    agg = {}
    for court, (juris, pop, lat, lon, loc) in POPULATION.items():
        n = (df[COURT_COL] == court).sum()
        if loc not in agg:
            agg[loc] = {"lat": lat, "lon": lon, "cases": 0, "pop": 0}
        agg[loc]["cases"] += n
        agg[loc]["pop"] += pop

    # compute per-10k rate for each locality (guard against missing/zero population)
    for loc, v in agg.items():
        v["per_10k"] = 10_000 * v["cases"] / v["pop"] if v["pop"] else np.nan

    fig, ax = plt.subplots(figsize=(11, 9))

    # oblast boundary (optional – downloaded on first run)
    boundary_ok = False
    if not os.path.exists(GEOJSON_FILE):
        try:
            urllib.request.urlretrieve(GEOJSON_URL, GEOJSON_FILE)
        except Exception as e:
            print(f"Could not download oblast boundary ({e}); map drawn without it.")
    if os.path.exists(GEOJSON_FILE):
        try:
            gj = json.load(open(GEOJSON_FILE))
            geom = gj["geometry"] if gj.get("type") == "Feature" else \
                   gj["features"][0]["geometry"]
            for poly in _iter_polygons(geom):
                for ring in poly:
                    xs, ys = zip(*ring)
                    ax.plot(xs, ys, color="grey", lw=1)
                    ax.fill(xs, ys, color="#f5f5f5", zorder=0)
            boundary_ok = True
        except Exception as e:
            print(f"Boundary file unreadable ({e}); map drawn without it.")

    vals = np.array([v["per_10k"] for v in agg.values() if not np.isnan(v["per_10k"])])
    norm = Normalize(vmin=0, vmax=vals.max())
    cmap = plt.get_cmap("Reds")
    for loc, v in sorted(agg.items(), key=lambda kv: -kv[1]["per_10k"]):
        if np.isnan(v["per_10k"]):
            continue
        color = cmap(0.15 + 0.85 * norm(v["per_10k"]))
        size = 120 + 2600 * norm(v["per_10k"])
        ax.scatter(v["lon"], v["lat"], s=size, color=color, edgecolor="#5a0000",
                   lw=0.8, zorder=3)
        ax.annotate(f"{loc}\n{v['per_10k']:.0f} / 10k\n({v['cases']:,} cases)",
                    (v["lon"], v["lat"]),
                    textcoords="offset points", xytext=(0, 14 + 0.004 * size),
                    ha="center", fontsize=8, zorder=4)

    sm = ScalarMappable(norm=norm, cmap=cmap)
    fig.colorbar(sm, ax=ax, shrink=0.7, label="Cases per 10,000 residents")
    ax.set_title("Geographic heat map of case intensity relative to population, Donetsk oblast\n"
                 "(darker/larger red = more cases per capita; localities = court jurisdictions, "
                 "city district courts aggregated;\npopulation figures approximate — see POPULATION dict)",
                 fontsize=11)
    ax.set_xlabel("Longitude"); ax.set_ylabel("Latitude")
    ax.set_aspect(1 / np.cos(np.deg2rad(47.8)))  # rough lat correction
    if boundary_ok:
        ax.set_xlim(36.4, 39.2); ax.set_ylim(46.7, 49.4)
    ax.grid(alpha=0.2)
    fig.tight_layout()
    fig.savefig(f"{OUT_DIR}/8_heatmap_geographic_per_capita.png", dpi=200)
    plt.close(fig)
# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    df = load_data(INPUT_FILE)
    df = df.dropna(subset=["date"])
    idx = monthly_index(df)

    sol_order = df[SOLUTION_COL].value_counts().index.tolist()
    palette = solution_palette(sol_order)
    court_order = df[COURT_COL].value_counts().index.tolist()  # most cases first

    chart_combined(df, idx)
    chart_combined_by_solution(df, idx, sol_order, palette)
    chart_per_court(df, idx, sol_order, palette, court_order)
    per_cap = chart_per_capita(df, court_order)
    chart_heatmap_court_month(df, idx, court_order)
    chart_geographic_heatmap(df, per_cap)
    chart_avg_processing_time(df, idx)
    chart_geographic_heatmap_per_capita(df, per_cap)

    print(f"Done. {len(df):,} cases, {len(court_order)} courts, "
          f"{df['month'].min():%b %Y} – {df['month'].max():%b %Y}.")
    print(f"Charts written to ./{OUT_DIR}/")


if __name__ == "__main__":
    main()
