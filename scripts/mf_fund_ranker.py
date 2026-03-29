"""
MF Fund Ranker — Decision Engine
=================================
Reads your dashboard_data.xlsx, scores every scheme within its category,
and outputs a ranked Excel file: mf_ranked_screener.xlsx

Scoring Formula (Best Absolute Returns):
  - 1Y Return  : 40%
  - 3Y CAGR    : 40%
  - AUM        : 10%  (larger = more trust)
  - Expense Ratio: 10% (lower = better, inverted)

Usage:
  pip install pandas openpyxl
  python mf_fund_ranker.py

Place dashboard_data.xlsx in the same folder as this script.
Output: mf_ranked_screener.xlsx
"""

import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

# ── CONFIG ──────────────────────────────────────────────────────────────────
INPUT_FILE  = "dashboard_data.xlsx"
OUTPUT_FILE = "mf_ranked_screener.xlsx"

WEIGHTS = {
    "return_1y":  0.40,
    "return_3y":  0.40,
    "aum":        0.10,
    "expense":    0.10,   # inverted: lower expense → higher score
}

# ── COLUMN AUTO-DETECTION ───────────────────────────────────────────────────
COLUMN_ALIASES = {
    "scheme_name": [
        "scheme_name", "scheme name", "fund name", "name", "schemename",
        "fund", "scheme"
    ],
    "category": [
        "category", "scheme_category", "scheme category", "cat",
        "fund_category", "fund category", "sub_category", "sub category",
        "category_name", "cat_level_3", "cat_level_2", "sub_type"
    ],
    "amc": [
        "amc", "amc_name", "amc name", "fund_house", "fund house",
        "amcname", "asset_management_company", "house"
    ],
    "return_1y": [
        "return_1y", "1y_return", "1yr_return", "returns_1y", "1y return",
        "1 year return", "1yr", "ret_1y", "cagr_1y", "1y_cagr",
        "trailing_1y", "1year_return", "ann_return_1y"
    ],
    "return_3y": [
        "return_3y", "3y_return", "3yr_return", "returns_3y", "3y return",
        "3 year return", "3yr", "ret_3y", "cagr_3y", "3y_cagr",
        "trailing_3y", "3year_return", "ann_return_3y"
    ],
    "aum": [
        "aum", "aum_cr", "aum_crore", "aum (cr)", "net_assets",
        "net assets", "corpus", "fund_size", "fund size", "aum_in_cr"
    ],
    "expense": [
        "expense_ratio", "expense ratio", "ter", "total_expense_ratio",
        "expense", "exp_ratio", "expenser_ratio"
    ],
    "nav": [
        "nav", "net_asset_value", "current_nav", "latest_nav"
    ],
    "plan": [
        "plan", "plan_type", "scheme_plan", "direct_regular"
    ],
}

COLORS = {
    "header_bg":   "1A3A5C",   # dark navy
    "header_fg":   "FFFFFF",
    "rank1_bg":    "FFD700",   # gold
    "rank2_bg":    "E8E8E8",   # silver
    "rank3_bg":    "D4956A",   # bronze
    "positive":    "1E7A4B",   # dark green
    "negative":    "C0392B",   # red
    "cat_header":  "2D6A8A",   # teal
    "cat_fg":      "FFFFFF",
    "alt_row":     "F5F0E8",   # cream
    "score_bg":    "EBF5FB",
    "border":      "CCCCCC",
}


def detect_column(df_cols, key):
    """Find the actual column name from aliases."""
    cols_lower = {c.lower().strip(): c for c in df_cols}
    for alias in COLUMN_ALIASES.get(key, []):
        if alias.lower() in cols_lower:
            return cols_lower[alias.lower()]
    return None


def load_data(filepath):
    """Load all sheets, combine, and return a unified DataFrame."""
    print(f"\n📂 Reading: {filepath}")
    sheets = pd.read_excel(filepath, sheet_name=None)
    print(f"   Sheets found: {list(sheets.keys())}")

    frames = []
    for name, df in sheets.items():
        df["_source_sheet"] = name
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    print(f"   Total rows across all sheets: {len(combined)}")
    print(f"   Columns: {list(combined.columns)}")
    return combined


def map_columns(df):
    """Detect and map standard field names."""
    mapping = {}
    for key in COLUMN_ALIASES:
        col = detect_column(df.columns, key)
        mapping[key] = col
        status = f"✅ '{col}'" if col else "❌ not found"
        print(f"   {key:<15} → {status}")
    return mapping


def to_numeric(series):
    """Clean and convert a series to numeric, coercing errors."""
    if series is None:
        return pd.Series(dtype=float)
    s = series.astype(str).str.replace('%', '').str.replace(',', '').str.strip()
    return pd.to_numeric(s, errors='coerce')


def percentile_score(series):
    """Rank-based percentile score 0–100 within the group."""
    ranks = series.rank(method='min', na_option='bottom')
    return (ranks - 1) / max(len(series) - 1, 1) * 100


def score_funds(df, mapping):
    """Compute composite score for each fund row."""
    df = df.copy()

    # Pull numeric columns
    r1  = to_numeric(df[mapping["return_1y"]] if mapping["return_1y"] else None)
    r3  = to_numeric(df[mapping["return_3y"]] if mapping["return_3y"] else None)
    aum = to_numeric(df[mapping["aum"]]       if mapping["aum"]       else None)
    exp = to_numeric(df[mapping["expense"]]   if mapping["expense"]   else None)

    df["_r1"]  = r1
    df["_r3"]  = r3
    df["_aum"] = aum
    df["_exp"] = exp

    cat_col = mapping["category"]
    if not cat_col:
        df["_category_clean"] = "All Funds"
    else:
        df["_category_clean"] = df[cat_col].astype(str).str.strip().str.title()

    # Score per category
    score_parts = []
    for cat, grp in df.groupby("_category_clean"):
        g = grp.copy()
        s1  = percentile_score(g["_r1"].fillna(g["_r1"].median()))  * WEIGHTS["return_1y"]
        s3  = percentile_score(g["_r3"].fillna(g["_r3"].median()))  * WEIGHTS["return_3y"]
        sa  = percentile_score(g["_aum"].fillna(g["_aum"].median()))* WEIGHTS["aum"]
        # Expense: lower is better → invert
        se  = percentile_score(-g["_exp"].fillna(g["_exp"].median()))* WEIGHTS["expense"]
        g["_composite_score"] = (s1 + s3 + sa + se).round(1)
        g["_rank"] = g["_composite_score"].rank(method='min', ascending=False).astype(int)
        score_parts.append(g)

    return pd.concat(score_parts, ignore_index=True)


def get_col_val(row, col):
    return row[col] if col and col in row.index else None


def build_excel(df, mapping, output_path):
    """Write the ranked screener to a well-formatted Excel workbook."""
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    thin = Side(style='thin', color=COLORS["border"])
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr_font(bold=True, color="FFFFFF", size=10):
        return Font(name="Arial", bold=bold, color=color, size=size)

    def cell_font(bold=False, color="000000", size=9):
        return Font(name="Arial", bold=bold, color=color, size=size)

    def fill(hex_color):
        return PatternFill("solid", start_color=hex_color, fgColor=hex_color)

    def pct(val):
        if pd.isna(val): return "—"
        return f"{val:.1f}%"

    def crore(val):
        if pd.isna(val): return "—"
        if val >= 100000: return f"₹{val/100000:.1f}L Cr"
        if val >= 1000:   return f"₹{val/1000:.1f}K Cr"
        return f"₹{val:.0f} Cr"

    def score_color(score):
        if score >= 80: return "1E7A4B"
        if score >= 60: return "2980B9"
        if score >= 40: return "F39C12"
        return "C0392B"

    categories = sorted(df["_category_clean"].unique())
    print(f"\n📊 Writing {len(categories)} category sheets...")

    # ── Per-category sheets ──────────────────────────────────────────────
    for cat in categories:
        cat_df = df[df["_category_clean"] == cat].sort_values("_rank")
        safe_name = cat[:31].replace("/", "-").replace("\\", "-").replace("*", "").replace("?", "").replace("[", "").replace("]", "")
        ws = wb.create_sheet(title=safe_name)

        # Title row
        ws.merge_cells("A1:K1")
        ws["A1"] = f"📈  {cat}  —  Fund Performance Ranker"
        ws["A1"].font = Font(name="Arial", bold=True, size=13, color=COLORS["cat_fg"])
        ws["A1"].fill = fill(COLORS["cat_header"])
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 28

        # Sub-header
        ws.merge_cells("A2:K2")
        ws["A2"] = f"Scoring: 1Y Returns 40%  |  3Y CAGR 40%  |  AUM 10%  |  Expense Ratio 10%   |   Total funds: {len(cat_df)}"
        ws["A2"].font = Font(name="Arial", italic=True, size=8, color="555555")
        ws["A2"].fill = fill("F0F4F8")
        ws["A2"].alignment = Alignment(horizontal="left", indent=1)
        ws.row_dimensions[2].height = 16

        # Column headers
        headers = ["Rank", "Scheme Name", "AMC", "1Y Return", "3Y CAGR",
                   "AUM", "Expense Ratio", "Composite Score", "Signal", "Plan", "Category"]
        ws.append(headers)
        for col_idx, hdr in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx)
            cell.font = hdr_font()
            cell.fill = fill(COLORS["header_bg"])
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        ws.row_dimensions[3].height = 18

        # Data rows
        for i, (_, row) in enumerate(cat_df.iterrows(), 4):
            rank   = row.get("_rank", i - 3)
            name   = get_col_val(row, mapping["scheme_name"]) or "—"
            amc    = get_col_val(row, mapping["amc"])         or "—"
            r1     = row.get("_r1")
            r3     = row.get("_r3")
            aum_v  = row.get("_aum")
            exp_v  = row.get("_exp")
            score  = row.get("_composite_score", 0)
            plan   = get_col_val(row, mapping["plan"])        or "—"

            # Signal
            if score >= 75:   signal = "⭐ Strong Buy"
            elif score >= 55: signal = "✅ Buy"
            elif score >= 40: signal = "⚠️ Watch"
            else:             signal = "❌ Avoid"

            values = [rank, name, amc, pct(r1), pct(r3),
                      crore(aum_v), pct(exp_v) if not pd.isna(exp_v) else "—",
                      round(score, 1), signal, str(plan), cat]

            # Row background
            if rank == 1:   row_bg = COLORS["rank1_bg"]
            elif rank == 2: row_bg = COLORS["rank2_bg"]
            elif rank == 3: row_bg = COLORS["rank3_bg"]
            elif i % 2 == 0: row_bg = COLORS["alt_row"]
            else:            row_bg = "FFFFFF"

            for col_idx, val in enumerate(values, 1):
                cell = ws.cell(row=i, column=col_idx, value=val)
                cell.font = cell_font(bold=(rank <= 3), size=9)
                cell.fill = fill(row_bg)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=(col_idx == 2))

                # Colour 1Y/3Y returns
                if col_idx in (4, 5) and isinstance(val, str) and val != "—":
                    num = float(val.replace('%', ''))
                    cell.font = Font(name="Arial", size=9, bold=(rank<=3),
                                     color=COLORS["positive"] if num > 0 else COLORS["negative"])
                # Score colour
                if col_idx == 8:
                    cell.font = Font(name="Arial", bold=True, size=9, color=score_color(score))

            ws.row_dimensions[i].height = 16

        # Column widths
        widths = [6, 52, 22, 12, 12, 14, 14, 16, 14, 12, 22]
        for ci, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

        ws.freeze_panes = "A4"
        print(f"   ✅ {cat} — {len(cat_df)} funds ranked")

    # ── Summary Sheet ────────────────────────────────────────────────────
    ws_sum = wb.create_sheet(title="🏆 SUMMARY", index=0)

    ws_sum.merge_cells("A1:G1")
    ws_sum["A1"] = "MF INTELLIGENCE — TOP FUND PER CATEGORY"
    ws_sum["A1"].font = Font(name="Arial", bold=True, size=14, color="FFFFFF")
    ws_sum["A1"].fill = fill("0D1117")
    ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 30

    ws_sum.merge_cells("A2:G2")
    ws_sum["A2"] = "Scoring: 1Y Returns (40%) + 3Y CAGR (40%) + AUM (10%) + Expense Ratio inverse (10%)"
    ws_sum["A2"].font = Font(name="Arial", italic=True, size=8, color="555555")
    ws_sum["A2"].fill = fill("F5F5F5")
    ws_sum["A2"].alignment = Alignment(horizontal="center")
    ws_sum.row_dimensions[2].height = 14

    sum_headers = ["Category", "#1 Fund", "AMC", "1Y Return", "3Y CAGR", "Score", "Signal"]
    ws_sum.append(sum_headers)
    for ci, hdr in enumerate(sum_headers, 1):
        c = ws_sum.cell(row=3, column=ci)
        c.font = hdr_font()
        c.fill = fill(COLORS["header_bg"])
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    ws_sum.row_dimensions[3].height = 18

    for i, cat in enumerate(categories, 4):
        cat_df = df[df["_category_clean"] == cat].sort_values("_rank")
        if cat_df.empty: continue
        top = cat_df.iloc[0]
        name  = get_col_val(top, mapping["scheme_name"]) or "—"
        amc   = get_col_val(top, mapping["amc"])         or "—"
        r1    = top.get("_r1")
        r3    = top.get("_r3")
        score = top.get("_composite_score", 0)
        signal = "⭐ Strong Buy" if score >= 75 else "✅ Buy" if score >= 55 else "⚠️ Watch"

        row_bg = COLORS["alt_row"] if i % 2 == 0 else "FFFFFF"
        vals = [cat, name, amc, pct(r1), pct(r3), round(score, 1), signal]
        for ci, v in enumerate(vals, 1):
            c = ws_sum.cell(row=i, column=ci, value=v)
            c.font = cell_font(size=9)
            c.fill = fill(row_bg)
            c.border = border
            c.alignment = Alignment(horizontal="center" if ci != 2 else "left",
                                     vertical="center", wrap_text=(ci == 2))
            if ci in (4, 5) and isinstance(v, str) and v != "—":
                num = float(v.replace('%', ''))
                c.font = Font(name="Arial", size=9,
                              color=COLORS["positive"] if num > 0 else COLORS["negative"])
            if ci == 6:
                c.font = Font(name="Arial", bold=True, size=9, color=score_color(score))
        ws_sum.row_dimensions[i].height = 18

    for ci, w in enumerate([28, 52, 22, 12, 12, 10, 14], 1):
        ws_sum.column_dimensions[get_column_letter(ci)].width = w
    ws_sum.freeze_panes = "A4"

    wb.save(output_path)
    print(f"\n✅ Output saved → {output_path}")


def main():
    if not os.path.exists(INPUT_FILE):
        print(f"\n❌ File not found: {INPUT_FILE}")
        print("   Place dashboard_data.xlsx in the same folder as this script and re-run.")
        return

    df = load_data(INPUT_FILE)

    print("\n🔍 Auto-detecting columns...")
    mapping = map_columns(df)

    missing_critical = [k for k in ("scheme_name", "return_1y", "return_3y") if not mapping[k]]
    if missing_critical:
        print(f"\n⚠️  Critical columns not found: {missing_critical}")
        print("   Available columns:", list(df.columns))
        print("   → Update COLUMN_ALIASES at the top of this script to match your column names.")
        return

    print("\n⚙️  Scoring funds...")
    df_scored = score_funds(df, mapping)

    print("\n📝 Building Excel output...")
    build_excel(df_scored, mapping, OUTPUT_FILE)

    print("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print(f"  📊 MF Ranked Screener → {OUTPUT_FILE}")
    print(f"  📁 Open this file to see:")
    print(f"     • 🏆 SUMMARY tab — top fund per category at a glance")
    print(f"     • One tab per fund category — all schemes ranked")
    print(f"     • Signals: ⭐ Strong Buy  ✅ Buy  ⚠️ Watch  ❌ Avoid")
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")


if __name__ == "__main__":
    main()
