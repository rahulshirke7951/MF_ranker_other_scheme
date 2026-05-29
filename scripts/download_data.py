import pandas as pd
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ───────────────────────────────────────────────────────────────────
INPUT_FILE  = "dashboard_data.xlsx"
OUTPUT_FILE = "mf_ranked_screener.xlsx"

QUALITY_FILTERS = {
    "cagr_2y_min": 0.10,
    "cagr_3y_min": 0.12,
}

# Engine 1: Momentum (Short-Term)
ENGINE1_WEIGHTS = {
    "return_6m":  0.30,
    "return_3m":  0.20,
    "return_1y":  0.25,
    "return_1m":  0.25,
}

# Engine 2: Quality (Long-Term)
ENGINE2_WEIGHTS = {
    "return_1y":  0.25,
    "return_2y":  0.30,
    "return_3y":  0.45,
}

COMPOSITE_BLEND = {
    "engine1_momentum": 0.55,
    "engine2_quality":  0.45,
}

TREND_BONUS = 5.0

FILTERS = {
    "cat_level_1": "Open Ended Schemes",
    "cat_level_2": "Other Scheme",
    "plan_type":   "Regular",
    "option_type": "Growth",
}

COLUMN_MAP = {
    "scheme_name": "scheme_name",
    "category":    "cat_level_3",
    "amc":         "amc_name",
    "return_1m":   "return_30d",
    "return_3m":   "return_90d",
    "return_6m":   "return_180d",
    "return_1y":   "return_365d",
    "return_2y":   "return_730d",
    "return_3y":   "return_1095d",
    "nav":         "latest_nav",
}

# ── COLORS ───────────────────────────────────────────────────────────────────
C = {
    "title_bg":      "0D1117",
    "title_fg":      "FFFFFF",
    "info_bg":       "F0F4F8",
    "info_fg":       "445566",
    "border":        "D0D8E4",
    "col_hdr_bg":    "1A3A5C",
    "col_hdr_fg":    "FFFFFF",
    "momentum_bg":   "E8560A",   # Orange
    "momentum_fg":   "FFFFFF",
    "longterm_bg":   "1E6B8C",   # Teal
    "longterm_fg":   "FFFFFF",
    "engine_bg":     "2D5016",   # Forest Green
    "engine_fg":     "FFFFFF",
    "rank1_bg":      "FFD700",
    "rank2_bg":      "E8E8E8",
    "rank3_bg":      "D4956A",
    "alt_row":       "F5F8FC",
    "white":         "FFFFFF",
    "positive":      "1E7A4B",
    "negative":      "C0392B",
    "engine1_tint":  "FFF3E0",   # Amber tint for Momentum
    "engine2_tint":  "E3F2FD",   # Blue tint for Quality
    "comp_tint":     "F0FFF4",   # Light green for Composite
    "asmp_title":    "0D1117",
    "asmp_blend":    "2D5016",
    "asmp_row_bl":   "EAFAF1",
}

# ── DATA LOADING ─────────────────────────────────────────────────────────────
def apply_filters(df):
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for col_key, value in FILTERS.items():
        actual = cols_lower.get(col_key.lower().strip())
        if actual:
            df = df[df[actual].astype(str).str.strip().str.lower() == str(value).strip().lower()]
    return df

def load_data():
    sheets = pd.read_excel(INPUT_FILE, sheet_name=None)
    frames = [df for df in sheets.values() if "scheme_name" in df.columns]
    combined = pd.concat(frames, ignore_index=True)
    return apply_filters(combined)

# ── SCORING ENGINE ───────────────────────────────────────────────────────────
def to_num(series):
    s = series.astype(str).str.replace('%', '').str.replace(',', '').str.strip()
    return pd.to_numeric(s, errors='coerce')

def pct_rank(series):
    if series.dropna().empty:
        return series.fillna(0.0)
    ranks = series.rank(method='min', na_option='bottom')
    return (ranks - 1) / max(len(series) - 1, 1) * 100

def cagr_2y(r):
    if pd.isna(r) or r <= -100: return float('nan')
    return ((1 + float(r) / 100) ** 0.5 - 1) * 100

def score_funds(df):
    df = df.copy()
    df["_r1m"]     = to_num(df[COLUMN_MAP["return_1m"]])
    df["_r3m"]     = to_num(df[COLUMN_MAP["return_3m"]])
    df["_r6m"]     = to_num(df[COLUMN_MAP["return_6m"]])
    df["_r1y"]     = to_num(df[COLUMN_MAP["return_1y"]])
    df["_r2y_raw"] = to_num(df[COLUMN_MAP["return_2y"]])
    df["_r2y_cagr"]= df["_r2y_raw"].apply(cagr_2y)
    df["_r3y"]     = to_num(df[COLUMN_MAP["return_3y"]])
    df["_cat"]     = df[COLUMN_MAP["category"]].astype(str).str.strip().str.title()

    # Dynamic protection wrapper for blank data cells
    essential_returns = ["_r1m", "_r3m", "_r6m", "_r1y", "_r2y_cagr", "_r3y"]
    df["_has_missing_data"] = df[essential_returns].isna().any(axis=1)

    # Asset tag filtering (Gold & Silver vs Standard Equity/Debt)
    is_metal = df[COLUMN_MAP["scheme_name"]].astype(str).str.lower().str.contains("gold|silver", regex=True)
    df["_asset_class"] = df.map(lambda x: "Gold & Silver" if is_metal[x.name] else "Standard Equity/Debt", axis=1)

    # Quality rules setup
    mask_2y = df["_r2y_cagr"] > (QUALITY_FILTERS["cagr_2y_min"] * 100)
    mask_3y = df["_r3y"]      > (QUALITY_FILTERS["cagr_3y_min"] * 100)
    df["_qualifies"] = mask_2y & mask_3y & (~df["_has_missing_data"])

    # Progressive momentum calculation
    trend = (df["_r6m"] > df["_r3m"]) & (df["_r3m"] > df["_r1m"])
    df["_trend"] = trend.map({True: "📈 Uptrend", False: ""})

    df["_e1"] = 0.0
    df["_e2"] = 0.0

    for cat in df["_cat"].unique():
        cm = df["_cat"] == cat
        cq = cm & df["_qualifies"]

        # Engine 1: Momentum
        valid_m_mask = cm & (~df["_has_missing_data"])
        if valid_m_mask.sum() > 0:
            e1 = (pct_rank(df.loc[valid_m_mask, "_r6m"]) * ENGINE1_WEIGHTS["return_6m"] +
                  pct_rank(df.loc[valid_m_mask, "_r3m"]) * ENGINE1_WEIGHTS["return_3m"] +
                  pct_rank(df.loc[valid_m_mask, "_r1y"]) * ENGINE1_WEIGHTS["return_1y"] +
                  pct_rank(df.loc[valid_m_mask, "_r1m"]) * ENGINE1_WEIGHTS["return_1m"])
            
            is_uptrend = (df.loc[valid_m_mask, "_trend"] == "📈 Uptrend").astype(float)
            e1 = e1 + (is_uptrend * TREND_BONUS)
            df.loc[valid_m_mask, "_e1"] = e1.clip(0, 100)

        # Engine 2: Quality
        if cq.sum() > 0:
            e2 = (pct_rank(df.loc[cq, "_r1y"])      * ENGINE2_WEIGHTS["return_1y"] +
                  pct_rank(df.loc[cq, "_r2y_cagr"]) * ENGINE2_WEIGHTS["return_2y"] +
                  pct_rank(df.loc[cq, "_r3y"])      * ENGINE2_WEIGHTS["return_3y"])
            df.loc[cq, "_e2"] = e2

    df["_comp"] = (df["_e1"] * COMPOSITE_BLEND["engine1_momentum"] +
                   df["_e2"] * COMPOSITE_BLEND["engine2_quality"])

    # Force incomplete/broken data to drop bottom out
    df.loc[df["_has_missing_data"], "_comp"] = -1.0

    df["_rank"] = (df.groupby("_cat")["_comp"]
                     .rank(method='min', ascending=False)
                     .astype(int))
    return df

# ── FORMATTING HELPERS ────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def border():
    s = Side(style='thin', color=C["border"])
    return Border(left=s, right=s, top=s, bottom=s)

def fmt_pct(val):
    if pd.isna(val) or val is None: return "—"
    try: return f"{float(val):+.2f}%"
    except: return "—"

def score_col(val):
    try: 
        v = float(val)
        if v < 0: return "888888"
    except: 
        return "888888"
    if v >= 75: return "1E7A4B"
    if v >= 50: return "E67E00"
    return "C0392B"

def signal(comp):
    if comp < 0: return "❌ Missing Data"
    if comp >= 75: return "⭐ Strong Buy"
    if comp >= 55: return "✅ Buy"
    return "⚠️ Watch"

def clean_name(name):
    return re.sub(r'[\\/*?:\[\]]', '', str(name))[:31]

def hfont(size=9, color="FFFFFF", bold=True):
    return Font(name="Arial", bold=bold, size=size, color=color)

def dfont(size=9, bold=False, color="000000"):
    return Font(name="Arial", bold=bold, size=size, color=color)

# Swapped locations adjusted for Asset Class insertion (Asset Class is Column 4)
MOMENTUM_COLS  = (5, 7)   # 1M to 6M
LONGTERM_COLS  = (8, 10)  # 1Y to 3Y CAGR
ENGINE_COLS    = (11, 13) # E1 to Composite

COL_HEADERS = [
    "Rank", "Scheme Name", "AMC", "Asset Class",
    "1M\nReturn", "3M\nReturn", "6M\nReturn",
    "1Y\nReturn", "2Y\nCAGR", "3Y\nCAGR",
    "Engine 1\n(Momentum)", "Engine 2\n(Quality)", "Composite\nScore",
]

COL_WIDTHS = [6, 50, 20, 22, 9, 9, 9, 9, 9, 9, 14, 14, 12]
RETURN_COLS_IDX = {5, 6, 7, 8, 9, 10}

# ── BUILD A CATEGORY SHEET ───────────────────────────────────────────────────
def build_category_sheet(wb, cat, cat_df):
    ws = wb.create_sheet(clean_name(cat))
    bd = border()
    total = len(cat_df)
    qualified = int(cat_df["_qualifies"].sum())
    ncols = len(COL_HEADERS)

    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    c = ws["A1"]
    c.value     = f"✦  {cat.upper()}  ✦"
    c.font      = Font(name="Arial", bold=True, size=14, color=C["title_fg"])
    c.fill      = fill(C["title_bg"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    ws.merge_cells(f"A2:{get_column_letter(ncols)}2")
    c = ws["A2"]
    e1_pct = int(COMPOSITE_BLEND["engine1_momentum"] * 100)
    e2_pct = int(COMPOSITE_BLEND["engine2_quality"] * 100)
    c.value     = f"Dual Engine Model: {e1_pct}% Momentum (ST) + {e2_pct}% Quality (LT) | Passed Gates: {qualified}/{total}"
    c.font      = Font(name="Arial", italic=True, size=8.5, color=C["info_fg"])
    c.fill      = fill(C["info_bg"])
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 16

    for ci in range(1, 5):
        ws.cell(row=3, column=ci).fill = fill(C["col_hdr_bg"])
        ws.cell(row=3, column=ci).border = bd

    for g_start, g_end, label, bg, fg in [
        (MOMENTUM_COLS[0], MOMENTUM_COLS[1], "◄  Momentum  ►", C["momentum_bg"], C["momentum_fg"]),
        (LONGTERM_COLS[0], LONGTERM_COLS[1], "◄  Long-Term  ►", C["longterm_bg"], C["longterm_fg"]),
        (ENGINE_COLS[0], ENGINE_COLS[1], "◄  Engine Scores  ►", C["engine_bg"], C["engine_fg"])
    ]:
        ws.merge_cells(f"{get_column_letter(g_start)}3:{get_column_letter(g_end)}3")
        cell = ws.cell(row=3, column=g_start, value=label)
        cell.font = hfont(color=fg); cell.fill = fill(bg); cell.alignment = Alignment(horizontal="center", vertical="center")
        for ci in range(g_start, g_end + 1):
            ws.cell(row=3, column=ci).fill = fill(bg)
            ws.cell(row=3, column=ci).border = bd
    ws.row_dimensions[3].height = 18

    for ci, hdr in enumerate(COL_HEADERS, 1):
        c = ws.cell(row=4, column=ci, value=hdr)
        c.font, c.fill, c.alignment, c.border = hfont(), fill(C["col_hdr_bg"]), Alignment(horizontal="center", vertical="center", wrap_text=True), bd
    ws.row_dimensions[4].height = 28

    for i, (_, row) in enumerate(cat_df.iterrows(), 5):
        rank = row["_rank"]
        is_missing = row["_has_missing_data"]
        
        if is_missing:
            rbg = C["white"]
        else:
            rbg = C["rank1_bg"] if rank == 1 else (C["rank2_bg"] if rank == 2 else (C["rank3_bg"] if rank == 3 else (C["alt_row"] if i % 2 == 0 else C["white"])))
        
        vals = [
            "—" if is_missing else rank, 
            row.get(COLUMN_MAP["scheme_name"], "—"), 
            row.get(COLUMN_MAP["amc"], "—"),
            row["_asset_class"],
            fmt_pct(row["_r1m"]), fmt_pct(row["_r3m"]), fmt_pct(row["_r6m"]),
            fmt_pct(row["_r1y"]), fmt_pct(row["_r2y_cagr"]), fmt_pct(row["_r3y"]),
            "—" if is_missing else round(row["_e1"], 1), 
            "—" if is_missing else round(row["_e2"], 1), 
            "—" if is_missing else round(row["_comp"], 1)
        ]

        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=ci, value=val)
            c.border, c.fill, c.alignment = bd, fill(rbg), (Alignment(horizontal="left", vertical="center") if ci in {2, 3, 4} else Alignment(horizontal="center", vertical="center"))
            
            if ci in RETURN_COLS_IDX and isinstance(val, str) and val != "—":
                num = float(val.replace('%', ''))
                c.font = Font(name="Arial", bold=(rank <= 3 and not is_missing), size=9, color=C["positive"] if num >= 0 else C["negative"])
            elif ci == 11 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val)); c.fill = fill(C["engine1_tint"])
            elif ci == 12 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val)); c.fill = fill(C["engine2_tint"])
            elif ci == 13 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val)); c.fill = fill(C["comp_tint"])
            else:
                c.font = dfont(bold=(rank <= 3 and not is_missing))
                if is_missing:
                    c.font = Font(name="Arial", italic=True, size=9, color="888888")
        ws.row_dimensions[i].height = 16

    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    
    # Auto-enable Filter Mode on Category Sheets
    ws.auto_filter.ref = f"A4:{get_column_letter(ncols)}{4 + len(cat_df)}"
    ws.freeze_panes = "A5"

# ── SUMMARY SHEET ─────────────────────────────────────────────────────────────
def build_summary(wb, df_scored, categories):
    ws = wb.create_sheet("🏆 SUMMARY", 0)
    bd = border()
    ncols = len(COL_HEADERS) + 1  

    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    c = ws["A1"]
    c.value     = "MF INTELLIGENCE — UNIFIED DUAL ENGINE RANKED SCREENER"
    c.font      = Font(name="Arial", bold=True, size=15, color=C["title_fg"])
    c.fill      = fill(C["title_bg"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    ws.merge_cells(f"A2:{get_column_letter(ncols)}2")
    c = ws["A2"]
    e1_pct = int(COMPOSITE_BLEND["engine1_momentum"] * 100)
    e2_pct = int(COMPOSITE_BLEND["engine2_quality"] * 100)
    c.value     = f"Composite Master Framework: {e1_pct}% Engine 1 (Momentum) + {e2_pct}% Engine 2 (Quality Setup) | Filter via 'Asset Class' column"
    c.font = Font(name="Arial", italic=True, size=8.5, color=C["info_fg"]); c.fill = fill(C["info_bg"]); c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 16

    for ci in range(1, 5):
        ws.cell(row=3, column=ci).fill = fill(C["col_hdr_bg"]); ws.cell(row=3, column=ci).border = bd

    for g_start, g_end, label, bg, fg in [
        (MOMENTUM_COLS[0], MOMENTUM_COLS[1], "◄  Momentum  ►", C["momentum_bg"], C["momentum_fg"]),
        (LONGTERM_COLS[0], LONGTERM_COLS[1], "◄  Long-Term  ►", C["longterm_bg"], C["longterm_fg"]),
        (ENGINE_COLS[0], ENGINE_COLS[1], "◄  Engine Scores  ►", C["engine_bg"], C["engine_fg"])
    ]:
        ws.merge_cells(f"{get_column_letter(g_start)}3:{get_column_letter(g_end)}3")
        cell = ws.cell(row=3, column=g_start, value=label)
        cell.font = hfont(color=fg); cell.fill = fill(bg); cell.alignment = Alignment(horizontal="center", vertical="center")
        for ci in range(g_start, g_end + 1):
            ws.cell(row=3, column=ci).fill = fill(bg); ws.cell(row=3, column=ci).border = bd
    ws.row_dimensions[3].height = 18

    for ci, hdr in enumerate(COL_HEADERS, 1):
        c = ws.cell(row=4, column=ci, value=hdr)
        c.font, c.fill, c.alignment, c.border = hfont(), fill(C["col_hdr_bg"]), Alignment(horizontal="center", vertical="center", wrap_text=True), bd

    sig_col = len(COL_HEADERS) + 1
    ws.cell(row=4, column=sig_col, value="Signal").font = hfont()
    ws.cell(row=4, column=sig_col).fill = fill(C["col_hdr_bg"])
    ws.cell(row=4, column=sig_col).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=4, column=sig_col).border = bd
    ws.column_dimensions[get_column_letter(sig_col)].width = 14

    for i, cat in enumerate(categories, 5):
        cat_df = df_scored[df_scored["_cat"] == cat].sort_values("_rank")
        if cat_df.empty: continue
        top = cat_df.iloc[0]
        rbg = C["alt_row"] if i % 2 == 0 else C["white"]
        is_missing = top["_has_missing_data"]

        vals = [
            cat, top.get(COLUMN_MAP["scheme_name"], "—"), top.get(COLUMN_MAP["amc"], "—"),
            top["_asset_class"],
            fmt_pct(top["_r1m"]), fmt_pct(top["_r3m"]), fmt_pct(top["_r6m"]),
            fmt_pct(top["_r1y"]), fmt_pct(top["_r2y_cagr"]), fmt_pct(top["_r3y"]),
            "—" if is_missing else round(top["_e1"], 1), 
            "—" if is_missing else round(top["_e2"], 1), 
            "—" if is_missing else round(top["_comp"], 1)
        ]

        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=i, column=ci, value=val)
            c.border, c.fill, c.alignment = bd, fill(rbg), (Alignment(horizontal="left", vertical="center") if ci in {1, 2, 3, 4} else Alignment(horizontal="center", vertical="center"))
            
            if ci in RETURN_COLS_IDX and isinstance(val, str) and val != "—":
                num = float(val.replace('%', ''))
                c.font = Font(name="Arial", size=9, color=C["positive"] if num >= 0 else C["negative"])
            elif ci == 11 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val)); c.fill = fill(C["engine1_tint"])
            elif ci == 12 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val)); c.fill = fill(C["engine2_tint"])
            elif ci == 13 and not is_missing:
                c.font = Font(name="Arial", bold=True, size=9, color=score_col(val)); c.fill = fill(C["comp_tint"])
            else:
                c.font = dfont()

        c_sig = ws.cell(row=i, column=sig_col, value=signal(top["_comp"]))
        c_sig.font, c_sig.fill, c_sig.border, c_sig.alignment = Font(name="Arial", bold=True, size=9), fill(rbg), bd, Alignment(horizontal="center", vertical="center")
        if is_missing:
            c_sig.font = Font(name="Arial", italic=True, size=9, color="888888")
        ws.row_dimensions[i].height = 18

    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
        
    # Auto-enable Filter Mode on Master Summary Sheet
    ws.auto_filter.ref = f"A4:{get_column_letter(ncols)}{4 + len(categories)}"
    ws.freeze_panes = "A5"

# ── ASSUMPTIONS SHEET ─────────────────────────────────────────────────────────
def build_assumptions(wb, df_scored):
    ws = wb.create_sheet("📋 ASSUMPTIONS", 1)
    bd = border()

    def section(row, title, bg, fg="FFFFFF", ncols=3):
        ws.merge_cells(f"A{row}:{get_column_letter(ncols)}{row}")
        c = ws[f"A{row}"]
        c.value = title; c.font = Font(name="Arial", bold=True, size=11, color=fg); c.fill = fill(bg); c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 22
        return row + 1

    def kv_row(row, key, val, desc="", row_bg="FFFFFF"):
        for ci, text in [(1, key), (2, val), (3, desc)]:
            c = ws.cell(row=row, column=ci, value=text)
            c.font, c.fill, c.border = Font(name="Arial", size=9), fill(row_bg), bd
            c.alignment = Alignment(horizontal="left", vertical="center") if ci in {1, 3} else Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 15
        return row + 1

    def col_headers(row, headers, bg):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.font, c.fill, c.border, c.alignment = Font(name="Arial", bold=True, size=9, color="FFFFFF"), fill(bg), bd, Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 16
        return row + 1

    ws.merge_cells("A1:C1")
    ws["A1"] = "DUAL ENGINE MODEL — ASSUMPTIONS & METHODOLOGY"
    ws["A1"].font, ws["A1"].fill, ws["A1"].alignment = Font(name="Arial", bold=True, size=14, color="FFFFFF"), fill(C["asmp_title"]), Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:C2")
    ws["A2"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Unified Asset Class Filtering Implemented"
    ws["A2"].font, ws["A2"].fill, ws["A2"].alignment = Font(name="Arial", italic=True, size=8.5, color=C["info_fg"]), fill(C["info_bg"]), Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 14

    r = 4
    r = section(r, "⚖  COMPOSITE BLEND WEIGHTS", C["asmp_blend"])
    r = col_headers(r, ["Parameter", "Value", "Description"], C["asmp_blend"])
    r = kv_row(r, "Engine 1 Weight (Momentum)", f"{int(COMPOSITE_BLEND['engine1_momentum']*100)}%", "Short-term momentum configuration weight", C["asmp_row_bl"])
    r = kv_row(r, "Engine 2 Weight (Quality)", f"{int(COMPOSITE_BLEND['engine2_quality']*100)}%", "Long-term core health score weight", C["asmp_row_bl"])
    
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 44

# ── MAIN PIPELINE EXECUTOR ───────────────────────────────────────────────────
def main():
    print("Loading data matrix...")
    df = load_data()
    
    print("Executing Scoring Engine and assigning Asset Tags...")
    df_scored = score_funds(df)

    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    categories = sorted(df_scored["_cat"].unique())

    print("Building Unified Summary Sheet with Excel Auto-Filters...")
    build_summary(wb, df_scored, categories)

    print("Writing structural configurations...")
    build_assumptions(wb, df_scored)

    print("Generating categorical breakout sheets...")
    for cat in categories:
        cat_df = df_scored[df_scored["_cat"] == cat].sort_values("_rank")
        if cat_df.empty: continue
        build_category_sheet(wb, cat, cat_df)

    wb.save(OUTPUT_FILE)
    print(f"\n✅ Done! Open '{OUTPUT_FILE}' and look for the 'Asset Class' filter dropdown.")

if __name__ == "__main__":
    main()
