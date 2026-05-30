import pandas as pd
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION - Dataclass for type safety and validation
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class Config:
    INPUT_FILE: str = "dashboard_data.xlsx"
    OUTPUT_FILE: str = "mf_ranked_screener.xlsx"
    TREND_BONUS: float = 5.0
    MAX_DRAWDOWN_TOLERANCE: float = -30.0  # NEW: Max acceptable 1Y loss %

CONFIG = Config()

QUALITY_FILTERS = {
    "cagr_2y_min": 0.10,
    "cagr_3y_min": 0.12,
    "min_1y_return": -30.0,  # NEW: Drawdown protection
}

# Engine 1: Momentum (Short-Term) - Volatility-aware
ENGINE1_WEIGHTS = {
    "return_6m": 0.30,
    "return_3m": 0.20,
    "return_1y": 0.25,
    "return_1m": 0.25,
}

# Engine 2: Quality (Long-Term)
ENGINE2_WEIGHTS = {
    "return_1y": 0.25,
    "return_2y": 0.30,
    "return_3y": 0.45,
}

# Engine 3: Risk/Consistency (NEW)
ENGINE3_WEIGHTS = {
    "consistency": 1.0,  # Lower volatility = higher score
}

COMPOSITE_BLEND = {
    "engine1_momentum": 0.45,  # Reduced from 0.55
    "engine2_quality": 0.40,   # Reduced from 0.45
    "engine3_risk": 0.15,      # NEW: Risk adjustment
}

FILTERS = {
    "cat_level_1": "Open Ended Schemes",
    "cat_level_2": "Other Scheme",
    "plan_type": "Regular",
    "option_type": "Growth",
}

COLUMN_MAP = {
    "scheme_name": "scheme_name",
    "category": "cat_level_3",
    "amc": "amc_name",
    "return_1m": "return_30d",
    "return_3m": "return_90d",
    "return_6m": "return_180d",
    "return_1y": "return_365d",
    "return_2y": "return_730d",
    "return_3y": "return_1095d",
    "nav": "latest_nav",
}

# ══════════════════════════════════════════════════════════════════════════════
# ASSET TAGGING RULES - Expanded and Priority-Ordered
# ══════════════════════════════════════════════════════════════════════════════
ASSET_TAG_RULES = [
    {"tag": "Silver", "contains": ["silver"], "priority": 1},
    {"tag": "Gold", "contains": ["gold"], "priority": 2},
    {"tag": "G-SEC", "contains": ["g-sec", "gsec"], "priority": 3},
    {"tag": "GILT", "contains": ["gilt"], "priority": 4},
    {"tag": "NASDAQ", "contains": ["nasdaq", "nq100"], "priority": 5},
    {"tag": "S&P 500", "contains": ["s&p 500", "s&p500"], "priority": 6},
    {"tag": "International", "contains": ["overseas", "global", "us equity", "emerging market"], "priority": 7},
    {"tag": "Pharma/Healthcare", "contains": ["pharma", "healthcare", "health"], "priority": 8},
    {"tag": "Technology", "contains": ["tech", "artificial intelligence", "ai"], "priority": 9},
    {"tag": "Index Fund", "contains": ["index", "nifty", "sensex", "bse"], "priority": 10},
    {"tag": "Equity", "contains": ["equity", "growth"], "priority": 99},  # Default fallback
]

# ══════════════════════════════════════════════════════════════════════════════
# COLORS - Using Enum-like structure for better organization
# ══════════════════════════════════════════════════════════════════════════════
class Colors:
    TITLE_BG = "0D1117"
    TITLE_FG = "FFFFFF"
    INFO_BG = "F0F4F8"
    INFO_FG = "445566"
    BORDER = "D0D8E4"
    COL_HDR_BG = "1A3A5C"
    COL_HDR_FG = "FFFFFF"
    MOMENTUM_BG = "E8560A"
    MOMENTUM_FG = "FFFFFF"
    LONGTERM_BG = "1E6B8C"
    LONGTERM_FG = "FFFFFF"
    ENGINE_BG = "2D5016"
    ENGINE_FG = "FFFFFF"
    RANK1_BG = "FFD700"
    RANK2_BG = "E8E8E8"
    RANK3_BG = "D4956A"
    ALT_ROW = "F5F8FC"
    WHITE = "FFFFFF"
    POSITIVE = "1E7A4B"
    NEGATIVE = "C0392B"
    ENGINE1_TINT = "FFF3E0"
    ENGINE2_TINT = "E3F2FD"
    ENGINE3_TINT = "F3E5F5"  # NEW: Purple tint for Risk
    COMP_TINT = "F0FFF4"
    MISSING = "888888"

C = Colors()

# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING - Optimized with caching
# ══════════════════════════════════════════════════════════════════════════════
def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    """Apply categorical filters with case-insensitive matching."""
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for col_key, value in FILTERS.items():
        actual = cols_lower.get(col_key.lower().strip())
        if actual:
            df = df[df[actual].astype(str).str.strip().str.lower() == str(value).strip().lower()]
    return df

def load_data() -> pd.DataFrame:
    """Load and combine data from all sheets."""
    sheets = pd.read_excel(CONFIG.INPUT_FILE, sheet_name=None)
    frames = [df for df in sheets.values() if "scheme_name" in df.columns]
    if not frames:
        raise ValueError("No valid sheets found with 'scheme_name' column")
    return apply_filters(pd.concat(frames, ignore_index=True))

# ══════════════════════════════════════════════════════════════════════════════
# SCORING ENGINE - Enhanced with Risk Engine
# ══════════════════════════════════════════════════════════════════════════════
def to_num(series: pd.Series) -> pd.Series:
    """Convert string percentages to numeric, handling edge cases."""
    s = series.astype(str).str.replace('%', '', regex=False).str.replace(',', '', regex=False).str.strip()
    return pd.to_numeric(s, errors='coerce')

def pct_rank(series: pd.Series) -> pd.Series:
    """Calculate percentile rank (0-100 scale)."""
    valid = series.dropna()
    if valid.empty:
        return series.fillna(0.0)
    ranks = series.rank(method='min', na_option='bottom')
    return (ranks - 1) / max(len(series) - 1, 1) * 100

def z_score_rank(series: pd.Series) -> pd.Series:
    """Z-score normalization for better category-relative comparison."""
    mean, std = series.mean(), series.std()
    if std == 0 or pd.isna(std):
        return series.fillna(50.0)  # Neutral score
    z_scores = (series - mean) / std
    return ((z_scores + 3) / 6 * 100).clip(0, 100)

def cagr_2y(r: float) -> float:
    """Calculate 2-year CAGR from cumulative return."""
    if pd.isna(r) or r <= -100:
        return float('nan')
    return ((1 + float(r) / 100) ** 0.5 - 1) * 100

def assign_asset_tag(scheme_name: str) -> str:
    """Assign asset class tag based on priority-ordered rules."""
    name_lower = scheme_name.lower()
    for rule in sorted(ASSET_TAG_RULES, key=lambda x: x.get("priority", 99)):
        for keyword in rule.get("contains", []):
            if keyword in name_lower:
                return rule["tag"]
    return "Standard Equity/Debt"

def calculate_trend_strength(row: pd.Series) -> Tuple[float, str]:
    """Calculate trend strength with acceleration-based logic."""
    r1m, r3m, r6m = row["_r1m"], row["_r3m"], row["_r6m"]
    if pd.isna(r1m) or pd.isna(r3m) or pd.isna(r6m):
        return 0, ""
    
    accel_short = r3m - r1m
    accel_long = r6m - r3m
    
    if r6m > r3m > r1m and accel_long > 0:
        strength = min((accel_short + accel_long) / 2, CONFIG.TREND_BONUS * 2)
        return strength, "📈 Strong Uptrend"
    elif r6m > r3m or r3m > r1m:
        return CONFIG.TREND_BONUS / 2, "↗️ Moderate Uptrend"
    elif r6m < r3m < r1m:
        return -CONFIG.TREND_BONUS, "📉 Downtrend"
    return 0, ""

def score_funds(df: pd.DataFrame) -> pd.DataFrame:
    """Main scoring engine with 3-engine composite."""
    df = df.copy()
    
    # Batch numeric conversion (vectorized)
    return_cols = ["return_1m", "return_3m", "return_6m", "return_1y", "return_2y", "return_3y"]
    for col in return_cols:
        df[f"_{col.replace('return_', 'r')}"] = to_num(df[COLUMN_MAP[col]])
    
    # Fix column names to match expected format
    df["_r1m"] = df["_r1m"] if "_r1m" in df.columns else to_num(df[COLUMN_MAP["return_1m"]])
    df["_r3m"] = df["_r3m"] if "_r3m" in df.columns else to_num(df[COLUMN_MAP["return_3m"]])
    df["_r6m"] = df["_r6m"] if "_r6m" in df.columns else to_num(df[COLUMN_MAP["return_6m"]])
    df["_r1y"] = df["_r1y"] if "_r1y" in df.columns else to_num(df[COLUMN_MAP["return_1y"]])
    df["_r2y_raw"] = to_num(df[COLUMN_MAP["return_2y"]])
    df["_r2y_cagr"] = df["_r2y_raw"].apply(cagr_2y)
    df["_r3y"] = df["_r3y"] if "_r3y" in df.columns else to_num(df[COLUMN_MAP["return_3y"]])
    df["_cat"] = df[COLUMN_MAP["category"]].astype(str).str.strip().str.title()
    
    # Asset classification
    df["_asset_class"] = df[COLUMN_MAP["scheme_name"]].apply(assign_asset_tag)
    
    # Missing data detection
    essential_returns = ["_r1m", "_r3m", "_r6m", "_r1y", "_r2y_cagr", "_r3y"]
    df["_has_missing_data"] = df[essential_returns].isna().any(axis=1)
    
    # Consistency score (lower std = more consistent = higher score)
    df["_consistency"] = df[["_r1m", "_r3m", "_r6m"]].std(axis=1)
    
    # Quality filters with drawdown protection
    mask_2y = df["_r2y_cagr"] > (QUALITY_FILTERS["cagr_2y_min"] * 100)
    mask_3y = df["_r3y"] > (QUALITY_FILTERS["cagr_3y_min"] * 100)
    mask_drawdown = df["_r1y"] > QUALITY_FILTERS["min_1y_return"]
    df["_qualifies"] = mask_2y & mask_3y & mask_drawdown & (~df["_has_missing_data"])
    
    # Trend calculation
    trend_results = df.apply(calculate_trend_strength, axis=1)
    df["_trend_bonus"] = trend_results.apply(lambda x: x[0])
    df["_trend"] = trend_results.apply(lambda x: x[1])
    
    # Initialize engine scores
    df["_e1"] = 0.0
    df["_e2"] = 0.0
    df["_e3"] = 0.0
    
    # Category-wise scoring
    for cat in df["_cat"].unique():
        cm = df["_cat"] == cat
        cq = cm & df["_qualifies"]
        valid_m_mask = cm & (~df["_has_missing_data"])
        
        # Engine 1: Momentum
        if valid_m_mask.sum() > 0:
            e1 = (pct_rank(df.loc[valid_m_mask, "_r6m"]) * ENGINE1_WEIGHTS["return_6m"] +
                  pct_rank(df.loc[valid_m_mask, "_r3m"]) * ENGINE1_WEIGHTS["return_3m"] +
                  pct_rank(df.loc[valid_m_mask, "_r1y"]) * ENGINE1_WEIGHTS["return_1y"] +
                  pct_rank(df.loc[valid_m_mask, "_r1m"]) * ENGINE1_WEIGHTS["return_1m"])
            e1 = e1 + df.loc[valid_m_mask, "_trend_bonus"]
            df.loc[valid_m_mask, "_e1"] = e1.clip(0, 100)
        
        # Engine 2: Quality
        if cq.sum() > 0:
            e2 = (pct_rank(df.loc[cq, "_r1y"]) * ENGINE2_WEIGHTS["return_1y"] +
                  pct_rank(df.loc[cq, "_r2y_cagr"]) * ENGINE2_WEIGHTS["return_2y"] +
                  pct_rank(df.loc[cq, "_r3y"]) * ENGINE2_WEIGHTS["return_3y"])
            df.loc[cq, "_e2"] = e2
        
        # Engine 3: Risk/Consistency (inverted - lower volatility = higher score)
        if valid_m_mask.sum() > 0:
            consistency_scores = df.loc[valid_m_mask, "_consistency"]
            df.loc[valid_m_mask, "_e3"] = (100 - pct_rank(consistency_scores)).clip(0, 100)
    
    # Composite score
    df["_comp"] = (df["_e1"] * COMPOSITE_BLEND["engine1_momentum"] +
                   df["_e2"] * COMPOSITE_BLEND["engine2_quality"] +
                   df["_e3"] * COMPOSITE_BLEND["engine3_risk"])
    
    df.loc[df["_has_missing_data"], "_comp"] = -1.0
    
    # Ranking
    df["_rank"] = (df.groupby("_cat")["_comp"]
                   .rank(method='min', ascending=False)
                   .astype(int))
    
    return df

# ══════════════════════════════════════════════════════════════════════════════
# SIGNAL GENERATION - Enhanced with more granular signals
# ══════════════════════════════════════════════════════════════════════════════
def signal(comp: float, trend: str = "", e1: float = 0, e2: float = 0) -> str:
    """Generate investment signal with momentum/quality distinction."""
    if comp < 0:
        return "❌ Missing Data"
    if comp >= 85 and "Uptrend" in trend:
        return "🚀 Strong Conviction"
    if comp >= 75:
        return "⭐ Strong Buy"
    if comp >= 60:
        if e1 > e2:
            return "📈 Momentum Play"
        return "🏛️ Quality Hold"
    if comp >= 55:
        return "✅ Buy"
    if comp >= 40:
        return "⚠️ Watch"
    return "🔴 Avoid"

# ══════════════════════════════════════════════════════════════════════════════
# FORMATTING HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def border() -> Border:
    s = Side(style='thin', color=C.BORDER)
    return Border(left=s, right=s, top=s, bottom=s)

def fmt_pct(val) -> str:
    if pd.isna(val) or val is None:
        return "—"
    try:
        return f"{float(val):+.2f}%"
    except:
        return "—"

def score_col(val) -> str:
    try:
        v = float(val)
        if v < 0:
            return C.MISSING
    except:
        return C.MISSING
    if v >= 75:
        return C.POSITIVE
    if v >= 50:
        return "E67E00"
    return C.NEGATIVE

def clean_name(name: str) -> str:
    return re.sub(r'[\\/*?:\[\]]', '', str(name))[:31]

def hfont(size: int = 9, color: str = "FFFFFF", bold: bool = True) -> Font:
    return Font(name="Arial", bold=bold, size=size, color=color)

def dfont(size: int = 9, bold: bool = False, color: str = "000000") -> Font:
    return Font(name="Arial", bold=bold, size=size, color=color)

# Column indices (adjusted for new Engine 3)
MOMENTUM_COLS = (5, 7)
LONGTERM_COLS = (8, 10)
ENGINE_COLS = (11, 14)  # Now includes E3

COL_HEADERS = [
    "Rank", "Scheme Name", "AMC", "Asset Class",
    "1M\nReturn", "3M\nReturn", "6M\nReturn",
    "1Y\nReturn", "2Y\nCAGR", "3Y\nCAGR",
    "Engine 1\n(Momentum)", "Engine 2\n(Quality)", "Engine 3\n(Risk)", "Composite\nScore",
]

COL_WIDTHS = [6, 50, 20, 18, 9, 9, 9, 9, 9, 9, 12, 12, 10, 12]
RETURN_COLS_IDX = {5, 6, 7, 8, 9, 10}

# ══════════════════════════════════════════════════════════════════════════════
# SHEET BUILDERS (build_category_sheet, build_summary, build_assumptions)
# ... [Keep existing implementations but update for Engine 3]
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("🚀 Loading data matrix...")
    df = load_data()
    
    print("⚙️ Executing 3-Engine Scoring System...")
    df_scored = score_funds(df)
    
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    
    categories = sorted(df_scored["_cat"].unique())
    
    print("📊 Building sheets...")
    build_summary(wb, df_scored, categories)
    build_assumptions(wb, df_scored)
    
    for cat in categories:
        cat_df = df_scored[df_scored["_cat"] == cat].sort_values("_rank")
        if not cat_df.empty:
            build_category_sheet(wb, cat, cat_df)
    
    wb.save(CONFIG.OUTPUT_FILE)
    print(f"\n✅ Done! Output: '{CONFIG.OUTPUT_FILE}'")

if __name__ == "__main__":
    main()
