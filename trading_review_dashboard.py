# -*- coding: utf-8 -*-
"""
江苏电力交易量价复盘平台 V3.2.11

本版优化：
1. 标题信息改为紧凑单行，避免异常空行
2. 展示日期按北京时间所属交易日识别，不再把“6月30日交易、7月1日执行”显示成6月30日
3. 删除基于绝对交易量计算的饼图和失真占比
4. 交易概览改为：净量柱状图 + 交易形成瀑布图 + 阶段量价摘要表
5. 顶部KPI改为净量、调整量和加权均价
6. Tab等宽铺满并优化间距
7. 删除能量块页面中无上下文的预测说明
8. 新增“中长期偏差预警”页面
9. 支持中长期覆盖率、70%/105%预警、回收费用区间估算及单值估算
10. 全省中长期均价默认支持330–360元/MWh区间，也可手动填写单一价格
11. Excel修改时间加入缓存Key，避免读取旧缓存
12. 修复能量块复盘页在部分标准化文件中无内容的问题
13. 交易概览增加峰、平、谷三个成交加权均价，价格口径跟随侧栏选择

运行：
streamlit run trading_review_dashboard.py
"""

import io
import re
from pathlib import Path
from typing import Dict, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =========================================================
# 1. 页面配置
# =========================================================
st.set_page_config(
    page_title="江苏电力交易量价复盘平台",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# 2. 路径与常量
# =========================================================
BASE_DIR = Path("data")

OUTPUT_FILE_PATTERN = "trading_review_standardized_*.xlsx"

MARKET_ORDER = ["年协/月协", "能量块", "日前", "实时"]

MARKET_CODE_TO_CN = {
    "long_term": "年协/月协",
    "energy_block": "能量块",
    "day_ahead": "日前",
    "real_time": "实时",
}


# =========================================================
# 3. 页面样式
# =========================================================
st.markdown(
    """
<style>
.stApp {
    background: #ffffff;
}

[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid #e5e7eb;
}

.main .block-container {
    padding-top: 1.15rem;
    padding-bottom: 2rem;
    max-width: 1600px;
}

.dashboard-title {
    font-size: 2rem;
    font-weight: 760;
    color: #172033;
    margin-bottom: 0.3rem;
}

.dashboard-meta {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 0.35rem 0.8rem;
    color: #667085;
    font-size: 0.95rem;
    margin-bottom: 1rem;
    line-height: 1.5;
}

.dashboard-meta .meta-item {
    white-space: nowrap;
}

.dashboard-meta .meta-separator {
    color: #c4c9d2;
}

.section-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #202939;
    margin-top: 0.4rem;
    margin-bottom: 0.5rem;
}

div[data-testid="stMetric"] {
    display: none;
}

.kpi-card {
    background: #ffffff;
    border: 1px solid #e6e9ef;
    padding: 0.9rem 1rem;
    border-radius: 12px;
    min-height: 105px;
    margin: 0.25rem 0.35rem;
    box-shadow: 0 2px 6px rgba(16, 24, 40, 0.06);
}

.kpi-title {
    color: #667085;
    font-size: 0.82rem;
    line-height: 1.25;
    margin-bottom: 0.45rem;
}

.kpi-value {
    color: #172033;
    font-size: 1.15rem;
    font-weight: 750;
    line-height: 1.3;
    margin-bottom: 0.45rem;
}

.kpi-direction {
    color: #475467;
    font-size: 0.82rem;
}

/* Tab等宽铺满 */
.stTabs [data-baseweb="tab-list"] {
    display: grid;
    grid-template-columns: repeat(6, minmax(0, 1fr));
    width: 100%;
    gap: 0.55rem;
    background: transparent;
}

.stTabs [data-baseweb="tab"] {
    flex: 1 1 0;
    justify-content: center;
    min-height: 2.8rem;
    padding: 0.65rem 0.7rem;
    background: #ffffff;
    border: 1px solid #e6e9ef;
    border-radius: 9px 9px 0 0;
    white-space: nowrap;
}

.stTabs [aria-selected="true"] {
    background: #eef4ff !important;
    color: #1849a9 !important;
    border-color: #b8cdf6 !important;
    font-weight: 700;
}

.stDownloadButton button {
    border-radius: 8px;
}

.risk-box {
    border: 1px solid #e6e9ef;
    border-radius: 12px;
    padding: 0.9rem 1rem;
    background: #fafbfc;
    margin: 0.5rem 0 1rem 0;
    color: #344054;
}

@media (max-width: 900px) {
    .stTabs [data-baseweb="tab-list"] {
        overflow-x: auto;
    }

    .stTabs [data-baseweb="tab"] {
        flex: 0 0 auto;
        min-width: 130px;
    }

    .dashboard-meta .meta-separator {
        display: none;
    }
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# 4. 基础函数
# =========================================================
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def parse_month_from_path(path: Path) -> Optional[str]:
    match = re.search(
        r"trading_review_standardized_(\d{4}-\d{2})\.xlsx$",
        path.name,
    )
    return match.group(1) if match else None


def find_month_files(base_dir: Path) -> Dict[str, Path]:
    result: Dict[str, Path] = {}

    if not base_dir.exists():
        return result

    for path in base_dir.rglob(OUTPUT_FILE_PATTERN):
        month = parse_month_from_path(path)
        if month:
            result[month] = path

    return dict(sorted(result.items()))


def normalize_time_fields(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    if "trade_date" in out.columns:
        out["trade_date"] = pd.to_datetime(
            out["trade_date"], errors="coerce"
        ).dt.normalize()

    if "datetime_bj" in out.columns:
        out["datetime_bj"] = pd.to_datetime(
            out["datetime_bj"], errors="coerce"
        )

    if "slot_index" in out.columns:
        out["slot_index"] = safe_numeric(out["slot_index"]).astype("Int64")

    return out


@st.cache_data(show_spinner=False)
def load_month_data(
    file_path: str,
    modified_time: float,
) -> Dict[str, pd.DataFrame]:
    """
    modified_time作为缓存Key的一部分。
    Excel更新后会自动重新读取。
    """
    _ = modified_time
    xls = pd.ExcelFile(file_path)

    output: Dict[str, pd.DataFrame] = {}
    sheets = [
        "long_term",
        "energy_block",
        "day_ahead",
        "real_time",
        "summary_15min",
        "check",
    ]

    for sheet in sheets:
        if sheet in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            output[sheet] = normalize_time_fields(clean_columns(df))
        else:
            output[sheet] = pd.DataFrame()

    return output


def derive_trade_fields(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    if out.empty or "datetime_bj" not in out.columns:
        return out

    out["datetime_bj"] = pd.to_datetime(
        out["datetime_bj"], errors="coerce"
    )
    out = out.dropna(subset=["datetime_bj"]).copy()

    if "slot_index" not in out.columns:
        out["slot_index"] = (
            out["datetime_bj"].dt.hour * 4
            + out["datetime_bj"].dt.minute // 15
        )

        zero_mask = (
            (out["datetime_bj"].dt.hour == 0)
            & (out["datetime_bj"].dt.minute == 0)
        )
        out.loc[zero_mask, "slot_index"] = 96

    out["slot_index"] = safe_numeric(out["slot_index"]).astype("Int64")

    if "time_label" not in out.columns:
        out["time_label"] = out["datetime_bj"].dt.strftime("%H:%M")
        out.loc[out["slot_index"] == 96, "time_label"] = "24:00"

    # 展示日期统一按北京时间对应的用电日识别。
    # 96点通常以次日00:00表示，因此仍归入前一自然日。
    out["display_date"] = out["datetime_bj"].dt.normalize()
    slot_96_mask = out["slot_index"].eq(96)
    out.loc[slot_96_mask, "display_date"] = (
        out.loc[slot_96_mask, "datetime_bj"]
        - pd.Timedelta(days=1)
    ).dt.normalize()

    # trade_date保留底层原始交易日；若缺失则用display_date补充。
    if "trade_date" not in out.columns:
        out["trade_date"] = out["display_date"]
    else:
        out["trade_date"] = pd.to_datetime(
            out["trade_date"], errors="coerce"
        ).dt.normalize()
        out["trade_date"] = out["trade_date"].fillna(out["display_date"])

    return out


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    buffer.seek(0)
    return buffer.getvalue()


def get_price_columns(price_mode: str):
    if price_mode == "电能量价格":
        return "energy_price_yuan_mwh", "energy_amount_yuan"
    return "price_yuan_mwh", "amount_yuan"


def get_options(
    df: pd.DataFrame,
    col: str,
    default: list[str],
) -> list[str]:
    if df.empty or col not in df.columns:
        return default

    values = [str(x) for x in df[col].dropna().unique()]
    ordered = [x for x in default if x in values]
    ordered += sorted(x for x in values if x not in ordered)
    return ordered


# =========================================================
# 江苏峰平谷规则
# =========================================================
def period_type(dt):
    """
    江苏现货市场峰平谷规则。
    返回：谷、平、峰。
    """
    if pd.isna(dt):
        return "未知"

    dt = pd.Timestamp(dt)
    h = dt.hour
    m = dt.month

    if dt < pd.Timestamp("2025-06-01"):
        if 0 <= h < 8:
            return "谷"
        if (8 <= h < 11) or (17 <= h < 22):
            return "峰"
        return "平"

    if m in [6, 7, 8, 12, 1, 2]:
        if (0 <= h < 6) or (11 <= h < 13):
            return "谷"
        if 14 <= h < 22:
            return "峰"
        return "平"

    if m in [3, 4, 5, 9, 10, 11]:
        if (2 <= h < 6) or (10 <= h < 14):
            return "谷"
        if 15 <= h < 22:
            return "峰"
        return "平"

    return "平"


def calculate_period_weighted_prices(
    df: pd.DataFrame,
    amount_col: str,
) -> dict[str, float]:
    """
    按当前所选价格口径计算峰、平、谷成交加权均价。

    计算口径与现有综合加权均价保持一致：
    各时段金额合计 ÷ 各时段净交易电量合计。
    """
    result = {"峰": np.nan, "平": np.nan, "谷": np.nan}

    if df.empty:
        return result

    out = df.copy()

    if "datetime_bj" not in out.columns:
        return result

    out["datetime_bj"] = pd.to_datetime(
        out["datetime_bj"],
        errors="coerce",
    )
    out["energy_mwh"] = safe_numeric(out["energy_mwh"]).fillna(0)

    if amount_col not in out.columns:
        out[amount_col] = 0
    out[amount_col] = safe_numeric(out[amount_col]).fillna(0)

    out = out.dropna(subset=["datetime_bj"]).copy()
    out["峰平谷"] = out["datetime_bj"].apply(period_type)

    grouped = (
        out.groupby("峰平谷", as_index=False)
        .agg(
            净交易电量_MWh=("energy_mwh", "sum"),
            成交金额_元=(amount_col, "sum"),
        )
    )

    for _, row in grouped.iterrows():
        period_name = str(row["峰平谷"])
        energy = float(row["净交易电量_MWh"])
        amount = float(row["成交金额_元"])

        if period_name in result and abs(energy) > 1e-9:
            result[period_name] = amount / energy

    return result


# =========================================================
# 5. 数据汇总函数
# =========================================================
def aggregate_to_hour(
    df: pd.DataFrame,
    amount_col: str,
) -> pd.DataFrame:
    out = df.copy()
    out["hour_no"] = ((out["slot_index"].astype(int) - 1) // 4) + 1
    out["hour_label"] = out["hour_no"].apply(
        lambda x: f"{x - 1:02d}:00-{x:02d}:00"
    )

    grouped = (
        out.groupby(
            [
                "display_date",
                "hour_no",
                "hour_label",
                "market_stage",
                "market_stage_cn",
            ],
            as_index=False,
        )
        .agg(
            energy_mwh=("energy_mwh", "sum"),
            amount_selected=(amount_col, "sum"),
        )
    )

    grouped["price_selected"] = np.where(
        grouped["energy_mwh"].abs() > 1e-9,
        grouped["amount_selected"] / grouped["energy_mwh"],
        np.nan,
    )
    return grouped


def aggregate_daily(
    df: pd.DataFrame,
    amount_col: str,
) -> pd.DataFrame:
    out = df.copy()
    out["display_date"] = pd.to_datetime(
        out["display_date"], errors="coerce"
    ).dt.normalize()
    out["market_stage_cn"] = out["market_stage_cn"].astype(str)
    out = out.dropna(subset=["display_date", "market_stage_cn"])

    grouped = (
        out.groupby(
            ["display_date", "market_stage_cn"],
            as_index=False,
        )
        .agg(
            energy_mwh=("energy_mwh", "sum"),
            amount_selected=(amount_col, "sum"),
        )
    )

    grouped["avg_price"] = np.where(
        grouped["energy_mwh"].abs() > 1e-9,
        grouped["amount_selected"] / grouped["energy_mwh"],
        np.nan,
    )
    return grouped


def market_summary_table(
    df: pd.DataFrame,
    amount_col: str,
) -> pd.DataFrame:
    grouped = (
        df.groupby("market_stage_cn", as_index=False)
        .agg(
            净交易电量_MWh=("energy_mwh", "sum"),
            买入量_MWh=(
                "energy_mwh",
                lambda x: x[x > 0].sum(),
            ),
            卖出量_MWh=(
                "energy_mwh",
                lambda x: -x[x < 0].sum(),
            ),
            金额_元=(amount_col, "sum"),
        )
    )

    grouped["加权均价_元每MWh"] = np.where(
        grouped["净交易电量_MWh"].abs() > 1e-9,
        grouped["金额_元"] / grouped["净交易电量_MWh"],
        np.nan,
    )

    grouped["交易方向"] = np.select(
        [
            grouped["净交易电量_MWh"] > 1e-9,
            grouped["净交易电量_MWh"] < -1e-9,
        ],
        ["净买入", "净卖出"],
        default="持平",
    )

    grouped["市场阶段"] = pd.Categorical(
        grouped["market_stage_cn"],
        categories=MARKET_ORDER,
        ordered=True,
    )

    return (
        grouped.sort_values("市场阶段")[
            [
                "市场阶段",
                "交易方向",
                "净交易电量_MWh",
                "买入量_MWh",
                "卖出量_MWh",
                "加权均价_元每MWh",
                "金额_元",
            ]
        ]
        .reset_index(drop=True)
    )


def prepare_summary(
    summary_df: pd.DataFrame,
    long_term_df: pd.DataFrame,
    selected_terms: list[str],
    selected_green: list[str],
    selected_regions: list[str],
) -> pd.DataFrame:
    summary = derive_trade_fields(summary_df)
    if summary.empty:
        return summary

    for col in [
        "energy_mwh",
        "price_yuan_mwh",
        "energy_price_yuan_mwh",
        "amount_yuan",
        "energy_amount_yuan",
    ]:
        if col not in summary.columns:
            summary[col] = np.nan
        summary[col] = safe_numeric(summary[col])

    non_lt = summary[
        summary["market_stage"] != "long_term"
    ].copy()

    lt = derive_trade_fields(long_term_df)

    if not lt.empty:
        for col in ["energy_mwh", "total_amount", "energy_amount"]:
            if col not in lt.columns:
                lt[col] = 0
            lt[col] = safe_numeric(lt[col]).fillna(0)

        if selected_terms and "contract_term" in lt.columns:
            lt = lt[
                lt["contract_term"].astype(str).isin(selected_terms)
            ]

        if selected_green and "green_type" in lt.columns:
            lt = lt[
                lt["green_type"].astype(str).isin(selected_green)
            ]

        if selected_regions and "region_type" in lt.columns:
            lt = lt[
                lt["region_type"].astype(str).isin(selected_regions)
            ]

        lt_group = (
            lt.groupby(
                [
                    "trade_date",
                    "display_date",
                    "datetime_bj",
                    "slot_index",
                    "time_label",
                ],
                as_index=False,
            )
            .agg(
                energy_mwh=("energy_mwh", "sum"),
                amount_yuan=("total_amount", "sum"),
                energy_amount_yuan=("energy_amount", "sum"),
            )
        )

        lt_group["price_yuan_mwh"] = np.where(
            lt_group["energy_mwh"].abs() > 1e-9,
            lt_group["amount_yuan"] / lt_group["energy_mwh"],
            np.nan,
        )

        lt_group["energy_price_yuan_mwh"] = np.where(
            lt_group["energy_mwh"].abs() > 1e-9,
            lt_group["energy_amount_yuan"] / lt_group["energy_mwh"],
            np.nan,
        )

        lt_group["market_stage"] = "long_term"

        final = pd.concat(
            [lt_group, non_lt],
            ignore_index=True,
        )
    else:
        final = non_lt

    final["market_stage_cn"] = (
        final["market_stage"]
        .map(MARKET_CODE_TO_CN)
        .fillna(final["market_stage"])
    )

    return final.sort_values(
        ["display_date", "slot_index", "market_stage_cn"]
    ).reset_index(drop=True)


def stage_net_energy(
    df: pd.DataFrame,
    stage: str,
) -> float:
    return float(
        df.loc[df["market_stage"] == stage, "energy_mwh"].sum()
    )


def stage_amount(
    df: pd.DataFrame,
    stage: str,
    amount_col: str,
) -> float:
    return float(
        df.loc[df["market_stage"] == stage, amount_col].sum()
    )


def format_energy(value: float) -> str:
    direction = "买入" if value > 1e-9 else "卖出" if value < -1e-9 else "持平"
    return f"{value:,.0f} MWh（{direction}）"


def calculate_spot_energy_weighted_price(
    df: pd.DataFrame,
) -> tuple[float, float, float]:
    """
    日前+实时净结算加权均价。

    固定规则：
    - 仅使用日前、实时市场；
    - 仅使用电能量价格 energy_price_yuan_mwh；
    - 不读取 amount_col；
    - 不受“综合成交价格（含环境权益价值）”切换影响。

    返回：
    净结算均价, 净交易电量, 净结算金额
    """
    spot = df.loc[
        df["market_stage"].isin(
            ["day_ahead", "real_time"]
        ),
        [
            "market_stage",
            "energy_mwh",
            "energy_price_yuan_mwh",
        ],
    ].copy()

    if spot.empty:
        return np.nan, 0.0, 0.0

    spot["energy_mwh"] = safe_numeric(
        spot["energy_mwh"]
    ).fillna(0)

    spot["energy_price_yuan_mwh"] = safe_numeric(
        spot["energy_price_yuan_mwh"]
    )

    spot = spot[
        spot["energy_price_yuan_mwh"].notna()
    ].copy()

    if spot.empty:
        return np.nan, 0.0, 0.0

    spot["signed_amount_yuan"] = (
        spot["energy_mwh"]
        *
        spot["energy_price_yuan_mwh"]
    )

    net_energy = float(
        spot["energy_mwh"].sum()
    )

    net_amount = float(
        spot["signed_amount_yuan"].sum()
    )

    if abs(net_energy) <= 1e-9:
        return np.nan, net_energy, net_amount

    return (
        net_amount / net_energy,
        net_energy,
        net_amount,
    )


def calculate_recovery_fee(
    coverage_ratio: float,
    actual_energy: float,
    long_term_energy: float,
    spot_price: float,
    market_lt_price: float,
    q1: float,
    q2: float,
    m: float,
) -> tuple[float, float, str]:
    """
    返回：
    recovery_fee：按C<0不回收后的费用
    deviation_energy：不足或超额电量
    risk_type：低签/超签/正常
    """
    if (
        actual_energy <= 0
        or pd.isna(coverage_ratio)
        or pd.isna(spot_price)
        or pd.isna(market_lt_price)
    ):
        return np.nan, np.nan, "无法计算"

    if coverage_ratio < q1:
        deviation_energy = q1 * actual_energy - long_term_energy
        raw_fee = (
            deviation_energy
            * (market_lt_price - spot_price)
            * m
        )
        return max(raw_fee, 0.0), deviation_energy, "低签"

    if coverage_ratio > q2:
        deviation_energy = long_term_energy - q2 * actual_energy
        raw_fee = (
            deviation_energy
            * (spot_price - market_lt_price)
            * m
        )
        return max(raw_fee, 0.0), deviation_energy, "超签"

    return 0.0, 0.0, "正常"


# =========================================================
# 6. 文件读取
# =========================================================
month_files = find_month_files(BASE_DIR)

if not month_files:
    st.error("未找到标准化交易复盘文件，请检查目录。")
    st.stop()

available_months = list(month_files.keys())


# =========================================================
# 7. 侧边栏
# =========================================================
with st.sidebar:
    st.markdown("### 数据筛选")

    selected_month = st.selectbox(
        "月份",
        options=available_months,
        index=len(available_months) - 1,
        format_func=lambda x: f"{x[:4]}年{int(x[5:])}月",
    )

    selected_file = month_files[selected_month]
    modified_time = selected_file.stat().st_mtime

    data = load_month_data(
        str(selected_file),
        modified_time,
    )

    long_term_raw = data["long_term"]
    summary_raw = data["summary_15min"]
    energy_block_raw = data["energy_block"]

    summary_dates = derive_trade_fields(summary_raw)

    if summary_dates.empty:
        st.error("summary_15min为空。")
        st.stop()

    # 仅展示所选月份对应的北京时间用电日
    month_start = pd.Timestamp(f"{selected_month}-01")
    month_end = month_start + pd.offsets.MonthEnd(0)

    available_display_dates = summary_dates[
        summary_dates["display_date"].between(
            month_start,
            month_end,
        )
    ]["display_date"].dropna()

    if available_display_dates.empty:
        st.error("所选月份没有可用的北京时间交易数据。")
        st.stop()

    min_date = available_display_dates.min().date()
    max_date = available_display_dates.max().date()

    selected_date_range = st.date_input(
        "日期范围",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

    if (
        isinstance(selected_date_range, tuple)
        and len(selected_date_range) == 2
    ):
        start_date, end_date = selected_date_range
    else:
        start_date = end_date = selected_date_range

    resolution = st.radio(
        "时间颗粒度",
        ["96点（15分钟）", "24小时"],
        horizontal=True,
    )

    price_mode = st.radio(
        "价格口径",
        ["电能量价格", "综合成交价格（含环境权益价值）"],
        horizontal=True,
    )

    st.markdown("---")
    st.markdown("### 年协/月协筛选")

    term_options = get_options(
        long_term_raw,
        "contract_term",
        ["年度", "月度"],
    )
    green_options = get_options(
        long_term_raw,
        "green_type",
        ["绿电", "非绿电"],
    )
    region_options = get_options(
        long_term_raw,
        "region_type",
        ["省内", "省间"],
    )

    selected_terms = st.multiselect(
        "合同周期",
        term_options,
        default=term_options,
    )

    selected_green = st.multiselect(
        "电力属性",
        green_options,
        default=green_options,
    )

    selected_regions = st.multiselect(
        "交易范围",
        region_options,
        default=region_options,
    )

    st.caption(f"数据文件：{selected_file.name}")


# =========================================================
# 8. 数据准备
# =========================================================
price_col, amount_col = get_price_columns(price_mode)

summary = prepare_summary(
    summary_raw,
    long_term_raw,
    selected_terms,
    selected_green,
    selected_regions,
)

start_ts = pd.Timestamp(start_date)
end_ts = pd.Timestamp(end_date)

filtered = summary[
    summary["display_date"].between(start_ts, end_ts)
].copy()

if filtered.empty:
    st.warning("当前条件没有数据。")
    st.stop()

# 先完成统一数值清洗。
# 避免现货计算数据源与页面其他计算使用不同数据路径。
for col in [
    "energy_mwh",
    "energy_price_yuan_mwh",
    "energy_amount_yuan",
    "price_yuan_mwh",
    "amount_yuan",
]:
    if col not in filtered.columns:
        filtered[col] = 0
    filtered[col] = safe_numeric(filtered[col]).fillna(0)


# 日前、实时市场没有绿电和环境权益。
# 在数值清洗完成后生成现货计算数据源。
spot_energy_source = filtered.loc[
    filtered["market_stage"].isin(
        ["day_ahead", "real_time"]
    ),
    [
        "market_stage",
        "energy_mwh",
        "energy_price_yuan_mwh",
    ],
].copy()

stage_summary = market_summary_table(filtered, amount_col)
period_weighted_prices = calculate_period_weighted_prices(
    filtered,
    amount_col,
)

total_net_energy = float(filtered["energy_mwh"].sum())
total_amount = float(filtered[amount_col].sum())

overall_price = (
    total_amount / total_net_energy
    if abs(total_net_energy) > 1e-9
    else np.nan
)

lt_net = stage_net_energy(filtered, "long_term")
eb_net = stage_net_energy(filtered, "energy_block")
da_net = stage_net_energy(filtered, "day_ahead")
rt_net = stage_net_energy(filtered, "real_time")

long_term_total_net = lt_net + eb_net
spot_net = da_net + rt_net

# 日前+实时市场不存在环境权益价值。
# 均价始终固定使用电能量价格口径，
# 与页面选择的综合成交价格口径相互独立。
(
    spot_weighted_price,
    spot_price_net_energy,
    spot_signed_energy_amount,
) = calculate_spot_energy_weighted_price(
    spot_energy_source
)


# =========================================================
# 9. 页面标题
# =========================================================
st.markdown(
    '<div class="dashboard-title">江苏电力交易量价复盘平台</div>',
    unsafe_allow_html=True,
)

st.markdown(
    f"""
<div class="dashboard-meta">
    <span class="meta-item">数据周期：{selected_month[:4]}年{int(selected_month[5:])}月</span>
    <span class="meta-separator">｜</span>
    <span class="meta-item">统计范围：{start_date:%Y-%m-%d} 至 {end_date:%Y-%m-%d}</span>
    <span class="meta-separator">｜</span>
    <span class="meta-item">价格口径：{price_mode}</span>
</div>
""",
    unsafe_allow_html=True,
)


# =========================================================
# 10. 顶部KPI
# =========================================================
def direction_text(value: float) -> str:
    if value > 1e-9:
        return "↑ 净买入"
    if value < -1e-9:
        return "↓ 净卖出"
    return "— 持平"


def render_kpi(title: str, value: str, direction: str = ""):
    st.markdown(
        f"""
<div class="kpi-card">
    <div class="kpi-title">{title}</div>
    <div class="kpi-value">{value}</div>
    <div class="kpi-direction">{direction}</div>
</div>
""",
        unsafe_allow_html=True,
    )


k1, k2, k3, k4 = st.columns(4, gap="medium")

with k1:
    render_kpi(
        "最终净交易电量",
        f"{total_net_energy:,.0f} MWh",
        direction_text(total_net_energy),
    )

with k2:
    render_kpi(
        "综合加权均价",
        "—"
        if pd.isna(overall_price)
        else f"{overall_price:,.2f} 元/MWh",
    )

with k3:
    render_kpi(
        "年协/月协净电量",
        f"{lt_net:,.0f} MWh",
        direction_text(lt_net),
    )

with k4:
    render_kpi(
        "能量块净调整量",
        f"{eb_net:,.0f} MWh",
        direction_text(eb_net),
    )

k5, k6, k7 = st.columns(3, gap="medium")

with k5:
    render_kpi(
        "现货净调整量",
        f"{spot_net:,.0f} MWh",
        direction_text(spot_net),
    )

with k6:
    render_kpi(
        "现货（日+实）净结算加权均价",
        "—"
        if pd.isna(spot_weighted_price)
        else f"{spot_weighted_price:,.2f} 元/MWh",
    )

with k7:
    if spot_signed_energy_amount < -1e-6:
        render_kpi(
            "日前/实时套利结果",
            f"{abs(spot_signed_energy_amount):,.0f} 元",
            "套利收益",
        )
    elif spot_signed_energy_amount > 1e-6:
        render_kpi(
            "日前/实时套利结果",
            f"{abs(spot_signed_energy_amount):,.0f} 元",
            "套利损失",
        )
    else:
        render_kpi(
            "日前/实时套利结果",
            "0 元",
            "— 持平",
        )


# =========================================================
# 11. Tabs
# =========================================================
(
    tab_overview,
    tab_risk,
    tab_intraday,
    tab_trend,
    tab_eb,
    tab_data,
) = st.tabs(
    [
        "交易概览",
        "中长期偏差预警",
        "日内量价",
        "每日趋势",
        "能量块复盘",
        "明细与下载",
    ]
)


# =========================================================
# Tab 1 交易概览
# =========================================================
with tab_overview:
    st.markdown(
        '<div class="section-title">交易结构与量价概览</div>',
        unsafe_allow_html=True,
    )

    peak_col, flat_col, valley_col = st.columns(3, gap="medium")

    with peak_col:
        peak_price = period_weighted_prices.get("峰", np.nan)
        render_kpi(
            "峰段成交加权均价",
            "—"
            if pd.isna(peak_price)
            else f"{peak_price:,.2f} 元/MWh",
            price_mode,
        )

    with flat_col:
        flat_price = period_weighted_prices.get("平", np.nan)
        render_kpi(
            "平段成交加权均价",
            "—"
            if pd.isna(flat_price)
            else f"{flat_price:,.2f} 元/MWh",
            price_mode,
        )

    with valley_col:
        valley_price = period_weighted_prices.get("谷", np.nan)
        render_kpi(
            "谷段成交加权均价",
            "—"
            if pd.isna(valley_price)
            else f"{valley_price:,.2f} 元/MWh",
            price_mode,
        )

    st.markdown("<div style='height:0.35rem'></div>", unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1.15])

    with col_left:
        plot_stage = stage_summary.copy()
        plot_stage["市场阶段文本"] = plot_stage["市场阶段"].astype(str)

        fig_stage = go.Figure()
        fig_stage.add_trace(
            go.Bar(
                x=plot_stage["市场阶段文本"],
                y=plot_stage["净交易电量_MWh"],
                customdata=np.column_stack(
                    [
                        plot_stage["交易方向"],
                        plot_stage["加权均价_元每MWh"],
                        plot_stage["买入量_MWh"],
                        plot_stage["卖出量_MWh"],
                    ]
                ),
                hovertemplate=(
                    "市场阶段：%{x}<br>"
                    "交易方向：%{customdata[0]}<br>"
                    "净交易量：%{y:,.2f} MWh<br>"
                    "买入量：%{customdata[2]:,.2f} MWh<br>"
                    "卖出量：%{customdata[3]:,.2f} MWh<br>"
                    "加权均价：%{customdata[1]:,.2f} 元/MWh"
                    "<extra></extra>"
                ),
            )
        )

        fig_stage.add_hline(y=0, line_width=1)

        fig_stage.update_layout(
            title="各阶段净交易电量",
            xaxis_title="",
            yaxis_title="MWh",
            height=420,
            margin=dict(l=10, r=10, t=55, b=20),
            showlegend=False,
        )

        st.plotly_chart(
            fig_stage,
            use_container_width=True,
        )

    with col_right:
        waterfall_x = [
            "年协/月协",
            "能量块调整",
            "日前调整",
            "实时调整",
            "最终净交易量",
        ]
        waterfall_y = [
            lt_net,
            eb_net,
            da_net,
            rt_net,
            total_net_energy,
        ]
        waterfall_measure = [
            "absolute",
            "relative",
            "relative",
            "relative",
            "total",
        ]

        fig_waterfall = go.Figure(
            go.Waterfall(
                x=waterfall_x,
                y=waterfall_y,
                measure=waterfall_measure,
                text=[
                    f"{lt_net:,.0f}",
                    f"{eb_net:+,.0f}",
                    f"{da_net:+,.0f}",
                    f"{rt_net:+,.0f}",
                    f"{total_net_energy:,.0f}",
                ],
                textposition="outside",
                connector={"line": {"width": 1}},
                hovertemplate=(
                    "阶段：%{x}<br>"
                    "电量：%{y:,.2f} MWh"
                    "<extra></extra>"
                ),
            )
        )

        fig_waterfall.update_layout(
            title="交易电量形成过程",
            xaxis_title="",
            yaxis_title="MWh",
            height=420,
            margin=dict(l=10, r=10, t=55, b=20),
            showlegend=False,
        )

        st.plotly_chart(
            fig_waterfall,
            use_container_width=True,
        )

    st.markdown("#### 现货净结算均价拆解")

    spot_audit_rows = []
    for stage_code, stage_name in [
        ("day_ahead", "日前"),
        ("real_time", "实时"),
    ]:
        stage_df = filtered[
            filtered["market_stage"] == stage_code
        ].copy()

        stage_df["energy_mwh"] = safe_numeric(
            stage_df["energy_mwh"]
        ).fillna(0)
        stage_df["energy_price_yuan_mwh"] = safe_numeric(
            stage_df["energy_price_yuan_mwh"]
        )

        stage_df = stage_df[
            stage_df["energy_price_yuan_mwh"].notna()
        ].copy()

        signed_energy = float(stage_df["energy_mwh"].sum())
        signed_amount = float(
            (
                stage_df["energy_mwh"]
                * stage_df["energy_price_yuan_mwh"]
            ).sum()
        )

        stage_price = (
            signed_amount / signed_energy
            if abs(signed_energy) > 1e-9
            else np.nan
        )

        spot_audit_rows.append(
            {
                "市场阶段": stage_name,
                "净电量(MWh)": signed_energy,
                "带符号电能量金额(元)": signed_amount,
                "净结算加权均价(元/MWh)": stage_price,
            }
        )

    spot_audit_rows.append(
        {
            "市场阶段": "日前+实时合计",
            "净电量(MWh)": spot_price_net_energy,
            "带符号电能量金额(元)": spot_signed_energy_amount,
            "净结算加权均价(元/MWh)": spot_weighted_price,
        }
    )

    spot_audit_df = pd.DataFrame(spot_audit_rows)

    st.dataframe(
        spot_audit_df.round(2),
        use_container_width=True,
        hide_index=True,
    )

    # 日前/实时市场交易效果解释
    if spot_signed_energy_amount < -1e-6:
        st.success(
            f"🟢 日前/实时市场套利收益："
            f"{abs(spot_signed_energy_amount):,.2f} 元"
        )
    elif spot_signed_energy_amount > 1e-6:
        st.error(
            f"🔴 日前/实时市场套利损失："
            f"{abs(spot_signed_energy_amount):,.2f} 元"
        )
    else:
        st.info(
            "日前/实时市场交易基本持平。"
        )

    st.caption(
        "净结算口径中，买入电量为正、卖出电量为负。"
        "交易收益/损失根据日前与实时市场组合后的净结算金额判断。"
    )

    st.markdown("#### 阶段量价摘要")

    display_summary = stage_summary.copy()
    display_summary.columns = [
        "市场阶段",
        "交易方向",
        "净交易量(MWh)",
        "买入量(MWh)",
        "卖出量(MWh)",
        "加权均价(元/MWh)",
        "交易金额(元)",
    ]

    st.dataframe(
        display_summary.round(2),
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# Tab 2 日内量价
# =========================================================
with tab_intraday:
    st.markdown(
        '<div class="section-title">日内交易量价分布</div>',
        unsafe_allow_html=True,
    )

    available_dates = sorted(
        filtered["display_date"].dropna().dt.date.unique()
    )

    selected_day = st.selectbox(
        "选择用电日",
        available_dates,
        index=len(available_dates) - 1,
        format_func=lambda x: x.strftime("%Y-%m-%d"),
        key="intraday_day",
    )

    day_df = filtered[
        filtered["display_date"] == pd.Timestamp(selected_day)
    ].copy()

    if resolution == "24小时":
        plot_df = aggregate_to_hour(day_df, amount_col)
        x_col = "hour_label"
        price_plot_col = "price_selected"
    else:
        plot_df = day_df.copy()
        plot_df["x_label"] = plot_df["time_label"].astype(str)
        x_col = "x_label"
        price_plot_col = price_col

    market_selection = st.multiselect(
        "显示市场阶段",
        MARKET_ORDER,
        default=MARKET_ORDER,
        key="intraday_market",
    )

    plot_df = plot_df[
        plot_df["market_stage_cn"].astype(str).isin(market_selection)
    ]

    fig_qty = px.bar(
        plot_df,
        x=x_col,
        y="energy_mwh",
        color="market_stage_cn",
        barmode="relative",
        category_orders={"market_stage_cn": MARKET_ORDER},
    )

    fig_qty.update_traces(
        hovertemplate=(
            "时段：%{x}<br>"
            "市场阶段：%{fullData.name}<br>"
            "交易电量：%{y:,.2f} MWh"
            "<extra></extra>"
        )
    )

    fig_qty.update_layout(
        title=f"{selected_day} 各阶段交易电量",
        yaxis_title="MWh",
        height=500,
        hovermode="x unified",
        xaxis_title=None,
    )

    st.plotly_chart(fig_qty, use_container_width=True)

    # 价格图：日前/实时保持连续曲线；能量块仅显示真实成交点，避免无交易时段显示0价格
    import plotly.graph_objects as go

    fig_price = go.Figure()
    for market in market_selection:
        tmp = plot_df[plot_df["market_stage_cn"] == market].copy()
        if tmp.empty:
            continue
        if market == "能量块":
            tmp = tmp[tmp["energy_mwh"].abs() > 0].copy()
            fig_price.add_trace(go.Scatter(
                x=tmp[x_col],
                y=tmp[price_plot_col],
                mode="markers",
                name=market,
            ))
        else:
            fig_price.add_trace(go.Scatter(
                x=tmp[x_col],
                y=tmp[price_plot_col],
                mode="lines+markers",
                name=market,
            ))

    fig_price.update_traces(
        hovertemplate=(
            "时段：%{x}<br>"
            "市场阶段：%{fullData.name}<br>"
            "价格：%{y:,.2f} 元/MWh"
            "<extra></extra>"
        )
    )

    fig_price.update_layout(
        title=f"{selected_day} 各阶段交易价格",
        yaxis_title="元/MWh",
        height=430,
        hovermode="x unified",
        xaxis_title=None,
    )

    st.plotly_chart(fig_price, use_container_width=True)


# =========================================================
# Tab 3 每日趋势
# =========================================================
with tab_trend:
    st.markdown(
        '<div class="section-title">每日交易量与价格趋势</div>',
        unsafe_allow_html=True,
    )

    daily = aggregate_daily(filtered, amount_col)
    col1, col2 = st.columns(2)

    with col1:
        fig_daily_qty = px.bar(
            daily,
            x="display_date",
            y="energy_mwh",
            color="market_stage_cn",
            barmode="relative",
            category_orders={"market_stage_cn": MARKET_ORDER},
        )

        fig_daily_qty.update_traces(
            hovertemplate=(
                "日期：%{x|%Y-%m-%d}<br>"
                "市场阶段：%{fullData.name}<br>"
                "交易电量：%{y:,.2f} MWh"
                "<extra></extra>"
            )
        )

        fig_daily_qty.update_xaxes(tickformat="%m-%d")
        fig_daily_qty.update_layout(
            title="每日各阶段净交易电量",
            xaxis_title="日期",
            yaxis_title="MWh",
            height=500,
            hovermode="x unified",
            legend_title_text="",
        )

        st.plotly_chart(
            fig_daily_qty,
            use_container_width=True,
        )

    with col2:
        fig_daily_price = px.line(
            daily,
            x="display_date",
            y="avg_price",
            color="market_stage_cn",
            markers=True,
            category_orders={"market_stage_cn": MARKET_ORDER},
        )

        fig_daily_price.update_traces(
            hovertemplate=(
                "日期：%{x|%Y-%m-%d}<br>"
                "市场阶段：%{fullData.name}<br>"
                "加权均价：%{y:,.2f} 元/MWh"
                "<extra></extra>"
            )
        )

        fig_daily_price.update_xaxes(tickformat="%m-%d")
        fig_daily_price.update_layout(
            title="每日各阶段加权均价",
            xaxis_title="日期",
            yaxis_title="元/MWh",
            height=500,
            hovermode="x unified",
            legend_title_text="",
        )

        st.plotly_chart(
            fig_daily_price,
            use_container_width=True,
        )

    daily_structure = (
        daily.pivot_table(
            index="display_date",
            columns="market_stage_cn",
            values="energy_mwh",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
    )

    for market in MARKET_ORDER:
        if market not in daily_structure.columns:
            daily_structure[market] = 0

    daily_structure["中长期净电量"] = (
        daily_structure["年协/月协"]
        + daily_structure["能量块"]
    )
    daily_structure["现货净调整量"] = (
        daily_structure["日前"]
        + daily_structure["实时"]
    )
    daily_structure["最终净交易量"] = (
        daily_structure["中长期净电量"]
        + daily_structure["现货净调整量"]
    )

    st.markdown("#### 每日交易结构变化")

    fig_structure = go.Figure()
    fig_structure.add_trace(
        go.Scatter(
            x=daily_structure["display_date"],
            y=daily_structure["中长期净电量"],
            name="中长期净电量",
            mode="lines+markers",
            hovertemplate=(
                "日期：%{x|%Y-%m-%d}<br>"
                "中长期净电量：%{y:,.2f} MWh"
                "<extra></extra>"
            ),
        )
    )
    fig_structure.add_trace(
        go.Scatter(
            x=daily_structure["display_date"],
            y=daily_structure["现货净调整量"],
            name="现货净调整量",
            mode="lines+markers",
            hovertemplate=(
                "日期：%{x|%Y-%m-%d}<br>"
                "现货净调整量：%{y:,.2f} MWh"
                "<extra></extra>"
            ),
        )
    )
    fig_structure.add_trace(
        go.Scatter(
            x=daily_structure["display_date"],
            y=daily_structure["最终净交易量"],
            name="最终净交易量",
            mode="lines+markers",
            hovertemplate=(
                "日期：%{x|%Y-%m-%d}<br>"
                "最终净交易量：%{y:,.2f} MWh"
                "<extra></extra>"
            ),
        )
    )

    fig_structure.add_hline(y=0, line_width=1)
    fig_structure.update_xaxes(tickformat="%m-%d")
    fig_structure.update_layout(
        title="每日中长期、现货调整与最终净量",
        xaxis_title="日期",
        yaxis_title="MWh",
        height=410,
        hovermode="x unified",
    )

    st.plotly_chart(
        fig_structure,
        use_container_width=True,
    )


# =========================================================
# Tab 4 能量块复盘
# =========================================================
with tab_eb:
    st.markdown(
        '<div class="section-title">能量块买卖与交易复盘</div>',
        unsafe_allow_html=True,
    )

    # 优先使用已经完成月份、日期及其他公共条件筛选后的汇总数据。
    # 这样能量块复盘与顶部KPI、交易概览使用完全一致的数据口径，
    # 同时避免部分energy_block工作表日期字段不完整导致页面为空。
    eb = filtered[
        filtered["market_stage"].eq("energy_block")
    ].copy()

    if eb.empty:
        st.info("当前筛选条件下没有能量块交易数据。")
    else:
        eb["energy_mwh"] = safe_numeric(
            eb["energy_mwh"]
        ).fillna(0)

        # 能量块不存在环境权益，固定采用电能量价格口径。
        eb_amount_col = "energy_amount_yuan"

        if eb_amount_col not in eb.columns:
            eb[eb_amount_col] = 0

        eb[eb_amount_col] = safe_numeric(
            eb[eb_amount_col]
        ).fillna(0)

        eb["buy_energy_mwh"] = eb["energy_mwh"].clip(lower=0)
        eb["sell_energy_mwh"] = (-eb["energy_mwh"]).clip(lower=0)

        eb["buy_amount"] = np.where(
            eb["energy_mwh"] > 0,
            eb[eb_amount_col],
            0,
        )
        eb["sell_amount"] = np.where(
            eb["energy_mwh"] < 0,
            -eb[eb_amount_col],
            0,
        )

        buy_energy = float(eb["buy_energy_mwh"].sum())
        sell_energy = float(eb["sell_energy_mwh"].sum())
        buy_amount = float(eb["buy_amount"].sum())
        sell_amount = float(eb["sell_amount"].sum())

        buy_price = (
            buy_amount / buy_energy
            if buy_energy > 1e-9
            else np.nan
        )
        sell_price = (
            sell_amount / sell_energy
            if sell_energy > 1e-9
            else np.nan
        )
        eb_net_detail = float(eb["energy_mwh"].sum())

        a, b, c, d, e = st.columns(5, gap="medium")

        with a:
            render_kpi(
                "能量块净调整量",
                f"{eb_net_detail:,.0f} MWh",
                direction_text(eb_net_detail),
            )

        with b:
            render_kpi(
                "买入量",
                f"{buy_energy:,.0f} MWh",
                "能量块买入",
            )

        with c:
            render_kpi(
                "卖出量",
                f"{sell_energy:,.0f} MWh",
                "能量块卖出",
            )

        with d:
            render_kpi(
                "买入加权均价",
                "—"
                if pd.isna(buy_price)
                else f"{buy_price:,.2f} 元/MWh",
                price_mode,
            )

        with e:
            render_kpi(
                "卖出加权均价",
                "—"
                if pd.isna(sell_price)
                else f"{sell_price:,.2f} 元/MWh",
                price_mode,
            )

        eb_daily = (
            eb.groupby("display_date", as_index=False)
            .agg(
                净调整量_MWh=("energy_mwh", "sum"),
                买入量_MWh=("buy_energy_mwh", "sum"),
                卖出量_MWh=("sell_energy_mwh", "sum"),
                买入金额_元=("buy_amount", "sum"),
                卖出金额_元=("sell_amount", "sum"),
            )
            .sort_values("display_date")
        )

        eb_daily["买入均价_元每MWh"] = np.where(
            eb_daily["买入量_MWh"] > 1e-9,
            eb_daily["买入金额_元"] / eb_daily["买入量_MWh"],
            np.nan,
        )
        eb_daily["卖出均价_元每MWh"] = np.where(
            eb_daily["卖出量_MWh"] > 1e-9,
            eb_daily["卖出金额_元"] / eb_daily["卖出量_MWh"],
            np.nan,
        )

        st.markdown(
            '<div class="section-title">每日能量块买卖情况</div>',
            unsafe_allow_html=True,
        )

        fig_eb = go.Figure()

        buy_hover_text = []
        sell_hover_text = []

        for _, row in eb_daily.iterrows():
            date_text = pd.Timestamp(
                row["display_date"]
            ).strftime("%Y-%m-%d")

            buy_text = (
                f"日期：{date_text}<br>"
                f"买入量：{row['买入量_MWh']:,.2f} MWh"
            )
            if (
                row["买入量_MWh"] > 1e-9
                and pd.notna(row["买入均价_元每MWh"])
            ):
                buy_text += (
                    f"<br>买入均价："
                    f"{row['买入均价_元每MWh']:,.2f} 元/MWh"
                )
            buy_hover_text.append(buy_text)

            sell_text = (
                f"日期：{date_text}<br>"
                f"卖出量：{row['卖出量_MWh']:,.2f} MWh"
            )
            if (
                row["卖出量_MWh"] > 1e-9
                and pd.notna(row["卖出均价_元每MWh"])
            ):
                sell_text += (
                    f"<br>卖出均价："
                    f"{row['卖出均价_元每MWh']:,.2f} 元/MWh"
                )
            sell_hover_text.append(sell_text)

        fig_eb.add_trace(
            go.Bar(
                name="买入量",
                x=eb_daily["display_date"],
                y=eb_daily["买入量_MWh"],
                customdata=np.array(buy_hover_text).reshape(-1, 1),
                hovertemplate="%{customdata[0]}<extra></extra>",
            )
        )

        fig_eb.add_trace(
            go.Bar(
                name="卖出量",
                x=eb_daily["display_date"],
                y=-eb_daily["卖出量_MWh"],
                customdata=np.array(sell_hover_text).reshape(-1, 1),
                hovertemplate="%{customdata[0]}<extra></extra>",
            )
        )

        fig_eb.add_hline(y=0, line_width=1)
        fig_eb.update_xaxes(tickformat="%m-%d")
        fig_eb.update_layout(
            barmode="relative",
            xaxis_title="日期",
            yaxis_title="MWh",
            height=390,
            hovermode="x unified",
            margin=dict(l=10, r=10, t=30, b=20),
            legend_title_text="",
        )

        st.plotly_chart(
            fig_eb,
            use_container_width=True,
        )

        eb_table = eb_daily[
            [
                "display_date",
                "净调整量_MWh",
                "买入量_MWh",
                "卖出量_MWh",
                "买入均价_元每MWh",
                "卖出均价_元每MWh",
            ]
        ].copy()

        eb_table = eb_table.rename(
            columns={
                "display_date": "北京时间用电日",
                "净调整量_MWh": "净调整量（MWh）",
                "买入量_MWh": "买入量（MWh）",
                "卖出量_MWh": "卖出量（MWh）",
                "买入均价_元每MWh": "买入加权均价（元/MWh）",
                "卖出均价_元每MWh": "卖出加权均价（元/MWh）",
            }
        )

        st.dataframe(
            eb_table.style.format(
                {
                    "北京时间用电日": lambda x: (
                        pd.Timestamp(x).strftime("%Y-%m-%d")
                        if pd.notna(x)
                        else ""
                    ),
                    "净调整量（MWh）": "{:,.2f}",
                    "买入量（MWh）": "{:,.2f}",
                    "卖出量（MWh）": "{:,.2f}",
                    "买入加权均价（元/MWh）": "{:,.2f}",
                    "卖出加权均价（元/MWh）": "{:,.2f}",
                },
                na_rep="—",
            ),
            use_container_width=True,
            hide_index=True,
        )


# =========================================================
# Tab 5 中长期偏差预警
# =========================================================
with tab_risk:
    st.markdown(
        '<div class="section-title">中长期偏差收益回收预警</div>',
        unsafe_allow_html=True,
    )

    st.caption(
        "中长期电量口径：年协/月协净电量 + 能量块净调整量。"
        "日前+实时价格采用带正负号的净结算口径："
        "Σ(电量×电能量价格) ÷ Σ电量；买入为正、卖出为负。"
        "实际结算以江苏市场正式结算结果及最终参数为准。"
    )

    default_actual_energy = max(total_net_energy, 0.0)

    actual_energy = default_actual_energy

    st.caption(
        f"当前实际用电量自动采用：{actual_energy:,.2f} MWh。"
        "后续如接入正式结算用电量数据，可替换此口径。"
    )

    st.markdown("#### 计算参数")
    p1, p2, p3, p4 = st.columns(4)

    q1_pct = p1.number_input(
        "下限 q1（%）",
        min_value=0.0,
        max_value=100.0,
        value=70.0,
        step=1.0,
        format="%.1f",
    )

    q2_pct = p2.number_input(
        "上限 q2（%）",
        min_value=0.0,
        max_value=200.0,
        value=105.0,
        step=1.0,
        format="%.1f",
    )

    m_value = p3.number_input(
        "回收系数 m",
        min_value=0.0,
        value=1.2,
        step=0.1,
        format="%.2f",
    )

    alert_buffer_pct = p4.number_input(
        "临界预警范围（百分点）",
        min_value=0.0,
        max_value=20.0,
        value=3.0,
        step=0.5,
        format="%.1f",
    )

    q1 = q1_pct / 100
    q2 = q2_pct / 100
    alert_buffer = alert_buffer_pct / 100

    if q1 >= q2:
        st.error("参数错误：q1必须小于q2。")
    elif actual_energy <= 0:
        st.warning("实际用电量必须大于0，暂时无法计算覆盖率。")
    else:
        coverage_ratio = long_term_total_net / actual_energy

        r1, r2, r3, r4, r5 = st.columns(5)

        r1.metric(
            "年协/月协净电量",
            f"{lt_net:,.0f} MWh",
        )
        r2.metric(
            "能量块净调整量",
            f"{eb_net:,.0f} MWh",
        )
        r3.metric(
            "中长期净电量",
            f"{long_term_total_net:,.0f} MWh",
        )
        r4.metric(
            "实际用电量",
            f"{actual_energy:,.0f} MWh",
        )
        r5.metric(
            "中长期覆盖率",
            f"{coverage_ratio * 100:.2f}%",
        )

        # 覆盖率预警
        if coverage_ratio < q1:
            gap = q1 * actual_energy - long_term_total_net
            st.error(
                f"🔴 低于下限：当前覆盖率为{coverage_ratio * 100:.2f}%，"
                f"低于q1={q1_pct:.1f}%。"
                f"按当前实际用电量计算，距离下限仍缺少约{gap:,.2f} MWh。"
            )
        elif coverage_ratio > q2:
            excess = long_term_total_net - q2 * actual_energy
            st.error(
                f"🔴 高于上限：当前覆盖率为{coverage_ratio * 100:.2f}%，"
                f"高于q2={q2_pct:.1f}%。"
                f"超出上限约{excess:,.2f} MWh。"
            )
        elif coverage_ratio < q1 + alert_buffer:
            distance = (
                coverage_ratio - q1
            ) * actual_energy
            st.warning(
                f"🟡 接近下限：当前覆盖率为{coverage_ratio * 100:.2f}%，"
                f"距离q1={q1_pct:.1f}%仅剩"
                f"{(coverage_ratio - q1) * 100:.2f}个百分点，"
                f"对应电量缓冲约{distance:,.2f} MWh。"
            )
        elif coverage_ratio > q2 - alert_buffer:
            distance = (
                q2 - coverage_ratio
            ) * actual_energy
            st.warning(
                f"🟡 接近上限：当前覆盖率为{coverage_ratio * 100:.2f}%，"
                f"距离q2={q2_pct:.1f}%仅剩"
                f"{(q2 - coverage_ratio) * 100:.2f}个百分点，"
                f"对应电量缓冲约{distance:,.2f} MWh。"
            )
        else:
            st.success(
                f"🟢 当前覆盖率为{coverage_ratio * 100:.2f}%，"
                f"处于{q1_pct:.1f}%—{q2_pct:.1f}%正常区间，"
                "未触发中长期偏差收益回收测算。"
            )

        # 仅超范围时测算回收费用
        outside_range = (
            coverage_ratio < q1
            or coverage_ratio > q2
        )

        if not outside_range:
            st.markdown(
                """
<div class="risk-box">
当前中长期覆盖率未超出设定范围，因此不展开收益回收费用估算。
</div>
""",
                unsafe_allow_html=True,
            )
        else:
            st.markdown("#### 收益回收费用估算")

            if pd.isna(spot_weighted_price):
                st.warning(
                    "当前筛选范围无法计算日前+实时加权均价，"
                    "因此暂时不能估算收益回收费用。"
                )
            else:
                s1, s2 = st.columns(2)

                with s1:
                    st.metric(
                        "日前+实时净结算加权均价",
                        f"{spot_weighted_price:,.2f} 元/MWh",
                    )

                with s2:
                    if coverage_ratio < q1:
                        deviation_energy = (
                            q1 * actual_energy
                            - long_term_total_net
                        )
                        deviation_label = "低于下限的电量"
                    else:
                        deviation_energy = (
                            long_term_total_net
                            - q2 * actual_energy
                        )
                        deviation_label = "高于上限的电量"

                    st.metric(
                        deviation_label,
                        f"{deviation_energy:,.2f} MWh",
                    )

                price_input_mode = st.radio(
                    "全省中长期结算加权均价口径",
                    [
                        "区间估算（330-360元/MWh）",
                        "手动填写已知均价",
                    ],
                    horizontal=True,
                    key="market_lt_price_mode",
                )

                if price_input_mode == "区间估算（330-360元/MWh）":
                    c1, c2 = st.columns(2)

                    market_lt_price_min = c1.number_input(
                        "全省中长期均价下限（元/MWh）",
                        min_value=0.0,
                        value=330.0,
                        step=1.0,
                        format="%.2f",
                    )

                    market_lt_price_max = c2.number_input(
                        "全省中长期均价上限（元/MWh）",
                        min_value=0.0,
                        value=360.0,
                        step=1.0,
                        format="%.2f",
                    )

                    low_price = min(
                        market_lt_price_min,
                        market_lt_price_max,
                    )
                    high_price = max(
                        market_lt_price_min,
                        market_lt_price_max,
                    )

                    fee_at_low, deviation_low, risk_type = (
                        calculate_recovery_fee(
                            coverage_ratio=coverage_ratio,
                            actual_energy=actual_energy,
                            long_term_energy=long_term_total_net,
                            spot_price=spot_weighted_price,
                            market_lt_price=low_price,
                            q1=q1,
                            q2=q2,
                            m=m_value,
                        )
                    )

                    fee_at_high, deviation_high, _ = (
                        calculate_recovery_fee(
                            coverage_ratio=coverage_ratio,
                            actual_energy=actual_energy,
                            long_term_energy=long_term_total_net,
                            spot_price=spot_weighted_price,
                            market_lt_price=high_price,
                            q1=q1,
                            q2=q2,
                            m=m_value,
                        )
                    )

                    fee_range_low = min(
                        fee_at_low,
                        fee_at_high,
                    )
                    fee_range_high = max(
                        fee_at_low,
                        fee_at_high,
                    )

                    f1, f2, f3 = st.columns(3)

                    f1.metric(
                        "全省中长期均价假设",
                        f"{low_price:.2f}—{high_price:.2f} 元/MWh",
                    )
                    f2.metric(
                        "预计回收费用下限",
                        f"{fee_range_low:,.2f} 元",
                    )
                    f3.metric(
                        "预计回收费用上限",
                        f"{fee_range_high:,.2f} 元",
                    )

                    if fee_range_high <= 1e-9:
                        st.success(
                            "虽然覆盖率已经超出范围，但在当前价格假设区间内，"
                            "计算结果均不形成正收益，因此预计不回收。"
                        )
                    else:
                        st.error(
                            f"🔴 预计收益回收费用范围："
                            f"{fee_range_low:,.2f}—{fee_range_high:,.2f} 元。"
                        )

                    price_points = np.linspace(
                        low_price,
                        high_price,
                        num=7,
                    )

                    sensitivity_rows = []
                    for price_point in price_points:
                        fee_value, _, _ = calculate_recovery_fee(
                            coverage_ratio=coverage_ratio,
                            actual_energy=actual_energy,
                            long_term_energy=long_term_total_net,
                            spot_price=spot_weighted_price,
                            market_lt_price=float(price_point),
                            q1=q1,
                            q2=q2,
                            m=m_value,
                        )
                        sensitivity_rows.append(
                            {
                                "全省中长期均价(元/MWh)": round(
                                    float(price_point), 2
                                ),
                                "预计回收费用(元)": round(
                                    float(fee_value), 2
                                ),
                            }
                        )

                    sensitivity_df = pd.DataFrame(
                        sensitivity_rows
                    )

                    st.markdown("#### 价格敏感性")
                    st.dataframe(
                        sensitivity_df,
                        use_container_width=True,
                        hide_index=True,
                    )

                    fig_sensitivity = px.line(
                        sensitivity_df,
                        x="全省中长期均价(元/MWh)",
                        y="预计回收费用(元)",
                        markers=True,
                    )
                    fig_sensitivity.update_traces(
                        hovertemplate=(
                            "全省中长期均价：%{x:,.2f} 元/MWh<br>"
                            "预计回收费用：%{y:,.2f} 元"
                            "<extra></extra>"
                        )
                    )
                    fig_sensitivity.update_layout(
                        title="全省中长期均价变化对回收费用的影响",
                        xaxis_title="元/MWh",
                        yaxis_title="元",
                        height=390,
                    )

                    st.plotly_chart(
                        fig_sensitivity,
                        use_container_width=True,
                    )

                else:
                    market_lt_price_single = st.number_input(
                        "已知全省中长期结算加权均价（元/MWh）",
                        min_value=0.0,
                        value=345.0,
                        step=1.0,
                        format="%.2f",
                    )

                    recovery_fee, deviation_energy, risk_type = (
                        calculate_recovery_fee(
                            coverage_ratio=coverage_ratio,
                            actual_energy=actual_energy,
                            long_term_energy=long_term_total_net,
                            spot_price=spot_weighted_price,
                            market_lt_price=market_lt_price_single,
                            q1=q1,
                            q2=q2,
                            m=m_value,
                        )
                    )

                    f1, f2, f3 = st.columns(3)

                    f1.metric(
                        "全省中长期均价",
                        f"{market_lt_price_single:,.2f} 元/MWh",
                    )
                    f2.metric(
                        "偏差电量",
                        f"{deviation_energy:,.2f} MWh",
                    )
                    f3.metric(
                        "预计回收费用",
                        f"{recovery_fee:,.2f} 元",
                    )

                    if recovery_fee <= 1e-9:
                        st.success(
                            "按当前参数计算，预计收益回收费用为0元，"
                            "不触发回收。"
                        )
                    else:
                        st.error(
                            f"🔴 按当前参数预计回收约"
                            f"{recovery_fee:,.2f} 元。"
                        )

                if coverage_ratio < q1:
                    formula_text = (
                        "低于下限时：不足电量 × "
                        "（全省中长期均价 − 本企业日前/实时净结算加权均价）"
                        f" × m（{m_value:.2f}），计算结果小于0时不回收。"
                    )
                else:
                    formula_text = (
                        "高于上限时：超额电量 × "
                        "（本企业日前/实时净结算加权均价 − 全省中长期均价）"
                        f" × m（{m_value:.2f}），计算结果小于0时不回收。"
                    )

                st.markdown(
                    f"""
<div class="risk-box">
<strong>当前计算逻辑：</strong>{formula_text}<br>
全省中长期均价目前为假设值或手工输入值，正式结算结果以市场公布口径为准。
</div>
""",
                    unsafe_allow_html=True,
                )


# =========================================================
# Tab 6 明细与下载
# =========================================================
with tab_data:
    st.markdown(
        '<div class="section-title">筛选结果与数据下载</div>',
        unsafe_allow_html=True,
    )

    display_cols = [
        "trade_date",
        "display_date",
        "datetime_bj",
        "slot_index",
        "time_label",
        "market_stage_cn",
        "energy_mwh",
        "energy_price_yuan_mwh",
        "price_yuan_mwh",
        "energy_amount_yuan",
        "amount_yuan",
    ]

    detail = filtered.copy()

    if "trade_date" in detail.columns:
        detail["trade_date"] = pd.to_datetime(
            detail["trade_date"], errors="coerce"
        ).dt.strftime("%Y-%m-%d")

    detail["display_date"] = pd.to_datetime(
        detail["display_date"], errors="coerce"
    ).dt.strftime("%Y-%m-%d")

    detail["datetime_bj"] = pd.to_datetime(
        detail["datetime_bj"], errors="coerce"
    ).dt.strftime("%Y-%m-%d %H:%M:%S")

    detail = detail[
        [col for col in display_cols if col in detail.columns]
    ].copy()

    detail = detail.rename(
        columns={
            "trade_date": "原始交易日",
            "display_date": "北京时间用电日",
            "datetime_bj": "北京时间",
            "slot_index": "96点序号",
            "time_label": "交易时段",
            "market_stage_cn": "市场阶段",
            "energy_mwh": "交易电量(MWh)",
            "energy_price_yuan_mwh": "电能量成交价(元/MWh)",
            "price_yuan_mwh": "综合成交价(元/MWh)",
            "energy_amount_yuan": "电能量金额(元)",
            "amount_yuan": "综合成交金额(元)",
        }
    )

    st.dataframe(
        detail.round(3),
        use_container_width=True,
        hide_index=True,
    )

    csv_bytes = detail.to_csv(
        index=False,
        encoding="utf-8-sig",
    )

    excel_bytes = to_excel_bytes(
        {
            "交易明细": detail,
            "阶段汇总": display_summary,
        }
    )

    d1, d2 = st.columns(2)

    with d1:
        st.download_button(
            "下载 CSV",
            data=csv_bytes,
            file_name=(
                f"江苏交易复盘_{selected_month}_"
                f"{start_date}_{end_date}.csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )

    with d2:
        st.download_button(
            "下载 Excel",
            data=excel_bytes,
            file_name=(
                f"江苏交易复盘_{selected_month}_"
                f"{start_date}_{end_date}.xlsx"
            ),
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    st.markdown("#### 数据完整性检查")

    check_df = data.get("check", pd.DataFrame())

    if check_df.empty:
        st.caption("当前文件没有check sheet。")
    else:
        st.dataframe(
            check_df,
            use_container_width=True,
            hide_index=True,
        )


# =========================================================
# END
# =========================================================

