# -*- coding: utf-8 -*-
"""
江苏电力交易量价复盘平台 V3.2.16


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
   "deviation_recovery": "日前偏差费用",
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
   grid-template-columns: repeat(9, minmax(0, 1fr));
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
       "deviation_recovery",
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




# 法定节假日：仅作为分析维度，不改变峰平谷划分规则。
HOLIDAYS = pd.to_datetime([
    # 2025
    "2025-01-01",
    "2025-01-28", "2025-01-29", "2025-01-30", "2025-01-31",
    "2025-02-01", "2025-02-02", "2025-02-03", "2025-02-04",
    "2025-04-04", "2025-04-05", "2025-04-06",
    "2025-05-01", "2025-05-02", "2025-05-03", "2025-05-04", "2025-05-05",
    "2025-05-31", "2025-06-01", "2025-06-02",
    "2025-10-01", "2025-10-02", "2025-10-03", "2025-10-04",
    "2025-10-05", "2025-10-06", "2025-10-07", "2025-10-08",
    # 2026
    "2026-01-01", "2026-01-02", "2026-01-03",
    "2026-02-15", "2026-02-16", "2026-02-17", "2026-02-18", "2026-02-19",
    "2026-02-20", "2026-02-21", "2026-02-22", "2026-02-23",
    "2026-04-04", "2026-04-05", "2026-04-06",
    "2026-05-01", "2026-05-02", "2026-05-03", "2026-05-04", "2026-05-05",
    "2026-06-19", "2026-06-20", "2026-06-21",
    "2026-09-25", "2026-09-26", "2026-09-27",
    "2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04",
    "2026-10-05", "2026-10-06", "2026-10-07",
]).normalize()

DEFAULT_ANALYSIS_PARAMS = {
    "resale_normal_max": 0.10,
    "resale_general_max": 0.25,
    "resale_load_drop_max": 0.50,
    # 用户当前采用10元/MWh作为策略关注的最小异常价差。
    "strategy_price_gap": 10.0,
    "strategy_volume_ratio": 0.05,
    "strategy_impact_yuan": 3000.0,
}


def get_day_type(dt) -> str:
    if pd.isna(dt):
        return "未知"
    d = pd.Timestamp(dt).normalize()
    if d in HOLIDAYS:
        return "法定节假日"
    if d.weekday() >= 5:
        return "周末"
    return "工作日"


def classify_resale_ratio(ratio: float, params: dict) -> str:
    if pd.isna(ratio):
        return "无基础电量"
    if ratio < params["resale_normal_max"]:
        return "正常调整"
    if ratio < params["resale_general_max"]:
        return "一般调整"
    if ratio < params["resale_load_drop_max"]:
        return "明显负荷下降"
    return "重大负荷下降 / 疑似停机"


def add_period_shading(fig, selected_day, resolution: str):
    """
    为现有日内图增加峰/平/谷连续背景。

    关键点：Plotly分类轴上的类别值位于整数坐标0,1,2...的中心。
    因此阴影边界使用-0.5/+0.5的类别边界，而不是类别中心，
    这样相邻峰/平/谷区间会首尾相接，不再出现白色间隔。
    """
    day = pd.Timestamp(selected_day).normalize()

    points = []
    if resolution == "24小时":
        for h in range(24):
            dt = day + pd.Timedelta(hours=h)
            label = f"{h:02d}:00-{h + 1:02d}:00"
            points.append((label, period_type(dt)))
    else:
        for slot in range(1, 97):
            # 96点第1点为00:15，对应00:00-00:15区间；第96点显示24:00。
            label_dt = day + pd.Timedelta(minutes=15 * slot)
            label = "24:00" if slot == 96 else label_dt.strftime("%H:%M")
            class_dt = day + pd.Timedelta(minutes=15 * (slot - 1))
            points.append((label, period_type(class_dt)))

    if not points:
        return fig

    categories = [x[0] for x in points]
    fig.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=categories,
    )

    shade = {
        "谷": "rgba(76, 144, 255, 0.08)",
        "平": "rgba(120, 120, 120, 0.055)",
        "峰": "rgba(255, 153, 51, 0.09)",
    }

    start_idx = 0
    current = points[0][1]
    for i in range(1, len(points) + 1):
        changed = i == len(points) or points[i][1] != current
        if changed:
            # 分类坐标第k类中心是k，所以边界分别是k-0.5和k+0.5。
            fig.add_vrect(
                x0=start_idx - 0.5,
                x1=i - 0.5,
                fillcolor=shade.get(current, "rgba(0,0,0,0)"),
                opacity=1,
                line_width=0,
                layer="below",
            )
            if i < len(points):
                start_idx = i
                current = points[i][1]

    return fig


def _infer_load_change_stage(lt, eb, da, rt, params):
    """
    根据交易路径推测负荷下降信息在哪个交易阶段开始明显体现。

    只做数据特征解释，不直接认定停机原因。
    阶段判断使用“正常调整上限”作为最小有意义下修比例，默认10% LT。
    """
    if lt <= 1e-9:
        return "无法判断", "缺少有效年协/月协基础电量，无法根据交易链判断负荷变化时点。"

    min_signal = params["resale_normal_max"] * lt
    neg_eb = max(-eb, 0.0)
    neg_da = max(-da, 0.0)
    neg_rt = max(-rt, 0.0)

    eb_sig = neg_eb >= min_signal
    da_sig = neg_da >= min_signal
    rt_sig = neg_rt >= min_signal

    # 最典型：前面仍在买，直到实时才大卖。
    if eb >= -1e-9 and da >= -1e-9 and rt_sig:
        return (
            "实时阶段突发下修",
            "能量块和日前阶段仍为净买入/未明显卖出，但实时阶段突然出现大规模净卖出；"
            "说明负荷下降主要在临近实际运行时才体现，更符合临时负荷下降、非计划停机/故障或临时减产特征，建议结合纸机运行记录确认。",
        )

    # 能量块买入，后续现货整体转卖。
    if eb > 1e-9 and (da + rt) < -min_signal:
        if rt_sig and not da_sig:
            return (
                "能量块买入后实时转卖",
                "能量块阶段仍在增加采购，后续主要在实时市场转为卖出；"
                "说明较早阶段仍按较高需求安排电量，之后才出现明显需求下修。",
            )
        return (
            "能量块买入后现货转卖",
            "能量块阶段仍为净买入，但日前/实时合计转为明显净卖出；"
            "说明生产负荷预期在能量块交易之后被下调，需结合当日生产计划变化判断是计划调整还是临时停机。",
        )

    # 能量块未卖，日前开始明显下修。
    if eb >= -1e-9 and da_sig:
        if rt_sig:
            return (
                "日前开始并持续下修",
                "能量块阶段未明显卖出，日前开始转为大幅卖出，实时阶段继续下修；"
                "说明负荷下降信息可能在日前阶段开始明确，并在实际运行前进一步确认。",
            )
        return (
            "日前阶段发现",
            "能量块阶段未明显卖出，但日前阶段出现明显净卖出；"
            "说明负荷下降信息可能在日前交易前后才被确认，可能对应生产计划调整或较晚确认的停机安排。",
        )

    # 多个阶段连续负调整。
    sig_count = int(eb_sig) + int(da_sig) + int(rt_sig)
    if sig_count >= 2:
        return (
            "多阶段持续下修",
            "多个交易阶段连续出现明显净卖出，说明负荷预期并非一次性变化，而是在交易链中持续下修；"
            "可能对应生产计划逐步调整、停机信息逐步确认或负荷持续下降。",
        )

    if eb_sig:
        return (
            "能量块阶段提前发现",
            "能量块阶段已经出现明显净卖出，说明负荷下降较早已反映在交易安排中；"
            "相较于实时突发型，更接近计划性负荷下降、计划停机或提前获知的生产调整。",
        )

    if rt_sig:
        return (
            "实时阶段主要下修",
            "主要负调整集中在实时阶段，说明较晚阶段才出现明显需求下修，建议结合生产运行记录确认是否存在临时停机或负荷突降。",
        )

    if da_sig:
        return (
            "日前阶段主要下修",
            "主要负调整集中在日前阶段，说明需求变化在日前交易阶段已经较明显。",
        )

    return "无明显阶段性下修", "各阶段净调整均未达到当前设定的显著下修门槛。"


def build_analysis_dataset(df: pd.DataFrame, amount_col: str, params: dict):
    if df.empty:
        return pd.DataFrame()

    out = df.copy()
    out["datetime_bj"] = pd.to_datetime(out["datetime_bj"], errors="coerce")
    out["display_date"] = pd.to_datetime(out["display_date"], errors="coerce").dt.normalize()
    out = out.dropna(subset=["datetime_bj", "display_date"]).copy()
    out["峰平谷"] = out["datetime_bj"].apply(period_type)
    out["energy_mwh"] = safe_numeric(out["energy_mwh"]).fillna(0)
    out[amount_col] = safe_numeric(out[amount_col]).fillna(0)

    grouped = (
        out.groupby(["display_date", "峰平谷", "market_stage_cn"], as_index=False)
        .agg(
            energy_mwh=("energy_mwh", "sum"),
            amount_selected=(amount_col, "sum"),
        )
    )

    rows = []
    for (d, period), g in grouped.groupby(["display_date", "峰平谷"]):
        stage_energy = g.set_index("market_stage_cn")["energy_mwh"].to_dict()
        stage_amount = g.set_index("market_stage_cn")["amount_selected"].to_dict()

        lt = float(stage_energy.get("年协/月协", 0.0))
        eb = float(stage_energy.get("能量块", 0.0))
        da = float(stage_energy.get("日前", 0.0))
        rt = float(stage_energy.get("实时", 0.0))

        # 交易路径：逐阶段观察当时认为需要保留的净电量。
        pos_lt = lt
        pos_eb = lt + eb
        pos_da = pos_eb + da
        pos_rt = pos_da + rt

        subsequent = eb + da + rt
        resale_energy = max(-subsequent, 0.0)
        resale_ratio = resale_energy / lt if lt > 1e-9 else np.nan

        # 新增：阶段性大幅卖出不能被前序买入抵消。
        # 例如 LT=100, EB=+20, DA=+10, RT=-80：最终回售率虽只有50%，
        # 但实时阶段从130骤降到50，本身就是很强的晚发现负荷下降信号。
        eb_down = max(-eb, 0.0)
        da_down = max(-da, 0.0)
        rt_down = max(-rt, 0.0)
        max_stage_down = max(eb_down, da_down, rt_down)
        max_stage_down_ratio_lt = max_stage_down / lt if lt > 1e-9 else np.nan

        # 相对进入该阶段前的持仓，反映“这一阶段突然削减了多少”。
        eb_down_ratio_pre = eb_down / pos_lt if pos_lt > 1e-9 else np.nan
        da_down_ratio_pre = da_down / pos_eb if pos_eb > 1e-9 else np.nan
        rt_down_ratio_pre = rt_down / pos_da if pos_da > 1e-9 else np.nan

        candidate_ratios = [x for x in [resale_ratio, max_stage_down_ratio_lt] if pd.notna(x)]
        load_signal_ratio = max(candidate_ratios) if candidate_ratios else np.nan
        load_signal_class = classify_resale_ratio(load_signal_ratio, params)

        stage_name, stage_reason = _infer_load_change_stage(lt, eb, da, rt, params)

        total_energy = pos_rt
        total_amount = float(sum(stage_amount.values()))
        final_price = total_amount / total_energy if abs(total_energy) > 1e-9 else np.nan

        lt_amount = float(stage_amount.get("年协/月协", 0.0))
        lt_price = lt_amount / lt if lt > 1e-9 else np.nan

        def _stage_price(stage_cn):
            e = float(stage_energy.get(stage_cn, 0.0))
            a = float(stage_amount.get(stage_cn, 0.0))
            return a / e if abs(e) > 1e-9 else np.nan

        eb_price = _stage_price("能量块")
        da_price = _stage_price("日前")
        rt_price = _stage_price("实时")

        # 回售价格仍用于解释“基础电量低价回售”对最终综合均价的影响。
        later_sell_amount = 0.0
        later_sell_energy = 0.0
        for stage_name_cn in ["能量块", "日前", "实时"]:
            e = float(stage_energy.get(stage_name_cn, 0.0))
            a = float(stage_amount.get(stage_name_cn, 0.0))
            if e < -1e-9:
                later_sell_energy += -e
                later_sell_amount += -a

        resale_price = (
            later_sell_amount / later_sell_energy
            if later_sell_energy > 1e-9
            else np.nan
        )
        resale_loss = (
            resale_energy * max(lt_price - resale_price, 0.0)
            if resale_energy > 1e-9
            and pd.notna(lt_price)
            and pd.notna(resale_price)
            else 0.0
        )

        # 便于管理层阅读的交易链文本。
        path_text = (
            f"LT {pos_lt:,.0f} → 能量块后 {pos_eb:,.0f} → "
            f"日前后 {pos_da:,.0f} → 实时后 {pos_rt:,.0f} MWh"
        )

        rows.append({
            "日期": d,
            "日类型": get_day_type(d),
            "时段": period,
            "LT基础量_MWh": lt,
            "能量块调整_MWh": eb,
            "日前调整_MWh": da,
            "实时调整_MWh": rt,
            "现货调整_MWh": da + rt,
            "后续净调整_MWh": subsequent,
            "基础电量回售_MWh": resale_energy,
            "基础电量回售率": resale_ratio,
            "回售判断": classify_resale_ratio(resale_ratio, params),
            "最大阶段下修_MWh": max_stage_down,
            "最大阶段下修率_LT": max_stage_down_ratio_lt,
            "能量块阶段下修率_前序": eb_down_ratio_pre,
            "日前阶段下修率_前序": da_down_ratio_pre,
            "实时阶段下修率_前序": rt_down_ratio_pre,
            "负荷下降信号率": load_signal_ratio,
            "负荷下降判断": load_signal_class,
            "负荷下降发现阶段": stage_name,
            "可能原因解释": stage_reason,
            "交易路径": path_text,
            "LT后持仓_MWh": pos_lt,
            "能量块后持仓_MWh": pos_eb,
            "日前后持仓_MWh": pos_da,
            "实时后持仓_MWh": pos_rt,
            "LT加权均价": lt_price,
            "能量块加权均价": eb_price,
            "日前加权均价": da_price,
            "实时加权均价": rt_price,
            "回售等效均价": resale_price,
            "总交易金额_元": total_amount,
            "回售价格损失估算_元": resale_loss,
            "最终净量_MWh": total_energy,
            "最终综合均价": final_price,
        })

    return pd.DataFrame(rows)


def build_strategy_watch(df: pd.DataFrame, params: dict):
    if df.empty:
        return pd.DataFrame()

    out = df.copy()
    out["datetime_bj"] = pd.to_datetime(out["datetime_bj"], errors="coerce")
    out["display_date"] = pd.to_datetime(out["display_date"], errors="coerce").dt.normalize()
    out = out.dropna(subset=["datetime_bj", "display_date"]).copy()
    out["峰平谷"] = out["datetime_bj"].apply(period_type)
    out["energy_mwh"] = safe_numeric(out["energy_mwh"]).fillna(0)
    out["energy_price_yuan_mwh"] = safe_numeric(out["energy_price_yuan_mwh"])
    out = out[out["energy_price_yuan_mwh"].notna()].copy()
    out["signed_amount"] = out["energy_mwh"] * out["energy_price_yuan_mwh"]

    rows = []
    for (d, period), g in out.groupby(["display_date", "峰平谷"]):
        lt_energy = float(g.loc[g["market_stage"].eq("long_term"), "energy_mwh"].sum())
        eb_buy = g[
            (g["market_stage"].eq("energy_block"))
            & (g["energy_mwh"] > 1e-9)
        ].copy()
        if eb_buy.empty:
            continue

        eb_energy = float(eb_buy["energy_mwh"].sum())
        eb_price = (
            float(eb_buy["signed_amount"].sum() / eb_energy)
            if eb_energy > 1e-9
            else np.nan
        )

        later = g[g["market_stage"].isin(["day_ahead", "real_time"])].copy()
        later_energy_abs = float(later["energy_mwh"].abs().sum())
        if later.empty or later_energy_abs <= 1e-9:
            continue

        later_price = float(
            (later["energy_mwh"].abs() * later["energy_price_yuan_mwh"]).sum()
            / later_energy_abs
        )
        gap = eb_price - later_price
        volume_ratio = eb_energy / lt_energy if lt_energy > 1e-9 else np.nan
        impact = max(gap, 0.0) * eb_energy

        is_watch = (
            gap >= params["strategy_price_gap"]
            and (
                pd.isna(volume_ratio)
                or volume_ratio >= params["strategy_volume_ratio"]
            )
            and impact >= params["strategy_impact_yuan"]
        )

        if is_watch:
            rows.append({
                "日期": d,
                "日类型": get_day_type(d),
                "时段": period,
                "能量块买入量_MWh": eb_energy,
                "占LT基础量": volume_ratio,
                "能量块买入均价": eb_price,
                "后续日前实时参考均价": later_price,
                "价差_元每MWh": gap,
                "潜在影响_元": impact,
            })

    return pd.DataFrame(rows)



# =========================================================
# 纸机停机历史与归因分析
# =========================================================
SHUTDOWN_FILE_NAME = "PM_shutdown_history.xlsx"
PM_LOAD_MW = {
    "PM1": 31.0,
    "PM2": 28.0,
    "PM3": 23.0,
}

# 连续性策略复盘默认参数；均可在分析结论Tab中调整。
DEFAULT_ANALYSIS_PARAMS.update({
    "strategy_repeat_count": 3,
    "strategy_repeat_ratio": 0.30,
    "strategy_cumulative_impact_yuan": 20000.0,
})


def find_shutdown_file(base_dir: Path) -> Optional[Path]:
    """优先读取data/PM_shutdown_history.xlsx，也兼容子目录同名文件。"""
    direct = base_dir / SHUTDOWN_FILE_NAME
    if direct.exists():
        return direct
    matches = list(base_dir.rglob(SHUTDOWN_FILE_NAME)) if base_dir.exists() else []
    return matches[0] if matches else None


@st.cache_data(show_spinner=False)
def load_shutdown_history(file_path: str, modified_time: float) -> pd.DataFrame:
    """
    读取停机历史。模板字段：
    machine / stop_start / stop_end / shutdown_type(可选) / remark / load_mw_override(可选)
    shutdown_type支持：计划 / 非计划 / 不确定；缺失时默认“不确定”。
    注意：该字段仅用于事件说明和复盘，不参与停机对均价影响的计算。
    stop_end建议填写“大致恢复正常生产负荷”的时间，而不仅是设备开始启动的时间。
    同时兼容常见中文列名。
    """
    _ = modified_time
    raw = pd.read_excel(file_path, sheet_name=0)
    raw = clean_columns(raw)

    aliases = {
        "machine": ["machine", "纸机", "纸机名称", "machine_name"],
        "stop_start": ["stop_start", "停机开始", "开始时间", "start", "start_time"],
        "stop_end": ["stop_end", "停机结束", "结束时间", "end", "end_time"],
        "shutdown_type": ["shutdown_type", "停机类型", "类型", "plan_type"],
        "remark": ["remark", "备注", "说明", "comment"],
        "load_mw_override": ["load_mw_override", "负荷MW", "负荷_mw", "load_mw", "estimated_load_mw"],
    }

    rename_map = {}
    lowered = {str(c).strip().lower(): c for c in raw.columns}
    for target, candidates in aliases.items():
        for cand in candidates:
            key = cand.lower()
            if key in lowered:
                rename_map[lowered[key]] = target
                break
    raw = raw.rename(columns=rename_map)

    required = ["machine", "stop_start", "stop_end"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(
            "停机历史Excel缺少必需列：" + ", ".join(missing)
            + "。请使用提供的PM_shutdown_history.xlsx模板。"
        )

    if "shutdown_type" not in raw.columns:
        raw["shutdown_type"] = "不确定"
    if "remark" not in raw.columns:
        raw["remark"] = ""
    if "load_mw_override" not in raw.columns:
        raw["load_mw_override"] = np.nan

    raw["machine"] = raw["machine"].astype(str).str.strip().str.upper()
    raw["stop_start"] = pd.to_datetime(raw["stop_start"], errors="coerce")
    raw["stop_end"] = pd.to_datetime(raw["stop_end"], errors="coerce")
    raw["load_mw_override"] = safe_numeric(raw["load_mw_override"])
    raw["remark"] = raw["remark"].fillna("").astype(str)

    type_map = {
        "计划": "计划", "计划停机": "计划", "planned": "计划", "plan": "计划",
        "非计划": "非计划", "非计划停机": "非计划", "unplanned": "非计划", "unexpected": "非计划",
        "不确定": "不确定", "未知": "不确定", "unknown": "不确定", "": "不确定",
    }
    raw["shutdown_type"] = (
        raw["shutdown_type"].fillna("不确定").astype(str).str.strip()
        .map(lambda x: type_map.get(x.lower(), type_map.get(x, "不确定")))
    )

    raw = raw[
        raw["machine"].isin(PM_LOAD_MW.keys())
        & raw["stop_start"].notna()
        & raw["stop_end"].notna()
        & raw["stop_end"].gt(raw["stop_start"])
    ].copy()

    raw["load_mw"] = raw.apply(
        lambda r: float(r["load_mw_override"])
        if pd.notna(r["load_mw_override"]) and float(r["load_mw_override"]) > 0
        else PM_LOAD_MW.get(r["machine"], np.nan),
        axis=1,
    )
    return raw.reset_index(drop=True)


def split_shutdown_to_periods(
    shutdown_df: pd.DataFrame,
    analysis_start,
    analysis_end,
) -> pd.DataFrame:
    """
    将任意分钟级、可跨天的停机区间按15分钟交易区间拆分，
    再归入对应日期的峰/平/谷。重叠分钟按实际分钟计算，不把13:20粗化成整小时。
    """
    if shutdown_df is None or shutdown_df.empty or pd.isna(analysis_end):
        return pd.DataFrame()

    window_start = pd.Timestamp(analysis_start).normalize()
    window_end = pd.Timestamp(analysis_end).normalize() + pd.Timedelta(days=1)
    rows = []

    for idx, r in shutdown_df.iterrows():
        start = max(pd.Timestamp(r["stop_start"]), window_start)
        end = min(pd.Timestamp(r["stop_end"]), window_end)
        if end <= start:
            continue

        cursor = start.floor("15min")
        while cursor < end:
            slot_end = cursor + pd.Timedelta(minutes=15)
            overlap_start = max(start, cursor)
            overlap_end = min(end, slot_end)
            minutes = max((overlap_end - overlap_start).total_seconds() / 60.0, 0.0)
            if minutes > 1e-9:
                period = period_type(cursor)
                load_mw = float(r["load_mw"])
                rows.append({
                    "shutdown_row": int(idx),
                    "日期": cursor.normalize(),
                    "时段": period,
                    "machine": r["machine"],
                    "stop_start": r["stop_start"],
                    "stop_end": r["stop_end"],
                    "shutdown_type": r.get("shutdown_type", "不确定"),
                    "remark": r.get("remark", ""),
                    "load_mw": load_mw,
                    "overlap_minutes": minutes,
                    "理论负荷下降_MWh": load_mw * minutes / 60.0,
                })
            cursor = slot_end

    return pd.DataFrame(rows)


def _counterfactual_reference_price(row: pd.Series) -> tuple[float, str]:
    """
    为“无停机情景”中仍需补购的电量选择参考价格。
    原则：优先使用交易链显示负荷变化主要被反映的市场阶段；
    若该阶段价格缺失，则按实时→日前→能量块→LT回退。
    这是运营复盘反事实假设，不代表当时一定会以该价格成交。
    """
    stage_text = str(row.get("负荷下降发现阶段", ""))
    candidates = []
    if "实时" in stage_text:
        candidates.append((row.get("实时加权均价", np.nan), "实时市场参考价"))
    elif "日前" in stage_text:
        candidates.append((row.get("日前加权均价", np.nan), "日前市场参考价"))
    elif "能量块" in stage_text:
        candidates.append((row.get("能量块加权均价", np.nan), "能量块参考价"))
    elif "多阶段" in stage_text:
        vals = []
        for c in ["能量块加权均价", "日前加权均价", "实时加权均价"]:
            v = row.get(c, np.nan)
            if pd.notna(v):
                vals.append(float(v))
        if vals:
            candidates.append((float(np.mean(vals)), "多阶段市场价格均值"))

    candidates.extend([
        (row.get("实时加权均价", np.nan), "实时市场参考价"),
        (row.get("日前加权均价", np.nan), "日前市场参考价"),
        (row.get("能量块加权均价", np.nan), "能量块参考价"),
        (row.get("LT加权均价", np.nan), "LT参考价"),
    ])
    for value, label in candidates:
        if pd.notna(value) and float(value) > 0:
            return float(value), label
    return np.nan, "无可用参考价"


def build_shutdown_period_summary(
    shutdown_split: pd.DataFrame,
    analysis_detail: pd.DataFrame,
) -> pd.DataFrame:
    """
    把已知停机与同日同峰平谷交易链对齐，形成生产事件归因数据。

    反事实逻辑：
    - 实际情景：直接使用真实交易净量与真实交易金额；
    - 无停机情景：实际最终净量 + 理论停机负荷下降；
    - 对已观察到、且可由停机解释的卖出，先按实际回售价“撤销回售”；
    - 若理论停机负荷仍高于可解释卖出，则剩余正常需求按负荷下降主要反映阶段的市场参考价补购；
    - 由此得到同日同峰平谷的无停机估算净量、总成本和综合均价。
    """
    if shutdown_split is None or shutdown_split.empty:
        return pd.DataFrame()

    machine_info = (
        shutdown_split.groupby(["日期", "时段"], as_index=False)
        .agg(
            理论停机负荷下降_MWh=("理论负荷下降_MWh", "sum"),
            停机机组=("machine", lambda s: "、".join(sorted(set(s.astype(str))))),
            停机类型=("shutdown_type", lambda s: "、".join(sorted(set(s.astype(str))))),
            停机机组记录数=("shutdown_row", "nunique"),
        )
    )

    if analysis_detail is None or analysis_detail.empty:
        return machine_info

    merged = machine_info.merge(analysis_detail, on=["日期", "时段"], how="left")

    # 各阶段负调整之和：避免“前面买、实时大卖”被最终净额抵消。
    merged["观察到的阶段卖出_MWh"] = (
        (-merged["能量块调整_MWh"].fillna(0)).clip(lower=0)
        + (-merged["日前调整_MWh"].fillna(0)).clip(lower=0)
        + (-merged["实时调整_MWh"].fillna(0)).clip(lower=0)
    )
    merged["停机可解释卖出_MWh"] = np.minimum(
        merged["理论停机负荷下降_MWh"].fillna(0),
        merged["观察到的阶段卖出_MWh"].fillna(0),
    )
    merged["生产事件解释率"] = np.where(
        merged["观察到的阶段卖出_MWh"] > 1e-9,
        merged["停机可解释卖出_MWh"] / merged["观察到的阶段卖出_MWh"],
        np.nan,
    )

    ref = merged.apply(_counterfactual_reference_price, axis=1)
    merged["无停机剩余补购参考价_元每MWh"] = [x[0] for x in ref]
    merged["无停机补购参考口径"] = [x[1] for x in ref]

    # 反事实第一步：撤销可由停机解释的实际回售。
    merged["撤销停机回售电量_MWh"] = merged["停机可解释卖出_MWh"]
    merged["撤销停机回售成本_元"] = np.where(
        merged["撤销停机回售电量_MWh"].gt(1e-9)
        & merged["回售等效均价"].notna(),
        merged["撤销停机回售电量_MWh"] * merged["回售等效均价"],
        0.0,
    )

    # 第二步：若理论停机减负荷 > 已观察到的停机相关卖出，
    # 正常运行时还需要额外补购剩余需求。
    merged["无停机额外补购电量_MWh"] = (
        merged["理论停机负荷下降_MWh"] - merged["撤销停机回售电量_MWh"]
    ).clip(lower=0)
    merged["无停机额外补购成本_元"] = np.where(
        merged["无停机额外补购电量_MWh"].gt(1e-9)
        & merged["无停机剩余补购参考价_元每MWh"].notna(),
        merged["无停机额外补购电量_MWh"]
        * merged["无停机剩余补购参考价_元每MWh"],
        0.0,
    )

    merged["无停机反事实恢复成本_元"] = (
        merged["撤销停机回售成本_元"] + merged["无停机额外补购成本_元"]
    )
    merged["无停机估算净量_MWh"] = (
        merged["最终净量_MWh"].fillna(0) + merged["理论停机负荷下降_MWh"].fillna(0)
    )
    merged["无停机估算总成本_元"] = (
        merged["总交易金额_元"].fillna(0) + merged["无停机反事实恢复成本_元"]
    )
    merged["无停机估算综合均价_元每MWh"] = np.where(
        merged["无停机估算净量_MWh"].abs() > 1e-9,
        merged["无停机估算总成本_元"] / merged["无停机估算净量_MWh"],
        np.nan,
    )
    merged["停机对该日时段均价影响_元每MWh"] = (
        merged["最终综合均价"] - merged["无停机估算综合均价_元每MWh"]
    )

    # 仅作为“均价影响的金额尺度”，不是总电费同比变化（因为两情景用电量不同）。
    merged["停机均价影响折算_元"] = (
        merged["停机对该日时段均价影响_元每MWh"]
        * merged["最终净量_MWh"].fillna(0)
    )
    return merged


def build_shutdown_event_detail(
    shutdown_raw: pd.DataFrame,
    shutdown_split: pd.DataFrame,
    shutdown_period_summary: pd.DataFrame,
) -> pd.DataFrame:
    """按每条实际停机记录输出明细，同时汇总其峰/平/谷理论影响。"""
    if shutdown_raw is None or shutdown_raw.empty:
        return pd.DataFrame()

    rows = []
    for idx, r in shutdown_raw.iterrows():
        part = shutdown_split[shutdown_split["shutdown_row"].eq(idx)].copy()
        if part.empty:
            continue
        period_energy = part.groupby("时段")["理论负荷下降_MWh"].sum().to_dict()
        rows.append({
            "纸机": r["machine"],
            "停机类型": r.get("shutdown_type", "不确定"),
            "停机开始": r["stop_start"],
            "停机结束": r["stop_end"],
            "持续小时": (r["stop_end"] - r["stop_start"]).total_seconds() / 3600.0,
            "采用负荷_MW": r["load_mw"],
            "理论负荷下降_MWh": float(part["理论负荷下降_MWh"].sum()),
            "峰段_MWh": float(period_energy.get("峰", 0.0)),
            "平段_MWh": float(period_energy.get("平", 0.0)),
            "谷段_MWh": float(period_energy.get("谷", 0.0)),
            "备注": r.get("remark", ""),
        })
    return pd.DataFrame(rows)


def calculate_shutdown_period_impact(
    analysis_df: pd.DataFrame,
    amount_col: str,
    shutdown_period_summary: pd.DataFrame,
) -> pd.DataFrame:
    """
    峰/平/谷月度反事实比较：
    实际综合均价 vs 假设已知停机没有发生时的估算综合均价。
    正值影响表示停机使该时段实际综合均价高于无停机估算均价。
    """
    actual_prices = calculate_period_weighted_prices(analysis_df, amount_col)
    base = analysis_df.copy()
    base["峰平谷"] = pd.to_datetime(base["datetime_bj"], errors="coerce").apply(period_type)
    base["energy_mwh"] = safe_numeric(base["energy_mwh"]).fillna(0)
    base[amount_col] = safe_numeric(base[amount_col]).fillna(0)
    actual_energy = base.groupby("峰平谷")["energy_mwh"].sum().to_dict()
    actual_amount = base.groupby("峰平谷")[amount_col].sum().to_dict()

    rows = []
    for p in ["峰", "平", "谷"]:
        e_actual = float(actual_energy.get(p, 0.0))
        a_actual = float(actual_amount.get(p, 0.0))
        actual_price = actual_prices.get(p, np.nan)

        if shutdown_period_summary is None or shutdown_period_summary.empty:
            q_restore = 0.0
            restore_cost = 0.0
        else:
            part = shutdown_period_summary[shutdown_period_summary["时段"].eq(p)].copy()
            q_restore = float(part["理论停机负荷下降_MWh"].fillna(0).sum())
            restore_cost = float(part["无停机反事实恢复成本_元"].fillna(0).sum())

        e_normal = e_actual + q_restore
        a_normal = a_actual + restore_cost
        normal_price = a_normal / e_normal if abs(e_normal) > 1e-9 else np.nan
        price_impact = (
            actual_price - normal_price
            if pd.notna(actual_price) and pd.notna(normal_price)
            else np.nan
        )
        equivalent_yuan = price_impact * e_actual if pd.notna(price_impact) else np.nan

        rows.append({
            "时段": p,
            "实际综合均价_元每MWh": actual_price,
            "实际最终净电量_MWh": e_actual,
            "无停机恢复电量估算_MWh": q_restore,
            "无停机反事实恢复成本_元": restore_cost,
            "无停机估算净电量_MWh": e_normal,
            "无停机估算综合均价_元每MWh": normal_price,
            "停机对均价影响估算_元每MWh": price_impact,
            "停机均价影响折算_元": equivalent_yuan,
        })
    return pd.DataFrame(rows)


def build_strategy_review_events(df: pd.DataFrame, params: dict) -> pd.DataFrame:
    """
    事后交易时点复盘：
    1) 能量块买入价 vs 后续日前/实时参考价；
    2) 日前买入价 vs 后续实时参考价。
    只在价差、电量占比、金额同时达到阈值时保留。
    这不是“失误认定”，因为交易当时受负荷预测、价格预测和合同覆盖率约束。
    """
    if df.empty:
        return pd.DataFrame()

    out = df.copy()
    out["datetime_bj"] = pd.to_datetime(out["datetime_bj"], errors="coerce")
    out["display_date"] = pd.to_datetime(out["display_date"], errors="coerce").dt.normalize()
    out = out.dropna(subset=["datetime_bj", "display_date"]).copy()
    out["峰平谷"] = out["datetime_bj"].apply(period_type)
    out["energy_mwh"] = safe_numeric(out["energy_mwh"]).fillna(0)
    out["energy_price_yuan_mwh"] = safe_numeric(out["energy_price_yuan_mwh"])
    out = out[out["energy_price_yuan_mwh"].notna()].copy()

    rows = []
    for (d, period), g in out.groupby(["display_date", "峰平谷"]):
        lt = float(g.loc[g["market_stage"].eq("long_term"), "energy_mwh"].sum())

        def add_case(stage_code, stage_name, later_codes, case_name):
            buy = g[(g["market_stage"].eq(stage_code)) & (g["energy_mwh"] > 1e-9)].copy()
            if buy.empty:
                return
            buy_energy = float(buy["energy_mwh"].sum())
            buy_price = float(
                (buy["energy_mwh"] * buy["energy_price_yuan_mwh"]).sum() / buy_energy
            )
            later = g[g["market_stage"].isin(later_codes)].copy()
            later_abs = float(later["energy_mwh"].abs().sum())
            if later.empty or later_abs <= 1e-9:
                return
            later_price = float(
                (later["energy_mwh"].abs() * later["energy_price_yuan_mwh"]).sum() / later_abs
            )
            gap = buy_price - later_price
            ratio = buy_energy / lt if lt > 1e-9 else np.nan
            impact = max(gap, 0.0) * buy_energy
            if (
                gap >= params["strategy_price_gap"]
                and (pd.isna(ratio) or ratio >= params["strategy_volume_ratio"])
                and impact >= params["strategy_impact_yuan"]
            ):
                rows.append({
                    "日期": d,
                    "日类型": get_day_type(d),
                    "时段": period,
                    "复盘类型": case_name,
                    "前序市场": stage_name,
                    "买入量_MWh": buy_energy,
                    "占LT基础量": ratio,
                    "前序买入均价": buy_price,
                    "后续市场参考均价": later_price,
                    "价差_元每MWh": gap,
                    "潜在影响_元": impact,
                })

        add_case("energy_block", "能量块", ["day_ahead", "real_time"], "能量块提前买入后后续市场更低")
        add_case("day_ahead", "日前", ["real_time"], "日前买入后实时市场更低")

    return pd.DataFrame(rows)


def _max_consecutive_days(dates) -> int:
    ds = sorted(pd.to_datetime(pd.Series(list(dates)), errors="coerce").dropna().dt.normalize().unique())
    if not ds:
        return 0
    best = cur = 1
    for prev, now in zip(ds[:-1], ds[1:]):
        if (pd.Timestamp(now) - pd.Timestamp(prev)).days == 1:
            cur += 1
            best = max(best, cur)
        else:
            cur = 1
    return best


def summarize_strategy_persistence(
    strategy_events: pd.DataFrame,
    eligible_df: pd.DataFrame,
    params: dict,
) -> pd.DataFrame:
    """把偶发不利结果和持续性模式分开。"""
    if strategy_events is None or strategy_events.empty:
        return pd.DataFrame()

    rows = []
    for (period, case), g in strategy_events.groupby(["时段", "复盘类型"]):
        event_days = g["日期"].nunique()
        consecutive = _max_consecutive_days(g["日期"])
        cumulative = float(g["潜在影响_元"].sum())

        # 分母用分析期有效实时交易日，避免月底尚无实时数据日期进入比例。
        eligible_days = int(eligible_df["display_date"].nunique()) if not eligible_df.empty else event_days
        ratio = event_days / eligible_days if eligible_days > 0 else np.nan

        persistent = (
            event_days >= int(params["strategy_repeat_count"])
            and (
                consecutive >= int(params["strategy_repeat_count"])
                or (pd.notna(ratio) and ratio >= params["strategy_repeat_ratio"])
                or cumulative >= params["strategy_cumulative_impact_yuan"]
            )
        )
        level = "持续性策略复盘警告" if persistent else "偶发不利结果"
        rows.append({
            "时段": period,
            "复盘类型": case,
            "不利事件天数": event_days,
            "分析期有效天数": eligible_days,
            "发生比例": ratio,
            "最长连续天数": consecutive,
            "累计潜在影响_元": cumulative,
            "判断": level,
        })
    return pd.DataFrame(rows).sort_values(["判断", "累计潜在影响_元"], ascending=[False, False])

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
# 到户电费模拟参数与计算
# =========================================================
DEFAULT_DELIVERY_PARAMS = {
   "peak_up_ratio": 0.80,
   "valley_down_ratio": 0.65,
   "line_loss_rate": 0.0318,
   "transmission_price": 0.0857,
   "fund_surcharge": 0.0294,
}


DEFAULT_SYSTEM_FEE_DETAIL = {
   "辅助服务费用": 0.0000,
   "抽水蓄能容量电费": 0.0052,
   "上网环节线损代理采购损益": -0.0021,
   "电价交叉补贴新增损益": 0.0074,
   "力调电费损益": 0.0032,
   "天然气发电容量电费（含气电联动）": 0.0230,
   "峰谷分时电价损益": -0.0028,
   "电力保供购电费用": 0.0001,
   "煤电容量电费": 0.0350,
   "新能源机制差价分摊电费": 0.0186,
}




def calc_delivery_price(flat_price_mwh, params):
   if pd.isna(flat_price_mwh):
       return None


   market_price = flat_price_mwh / 1000


   line_loss = (
       market_price
       * params["line_loss_rate"]
       / (1 - params["line_loss_rate"])
   )


   adders = (
       line_loss
       + params["transmission_price"]
       + params["system_fee"]
       + params["fund_surcharge"]
   )


   return {
       "market": market_price,
       "line_loss": line_loss,
       "peak": market_price * (1 + params["peak_up_ratio"]) + adders,
       "flat": market_price + adders,
       "valley": market_price * (1 - params["valley_down_ratio"]) + adders,
       "adders": adders,
   }








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





def calculate_strategy_profit(df: pd.DataFrame) -> float:
   """
   日前/实时交易策略收益。

   按当前所选时段整体复盘口径计算：
   1. 分别计算日前、实时市场的净电量与带符号电能量金额；
   2. 加权均价 = 带符号电能量金额 / 净电量；
   3. 若实时净卖出：收益 = |实时净卖出量| × (实时加权均价 - 日前加权均价)；
   4. 若实时净买入：收益 = 实时净买入量 × (日前加权均价 - 实时加权均价)。

   正值表示策略收益，负值表示策略损失。
   """
   spot = df.loc[
       df["market_stage"].isin(["day_ahead", "real_time"]),
       ["market_stage", "energy_mwh", "energy_price_yuan_mwh"],
   ].copy()

   if spot.empty:
       return 0.0

   spot["energy_mwh"] = safe_numeric(spot["energy_mwh"]).fillna(0)
   spot["energy_price_yuan_mwh"] = safe_numeric(
       spot["energy_price_yuan_mwh"]
   )
   spot = spot[spot["energy_price_yuan_mwh"].notna()].copy()

   if spot.empty:
       return 0.0

   spot["signed_amount_yuan"] = (
       spot["energy_mwh"] * spot["energy_price_yuan_mwh"]
   )

   stage = (
       spot.groupby("market_stage", as_index=False)
       .agg(
           net_energy_mwh=("energy_mwh", "sum"),
           signed_amount_yuan=("signed_amount_yuan", "sum"),
       )
   )

   stage["weighted_price"] = np.where(
       stage["net_energy_mwh"].abs() > 1e-9,
       stage["signed_amount_yuan"] / stage["net_energy_mwh"],
       np.nan,
   )

   da_row = stage[stage["market_stage"] == "day_ahead"]
   rt_row = stage[stage["market_stage"] == "real_time"]

   if da_row.empty or rt_row.empty:
       return 0.0

   da_price = float(da_row.iloc[0]["weighted_price"])
   rt_price = float(rt_row.iloc[0]["weighted_price"])
   rt_net = float(rt_row.iloc[0]["net_energy_mwh"])

   if pd.isna(da_price) or pd.isna(rt_price) or abs(rt_net) <= 1e-9:
       return 0.0

   if rt_net < 0:
       # 实时净卖出：实时价高于日前价为收益
       return abs(rt_net) * (rt_price - da_price)

   # 实时净买入：实时价低于日前价为收益
   return rt_net * (da_price - rt_price)

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
   deviation_raw = data.get("deviation_recovery", pd.DataFrame())


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


# 日前偏差费用：直接读取交易中心结算结果
if not deviation_raw.empty:
   deviation_raw = derive_trade_fields(deviation_raw)
   deviation_filtered = deviation_raw[
       deviation_raw["display_date"].between(start_ts, end_ts)
   ].copy()
else:
   deviation_filtered = pd.DataFrame()

if "deviation_recovery_yuan" in deviation_filtered.columns:
   deviation_fee_total = float(
       safe_numeric(deviation_filtered["deviation_recovery_yuan"])
       .fillna(0).sum()
   )
else:
   deviation_fee_total = 0.0


# 偏差电量读取新版清洗字段：
# 优先读取 deviation_energy_mwh，
# 兼容旧版本 energy_mwh。
if (
    not deviation_filtered.empty
    and "deviation_energy_mwh" in deviation_filtered.columns
):
    deviation_energy_total = float(
        safe_numeric(
            deviation_filtered["deviation_energy_mwh"]
        ).fillna(0).sum()
    )
elif (
    not deviation_filtered.empty
    and "energy_mwh" in deviation_filtered.columns
):
    deviation_energy_total = float(
        safe_numeric(
            deviation_filtered["energy_mwh"]
        ).fillna(0).sum()
    )
else:
    deviation_energy_total = 0.0


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

# 实际结算成本 = 市场交易费用 + 日前偏差费用
actual_settlement_cost = total_amount + deviation_fee_total


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

strategy_profit = calculate_strategy_profit(filtered)


# 新增分析结论的数据截止日：只分析到存在有效实时量/价数据的最后一天。
_rt_for_analysis = filtered[filtered["market_stage"].eq("real_time")].copy()
_rt_valid_mask = (
    _rt_for_analysis["energy_mwh"].abs().gt(1e-9)
    | _rt_for_analysis["energy_price_yuan_mwh"].abs().gt(1e-9)
)
_rt_for_analysis = _rt_for_analysis[_rt_valid_mask]
analysis_end_date = (
    pd.to_datetime(_rt_for_analysis["display_date"], errors="coerce").max()
    if not _rt_for_analysis.empty else pd.NaT
)
analysis_filtered = (
    filtered[filtered["display_date"].le(analysis_end_date)].copy()
    if pd.notna(analysis_end_date) else pd.DataFrame(columns=filtered.columns)
)


# 已知纸机停机记录：独立Excel维护，不改变交易底表。
shutdown_file = find_shutdown_file(BASE_DIR)
shutdown_raw = pd.DataFrame()
shutdown_load_error = ""
if shutdown_file is not None:
    try:
        shutdown_raw = load_shutdown_history(
            str(shutdown_file),
            shutdown_file.stat().st_mtime,
        )
    except Exception as exc:
        shutdown_load_error = str(exc)





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
    if strategy_profit > 1e-6:
        render_kpi(
            "日前/实时交易策略收益",
            f"{strategy_profit:,.0f} 元",
            "低买高卖 / 价格优化收益",
        )
    elif strategy_profit < -1e-6:
        render_kpi(
            "日前/实时交易策略损失",
            f"{abs(strategy_profit):,.0f} 元",
            "高买低卖 / 价格损失",
        )
    else:
        render_kpi(
            "日前/实时交易策略收益",
            "0 元",
            "— 持平",
        )



# 新增成本类KPI
c1, c2 = st.columns(2, gap="medium")

with c1:
   render_kpi(
       "日前偏差费用",
       f"{deviation_fee_total:,.2f} 元",
       "交易中心结算考核费用",
   )

with c2:
   render_kpi(
       "实际结算成本",
       f"{actual_settlement_cost:,.2f} 元",
       "交易费用 + 日前偏差费用",
   )


# =========================================================
# 11. Tabs
# =========================================================
(
   tab_overview,
   tab_analysis,
   tab_deviation,
   tab_risk,
   tab_intraday,
   tab_trend,
   tab_eb,
   tab_cost,
   tab_data,
) = st.tabs(
   [
       "交易概览",
       "分析结论",
       "日前偏差费用",
       "中长期偏差预警",
       "日内量价",
       "每日趋势",
       "能量块复盘",
       "到户电费模拟",
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
   if strategy_profit > 1e-6:
       st.success(
           f"🟢 日前/实时交易策略收益：{strategy_profit:,.2f} 元"
       )
   elif strategy_profit < -1e-6:
       st.error(
           f"🔴 日前/实时交易策略损失：{abs(strategy_profit):,.2f} 元"
       )
   else:
       st.info("日前/实时交易策略基本持平。")

   st.caption(
       "策略收益按所选时段的日前、实时加权均价计算，"
       "反映低买高卖或高买低卖带来的交易收益或损失。"
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
# 新增 Tab 分析结论
# =========================================================
with tab_analysis:
   st.markdown(
       '<div class="section-title">交易分析结论</div>',
       unsafe_allow_html=True,
   )

   if pd.isna(analysis_end_date) or analysis_filtered.empty:
       st.info("当前筛选范围尚无有效实时交易数据，暂不生成分析结论。")
   else:
       analysis_start_date = pd.Timestamp(start_date).normalize()
       st.caption(
           f"分析期间：{analysis_start_date:%Y-%m-%d} 至 {pd.Timestamp(analysis_end_date):%Y-%m-%d} "
           "（截止最新有效实时交易数据日；之后仅有年协/月协或能量块的数据不进入本页结论）。"
       )

       with st.expander("⚙️ 分析判断参数", expanded=False):
           st.markdown("**负荷下降 / 回售判断**")
           a1, a2, a3 = st.columns(3)
           resale_normal_pct = a1.number_input(
               "正常调整上限（%）", min_value=0.0, max_value=100.0,
               value=DEFAULT_ANALYSIS_PARAMS["resale_normal_max"] * 100,
               step=1.0, format="%.1f", key="analysis_resale_normal"
           )
           resale_general_pct = a2.number_input(
               "一般调整上限（%）", min_value=0.0, max_value=100.0,
               value=DEFAULT_ANALYSIS_PARAMS["resale_general_max"] * 100,
               step=1.0, format="%.1f", key="analysis_resale_general"
           )
           resale_drop_pct = a3.number_input(
               "重大负荷下降起点（%）", min_value=0.0, max_value=100.0,
               value=DEFAULT_ANALYSIS_PARAMS["resale_load_drop_max"] * 100,
               step=1.0, format="%.1f", key="analysis_resale_drop"
           )

           st.markdown("**单次交易结果复盘过滤**")
           b1, b2, b3 = st.columns(3)
           strategy_gap = b1.number_input(
               "最小异常价差（元/MWh）", min_value=0.0,
               value=DEFAULT_ANALYSIS_PARAMS["strategy_price_gap"],
               step=1.0, format="%.1f", key="analysis_strategy_gap"
           )
           strategy_volume_pct = b2.number_input(
               "最小涉及电量 / LT（%）", min_value=0.0, max_value=100.0,
               value=DEFAULT_ANALYSIS_PARAMS["strategy_volume_ratio"] * 100,
               step=1.0, format="%.1f", key="analysis_strategy_volume"
           )
           strategy_impact = b3.number_input(
               "最小潜在金额影响（元）", min_value=0.0,
               value=DEFAULT_ANALYSIS_PARAMS["strategy_impact_yuan"],
               step=500.0, format="%.0f", key="analysis_strategy_impact"
           )

           st.markdown("**持续性策略复盘**")
           c1, c2, c3 = st.columns(3)
           repeat_count = c1.number_input(
               "重复事件起点（次/天）", min_value=2, max_value=20,
               value=int(DEFAULT_ANALYSIS_PARAMS["strategy_repeat_count"]),
               step=1, key="analysis_repeat_count"
           )
           repeat_ratio_pct = c2.number_input(
               "不利事件占有效日比例（%）", min_value=0.0, max_value=100.0,
               value=DEFAULT_ANALYSIS_PARAMS["strategy_repeat_ratio"] * 100,
               step=5.0, format="%.1f", key="analysis_repeat_ratio"
           )
           cumulative_impact = c3.number_input(
               "累计潜在影响警戒值（元）", min_value=0.0,
               value=DEFAULT_ANALYSIS_PARAMS["strategy_cumulative_impact_yuan"],
               step=5000.0, format="%.0f", key="analysis_cumulative_impact"
           )

       analysis_params = {
           "resale_normal_max": resale_normal_pct / 100,
           "resale_general_max": resale_general_pct / 100,
           "resale_load_drop_max": resale_drop_pct / 100,
           "strategy_price_gap": strategy_gap,
           "strategy_volume_ratio": strategy_volume_pct / 100,
           "strategy_impact_yuan": strategy_impact,
           "strategy_repeat_count": int(repeat_count),
           "strategy_repeat_ratio": repeat_ratio_pct / 100,
           "strategy_cumulative_impact_yuan": cumulative_impact,
       }

       if not (
           analysis_params["resale_normal_max"]
           < analysis_params["resale_general_max"]
           < analysis_params["resale_load_drop_max"]
       ):
           st.error("回售率参数需满足：正常调整上限 < 一般调整上限 < 重大负荷下降起点。")
       else:
           analysis_detail = build_analysis_dataset(
               analysis_filtered, amount_col, analysis_params
           )

           # -----------------------------------------------------
           # 1) 已知停机记录读取、分钟级拆分与交易归因
           # -----------------------------------------------------
           shutdown_split = pd.DataFrame()
           shutdown_period_summary = pd.DataFrame()
           shutdown_event_detail = pd.DataFrame()
           shutdown_period_impact = pd.DataFrame()

           if shutdown_load_error:
               st.warning(f"停机历史文件读取失败：{shutdown_load_error}")
           elif shutdown_file is None:
               st.info(
                   f"尚未检测到 {SHUTDOWN_FILE_NAME}。本页仍会保留交易数据自诊断；"
                   "将停机模板填写后放入data目录，即可自动增加已知停机归因。"
               )
           elif shutdown_raw.empty:
               st.info("停机历史文件已读取，但当前没有有效PM1/PM2/PM3停机记录。")
           else:
               shutdown_split = split_shutdown_to_periods(
                   shutdown_raw,
                   analysis_start_date,
                   analysis_end_date,
               )
               shutdown_period_summary = build_shutdown_period_summary(
                   shutdown_split,
                   analysis_detail,
               )
               shutdown_event_detail = build_shutdown_event_detail(
                   shutdown_raw,
                   shutdown_split,
                   shutdown_period_summary,
               )
               shutdown_period_impact = calculate_shutdown_period_impact(
                   analysis_filtered,
                   amount_col,
                   shutdown_period_summary,
               )

           # -----------------------------------------------------
           # 2) 策略事件：先识别，再排除已知停机时段
           # -----------------------------------------------------
           strategy_events = build_strategy_review_events(
               analysis_filtered,
               analysis_params,
           )

           if not strategy_events.empty and not shutdown_period_summary.empty:
               shutdown_keys = set(
                   (pd.Timestamp(r["日期"]).normalize(), str(r["时段"]))
                   for _, r in shutdown_period_summary.iterrows()
                   if float(r.get("理论停机负荷下降_MWh", 0) or 0) > 1e-9
               )
               strategy_events = strategy_events[
                   ~strategy_events.apply(
                       lambda r: (pd.Timestamp(r["日期"]).normalize(), str(r["时段"])) in shutdown_keys,
                       axis=1,
                   )
               ].copy()

           # 没有停机记录时，继续沿用交易数据自诊断：明显负荷下降时段不评价策略。
           if not strategy_events.empty and shutdown_period_summary.empty and not analysis_detail.empty:
               load_drop_keys = set(
                   (pd.Timestamp(r["日期"]).normalize(), str(r["时段"]))
                   for _, r in analysis_detail[
                       analysis_detail["负荷下降信号率"].ge(
                           analysis_params["resale_general_max"]
                       )
                   ].iterrows()
               )
               strategy_events = strategy_events[
                   ~strategy_events.apply(
                       lambda r: (pd.Timestamp(r["日期"]).normalize(), str(r["时段"])) in load_drop_keys,
                       axis=1,
                   )
               ].copy()

           strategy_persistence = summarize_strategy_persistence(
               strategy_events,
               analysis_filtered,
               analysis_params,
           )

           # -----------------------------------------------------
           # 3) 中长期覆盖率背景：用于约束策略结论措辞
           # -----------------------------------------------------
           analysis_final_energy = float(analysis_filtered["energy_mwh"].sum())
           analysis_midlong_energy = float(
               analysis_filtered.loc[
                   analysis_filtered["market_stage"].isin(["long_term", "energy_block"]),
                   "energy_mwh",
               ].sum()
           )
           analysis_coverage = (
               analysis_midlong_energy / analysis_final_energy
               if analysis_final_energy > 1e-9 else np.nan
           )

           # =====================================================
           # 管理层摘要：先回答“发生了什么、哪些天、多少钱、是否持续”
           # =====================================================
           st.markdown("### 月度分析摘要")
           period_prices_analysis = calculate_period_weighted_prices(
               analysis_filtered, amount_col
           )

           summary_lines = []
           valid_prices = {
               k: v for k, v in period_prices_analysis.items() if pd.notna(v)
           }
           if valid_prices:
               highest_period = max(valid_prices, key=valid_prices.get)
               lowest_period = min(valid_prices, key=valid_prices.get)
               spread = valid_prices[highest_period] - valid_prices[lowest_period]
               summary_lines.append(
                   f"分析期内{highest_period}段综合成交加权均价最高，为 "
                   f"**{valid_prices[highest_period]:,.2f} 元/MWh**；"
                   f"较{lowest_period}段高 **{spread:,.2f} 元/MWh**。"
                   "该价差本身不直接定义为异常，需结合生产停机和各市场交易结构解释。"
               )

           # 停机影响摘要
           if not shutdown_period_summary.empty:
               total_shutdown_theoretical = float(
                   shutdown_period_summary["理论停机负荷下降_MWh"].sum()
               )
               total_shutdown_cost = float(
                   shutdown_period_summary["停机均价影响折算_元"].sum()
               )
               affected_days = shutdown_period_summary["日期"].nunique()
               summary_lines.append(
                   f"已知纸机停机涉及 **{affected_days} 个用电日**，按PM1/PM2/PM3典型负荷估算，"
                   f"对应理论负荷下降约 **{total_shutdown_theoretical:,.0f} MWh**；"
                   f"不区分计划/非计划，按所有已知停机的无停机反事实情景折算，停机对剩余用电均价的金额尺度影响合计约 "
                   f"**{total_shutdown_cost:+,.0f} 元**。该金额用于表达均价影响尺度，不等同于两种情景的总电费差额。"
               )

               if not shutdown_period_impact.empty:
                   impact_abs = shutdown_period_impact.copy()
                   impact_abs["abs_uplift"] = impact_abs["停机对均价影响估算_元每MWh"].abs()
                   top_impact = impact_abs.sort_values("abs_uplift", ascending=False).iloc[0]
                   if pd.notna(top_impact["停机对均价影响估算_元每MWh"]):
                       summary_lines.append(
                           f"停机对峰平谷均价影响中，以**{top_impact['时段']}段**最明显："
                           f"估算影响 **{top_impact['停机对均价影响估算_元每MWh']:+.2f} 元/MWh**，"
                           f"实际均价 {top_impact['实际综合均价_元每MWh']:.2f} 元/MWh，"
                           f"无停机反事实估算均价约 "
                           f"{top_impact['无停机估算综合均价_元每MWh']:.2f} 元/MWh。"
                       )

               top_shutdown = shutdown_period_summary.sort_values(
                   "停机均价影响折算_元", ascending=False
               ).head(3)
               if not top_shutdown.empty:
                   labels = [
                       f"{pd.Timestamp(r['日期']):%m月%d日}{r['时段']}段({r['停机机组']})"
                       for _, r in top_shutdown.iterrows()
                   ]
                   summary_lines.append(
                       "对停机相关价格影响贡献较大的日期/时段包括：**"
                       + "、".join(labels) + "**。详细交易路径见下方。"
                   )
           else:
               # 无停机表时保留旧版交易自诊断价值
               auto_events = analysis_detail[
                   analysis_detail["负荷下降信号率"].ge(
                       analysis_params["resale_general_max"]
                   )
               ].copy()
               if not auto_events.empty:
                   top_auto = auto_events.sort_values(
                       "负荷下降信号率", ascending=False
                   ).head(3)
                   labels = [
                       f"{pd.Timestamp(r['日期']):%m月%d日}{r['时段']}段"
                       for _, r in top_auto.iterrows()
                   ]
                   summary_lines.append(
                       "当前未接入已知停机记录，但交易链识别到明显负荷下降特征，"
                       "重点日期/时段包括 **" + "、".join(labels) + "**。"
                   )

           # 策略连续性摘要：偶发不下“策略失误”结论
           persistent_rows = (
               strategy_persistence[
                   strategy_persistence["判断"].eq("持续性策略复盘警告")
               ] if not strategy_persistence.empty else pd.DataFrame()
           )
           if not persistent_rows.empty:
               total_persistent = float(persistent_rows["累计潜在影响_元"].sum())
               main = persistent_rows.sort_values("累计潜在影响_元", ascending=False).iloc[0]
               summary_lines.append(
                   f"剔除已知停机时段及小电量/小金额结果后，出现**持续性不利交易结果**："
                   f"{main['时段']}段“{main['复盘类型']}”共 {int(main['不利事件天数'])} 个交易日，"
                   f"最长连续 {int(main['最长连续天数'])} 天；持续性事件累计潜在影响约 "
                   f"**{total_persistent:,.0f} 元**。建议复盘负荷预测、价格预测和提前采购比例，"
                   "但不直接认定为交易策略失误。"
               )
           elif not strategy_events.empty:
               summary_lines.append(
                   f"剔除已知停机/明显负荷下降后，共有 **{len(strategy_events)} 个**单次交易结果达到复盘阈值，"
                   "但当前未形成持续性模式，因此归为**偶发不利结果**，不足以判断存在系统性策略问题。"
               )
           else:
               summary_lines.append(
                   "剔除已知停机/明显负荷下降和小电量、小金额波动后，当前没有识别到需要持续性警告的交易时点问题。"
               )

           if pd.notna(analysis_coverage):
               summary_lines.append(
                   f"分析期中长期净电量（年协/月协+能量块）覆盖率约 **{analysis_coverage*100:.1f}%**。"
                   "交易决策同时受70%–105%中长期合同覆盖约束，因此不能仅根据事后更低的现货价格判断前序采购是否合理。"
               )

           for i, line in enumerate(summary_lines, 1):
               st.markdown(f"**{i}.** {line}")

           # =====================================================
           # 已知停机影响
           # =====================================================
           st.markdown("### 已知纸机停机影响")
           st.caption(
               "PM典型运行负荷默认：PM1≈31 MW、PM2≈28 MW、PM3≈23 MW；"
               "停机Excel可用load_mw_override为单次记录覆盖默认值，并可填写shutdown_type=计划/非计划/不确定。"
               " 均价影响计算不区分计划或非计划：只要该时段实际发生纸机停机，就统一作为生产负荷下降事实进行估算；"
               " shutdown_type仅用于事件解释和后续复盘。"
               "停机开始/结束支持精确到分钟和跨天；stop_end建议填写大致恢复正常生产负荷的时间，"
               "代码按实际重叠分钟拆入峰/平/谷。"
           )

           if shutdown_file is not None and not shutdown_load_error:
               st.caption(f"停机记录来源：{shutdown_file.name}")

           if shutdown_period_impact.empty:
               st.info("当前分析期没有可用的已知停机归因数据。")
           else:
               imp = shutdown_period_impact.rename(columns={
                   "实际综合均价_元每MWh": "实际综合均价(元/MWh)",
                   "实际最终净电量_MWh": "实际最终净电量(MWh)",
                   "无停机恢复电量估算_MWh": "无停机恢复电量估算(MWh)",
                   "无停机反事实恢复成本_元": "无停机恢复交易成本估算(元)",
                   "无停机估算综合均价_元每MWh": "无停机估算综合均价(元/MWh)",
                   "停机对均价影响估算_元每MWh": "停机对均价影响估算(元/MWh)",
                   "停机均价影响折算_元": "停机均价影响折算(元)",
               })
               st.dataframe(imp.round(2), use_container_width=True, hide_index=True)
               st.caption(
                   "“无停机估算均价”采用完整反事实，且不区分计划/非计划停机：实际交易结果基础上先撤销可由停机解释的回售，"
                   "若理论正常需求仍有缺口，再按负荷变化主要被交易侧反映阶段的市场价格估算补购。"
                   "因此它同时考虑了正常运行时可能仍需在后续市场补买的电量；结果用于运营复盘，不替代正式结算。"
               )

           st.markdown("#### 停机记录与理论负荷下降")
           if shutdown_event_detail.empty:
               st.info("暂无当前分析期内的有效停机记录。")
           else:
               show_stop = shutdown_event_detail.copy()
               show_stop["停机开始"] = pd.to_datetime(show_stop["停机开始"]).dt.strftime("%Y-%m-%d %H:%M")
               show_stop["停机结束"] = pd.to_datetime(show_stop["停机结束"]).dt.strftime("%Y-%m-%d %H:%M")
               st.dataframe(show_stop.round(2), use_container_width=True, hide_index=True)

           st.markdown("#### 停机日期 / 时段交易归因")
           if shutdown_period_summary.empty:
               st.info("暂无可与交易数据对齐的停机时段。")
           else:
               shutdown_sorted = shutdown_period_summary.sort_values(
                   ["日期", "时段"]
               )
               for _, r in shutdown_sorted.iterrows():
                   explain_rate = (
                       r["生产事件解释率"] * 100
                       if pd.notna(r.get("生产事件解释率", np.nan)) else np.nan
                   )
                   title = (
                       f"{pd.Timestamp(r['日期']):%Y-%m-%d}｜{r['时段']}段｜"
                       f"{r['停机机组']}｜{r.get('停机类型', '不确定')}｜"
                       f"{r.get('负荷下降发现阶段', '交易阶段未知')}"
                   )
                   with st.expander(title, expanded=True):
                       st.markdown(
                           f"**理论停机负荷下降：** {r['理论停机负荷下降_MWh']:,.1f} MWh  "
                           f"  \n**观察到的各阶段卖出：** {r['观察到的阶段卖出_MWh']:,.1f} MWh  "
                           f"  \n**可由已知停机解释的卖出：** {r['停机可解释卖出_MWh']:,.1f} MWh"
                       )
                       if pd.notna(explain_rate):
                           st.markdown(f"**生产事件解释率：** {explain_rate:.1f}%")
                       if pd.notna(r.get("交易路径", np.nan)):
                           st.markdown(f"**交易路径：** {r['交易路径']}")
                           st.markdown(
                               f"**阶段净调整：** 能量块 {r['能量块调整_MWh']:+,.1f} MWh；"
                               f"日前 {r['日前调整_MWh']:+,.1f} MWh；"
                               f"实时 {r['实时调整_MWh']:+,.1f} MWh。"
                           )
                       if pd.notna(r.get("负荷下降发现阶段", np.nan)):
                           st.markdown(
                               f"**交易侧反映时点：** {r['负荷下降发现阶段']}。"
                               f"{r.get('可能原因解释', '')}"
                           )
                       if pd.notna(r.get("无停机估算综合均价_元每MWh", np.nan)):
                           delta_p = float(r.get("停机对该日时段均价影响_元每MWh", 0) or 0)
                           st.caption(
                               f"该日该时段实际综合均价 {r.get('最终综合均价', np.nan):,.2f} 元/MWh；"
                               f"无停机反事实估算约 {r.get('无停机估算综合均价_元每MWh', np.nan):,.2f} 元/MWh；"
                               f"停机对单位均价影响约 {delta_p:+.2f} 元/MWh。"
                           )
                       if "计划" in str(r.get("停机类型", "")) and "非计划" not in str(r.get("停机类型", "")) and "实时" in str(r.get("负荷下降发现阶段", "")):
                           st.warning(
                               "该记录标记为计划停机，但主要负荷下修直到实时阶段才明显反映。"
                               "建议复盘生产计划信息传递、负荷预测更新和交易头寸调整时点；这不直接等同于交易策略失误。"
                           )

           # =====================================================
           # 无停机记录/停机无法完全解释的负荷异常仍保留
           # =====================================================
           st.markdown("### 未解释的负荷下降 / 异常交易链")
           auto_events = analysis_detail[
               analysis_detail["负荷下降信号率"].ge(
                   analysis_params["resale_general_max"]
               )
           ].copy()
           if not shutdown_period_summary.empty and not auto_events.empty:
               known_keys = set(
                   (pd.Timestamp(r["日期"]).normalize(), str(r["时段"]))
                   for _, r in shutdown_period_summary.iterrows()
               )
               auto_events = auto_events[
                   ~auto_events.apply(
                       lambda r: (pd.Timestamp(r["日期"]).normalize(), str(r["时段"])) in known_keys,
                       axis=1,
                   )
               ].copy()

           if auto_events.empty:
               st.success("当前阈值下，没有已知停机之外需要重点解释的明显负荷下降时段。")
           else:
               for _, r in auto_events.sort_values("负荷下降信号率", ascending=False).iterrows():
                   with st.expander(
                       f"{pd.Timestamp(r['日期']):%Y-%m-%d}｜{r['时段']}段｜"
                       f"{r['负荷下降判断']}｜{r['负荷下降发现阶段']}",
                       expanded=False,
                   ):
                       st.markdown(f"**交易路径：** {r['交易路径']}")
                       st.markdown(f"**可能原因：** {r['可能原因解释']}")
                       st.caption(
                           "该时段没有被当前停机历史解释，可能来自其他生产负荷变化、预测偏差或尚未记录的停机事件，"
                           "不直接归为交易策略问题。"
                       )

           # =====================================================
           # 策略复盘：偶发 vs 持续性
           # =====================================================
           st.markdown("### 交易策略结果复盘")
           st.caption(
               "一般市场策略复盘只评价已知停机/明显负荷下降之外的交易结果；shutdown_type只作为事件背景。若已知停机直到较晚市场才反映，"
               "会在上方生产事件归因中单独提示生产—预测—交易协同复盘。单次达到阈值仅称为‘偶发不利结果’，"
               "不会写成交易策略失误；只有同类型结果重复、连续或累计金额明显时才触发持续性复盘警告。"
           )
           st.caption(
               "这些结果是事后价格比较。交易当时依赖负荷预测和价格预测，同时受70%–105%中长期合同覆盖约束；"
               "持续不利结果可能来自负荷预测偏差、价格预测偏差、提前采购比例或交易时点，需要结合当日预测信息复盘。"
           )

           if strategy_persistence.empty:
               st.success("当前阈值下没有需要展示的交易时点复盘事件。")
           else:
               persist_show = strategy_persistence.copy()
               persist_show["发生比例(%)"] = persist_show["发生比例"] * 100
               persist_show = persist_show.drop(columns=["发生比例"])
               st.dataframe(persist_show.round(2), use_container_width=True, hide_index=True)

               for _, r in strategy_persistence.iterrows():
                   if r["判断"] == "持续性策略复盘警告":
                       st.warning(
                           f"{r['时段']}段｜{r['复盘类型']}：分析期内出现 {int(r['不利事件天数'])} 个交易日，"
                           f"最长连续 {int(r['最长连续天数'])} 天，累计潜在影响约 {r['累计潜在影响_元']:,.0f} 元。"
                           "该模式值得重点复盘，但仍不能仅凭事后价格认定策略失误。"
                       )

           st.markdown("#### 单次复盘事件明细")
           if strategy_events.empty:
               st.info("没有达到当前价差、电量和金额阈值的单次复盘事件。")
           else:
               se = strategy_events.copy()
               se["日期"] = pd.to_datetime(se["日期"]).dt.strftime("%Y-%m-%d")
               se["占LT基础量(%)"] = se["占LT基础量"] * 100
               display_se = se[[
                   "日期", "日类型", "时段", "复盘类型", "前序市场",
                   "买入量_MWh", "占LT基础量(%)", "前序买入均价",
                   "后续市场参考均价", "价差_元每MWh", "潜在影响_元",
               ]].rename(columns={
                   "买入量_MWh": "买入量(MWh)",
                   "前序买入均价": "前序买入均价(元/MWh)",
                   "后续市场参考均价": "后续市场参考均价(元/MWh)",
                   "价差_元每MWh": "价差(元/MWh)",
                   "潜在影响_元": "潜在影响(元)",
               })
               st.dataframe(display_se.round(2), use_container_width=True, hide_index=True)

           # =====================================================
           # 算法说明：保留原回售逻辑，并明确新停机归因边界
           # =====================================================
           with st.expander("查看分析口径与边界", expanded=False):
               st.markdown(
                   f"""
**1. 基础电量回售率**  
`max[-(能量块 + 日前 + 实时), 0] ÷ 年协/月协基础电量`。  
默认阈值：<{resale_normal_pct:.0f}% 正常；{resale_normal_pct:.0f}%–{resale_general_pct:.0f}% 一般调整；
{resale_general_pct:.0f}%–{resale_drop_pct:.0f}% 明显负荷下降；≥{resale_drop_pct:.0f}% 重大负荷下降。

**2. 已知停机理论负荷下降**  
每条停机记录按实际重叠分钟计算：`纸机典型MW × 停机分钟 ÷ 60`，并自动拆分到峰/平/谷。  
默认PM1=31 MW、PM2=28 MW、PM3=23 MW，可在Excel用`load_mw_override`覆盖；`shutdown_type`可填计划/非计划/不确定，**但该字段不参与均价影响计算**，只用于事件说明和复盘。`stop_end`建议填写大致恢复正常负荷的时点。

**3. 停机可解释卖出量**  
取`理论停机负荷下降`与`能量块/日前/实时各阶段实际负调整之和`的较小值。这样既不把超过理论停机负荷的卖出全部归因于停机，也不会因为前序买入抵消实时大卖而漏判。

**4. 无停机反事实均价**  
这里不区分停机属于计划、非计划还是不确定，只依据“该时段实际发生停机”这一事实估算生产影响。不是简单计算`LT价−回售价`。先从实际交易结果出发：①把可由停机解释的实际卖出按其真实回售价撤销；②若纸机正常运行需求仍高于已撤销的卖出量，则把剩余需求按负荷下降主要被反映的交易阶段（能量块/日前/实时）的价格估算补购；③得到无停机估算净电量、总成本和综合均价。  
`停机对均价影响 = 实际综合均价 − 无停机估算综合均价`。正值表示停机使剩余用电的单位综合成本上升。由于两种情景用电量不同，页面中的“均价影响折算金额”仅用于表达影响尺度，不等同于总电费差额。

**5. 策略复盘**  
先排除已知停机时段；没有停机表时排除交易链已识别的明显负荷下降时段。单次事件只有在价差≥{strategy_gap:.1f}元/MWh、涉及电量占LT≥{strategy_volume_pct:.1f}%且潜在金额≥{strategy_impact:,.0f}元时才展示。  
单次结果只称为“偶发不利结果”。达到重复次数、连续性/发生比例或累计金额条件后，才提示“持续性策略复盘警告”。

**6. 重要边界**  
本页是运营与交易复盘工具，不是事后最优交易证明。交易当时受负荷预测、价格预测、生产信息可得时点以及70%–105%中长期合同覆盖约束；因此所有策略提示均使用“值得复盘 / 可能存在预测偏差”等措辞，不直接认定人为失误。
"""
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


   add_period_shading(fig_qty, selected_day, resolution)
   st.plotly_chart(fig_qty, use_container_width=True)


   # 价格图：所有市场采用散点展示，避免连续线误导交易状态
   # 颜色 = 市场阶段，点形 = 买入/卖出方向
   import plotly.graph_objects as go


   fig_price = go.Figure()


   for market in market_selection:
       tmp = plot_df[
           plot_df["market_stage_cn"] == market
       ].copy()


       if tmp.empty:
           continue


       # 无成交不显示
       tmp = tmp[tmp["energy_mwh"].abs() > 0].copy()
       if tmp.empty:
           continue


       tmp["direction"] = np.where(
           tmp["energy_mwh"] >= 0,
           "买入",
           "卖出",
       )


       # 价格图使用点形区分交易方向：
       # ▲ 买入，▼ 卖出；颜色仍表示市场阶段


       tmp["marker_symbol"] = np.where(
           tmp["direction"] == "买入",
           "triangle-up",
           "triangle-down",
       )


       fig_price.add_trace(
           go.Scatter(
               x=tmp[x_col],
               y=tmp[price_plot_col],
               mode="markers",
               name=market,
               marker=dict(
                   size=10,
                   symbol=tmp["marker_symbol"],
               ),
               customdata=np.stack(
                   [
                       tmp["market_stage_cn"],
                       tmp["direction"],
                       tmp["energy_mwh"].abs(),
                   ],
                   axis=-1,
               ),
               hovertemplate=(
                   "时间：%{x}<br>"
                   "市场：%{customdata[0]}<br>"
                   "方向：%{customdata[1]}<br>"
                   "成交电量：%{customdata[2]:,.2f} MWh<br>"
                   "成交价格：%{y:,.2f} 元/MWh"
                   "<extra></extra>"
               ),
           )
       )


   fig_price.update_layout(
       title=f"{selected_day} 各阶段交易价格（▲买入 ▼卖出）",
       yaxis_title="元/MWh",
       annotations=[
           dict(
               text="▲ 买入，▼ 卖出；颜色表示市场阶段",
               xref="paper",
               yref="paper",
               x=0,
               y=1.08,
               showarrow=False,
               font=dict(size=12),
           )
       ],
       height=430,
       hovermode="closest",
       xaxis_title=None,
   )


   add_period_shading(fig_price, selected_day, resolution)
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
           legend_title_text="市场阶段",
       )


       st.plotly_chart(
           fig_daily_qty,
           use_container_width=True,
       )


   with col2:
       daily_price = daily.copy()


       # 未更新市场/未成交阶段不显示，不用0代表缺失
       daily_price.loc[
           daily_price["avg_price"] == 0,
           "avg_price",
       ] = np.nan


       fig_daily_price = px.scatter(
           daily_price,
           x="display_date",
           y="avg_price",
           color="market_stage_cn",
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
           legend_title_text="市场阶段",
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


   # 每日最终综合净价：
   # 总交易金额 / 最终净交易电量
   daily_amount_structure = (
       daily.pivot_table(
           index="display_date",
           columns="market_stage_cn",
           values="amount_selected",
           aggfunc="sum",
           fill_value=0,
       )
       .reset_index()
   )


   for market in MARKET_ORDER:
       if market not in daily_amount_structure.columns:
           daily_amount_structure[market] = 0


   daily_structure["最终净金额"] = (
       daily_amount_structure["年协/月协"]
       + daily_amount_structure["能量块"]
       + daily_amount_structure["日前"]
       + daily_amount_structure["实时"]
   )


   daily_structure["最终净价"] = np.where(
       daily_structure["最终净交易量"].abs() > 1e-9,
       daily_structure["最终净金额"] / daily_structure["最终净交易量"],
       np.nan,
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


   st.markdown("#### 每日最终净价趋势")


   fig_final_price = px.scatter(
       daily_structure,
       x="display_date",
       y="最终净价",
   )


   fig_final_price.update_traces(
       hovertemplate=(
           "日期：%{x|%Y-%m-%d}<br>"
           "最终净金额："
           f"%{{customdata[0]:,.2f}} 元<br>"
           "最终净量："
           f"%{{customdata[1]:,.2f}} MWh<br>"
           "最终净价："
           f"%{{y:,.2f}} 元/MWh"
           "<extra></extra>"
       ),
       customdata=np.column_stack(
           [
               daily_structure["最终净金额"],
               daily_structure["最终净交易量"],
           ]
       ),
   )


   fig_final_price.update_xaxes(tickformat="%m-%d")
   fig_final_price.update_layout(
       title="每日最终综合成交净价",
       xaxis_title="日期",
       yaxis_title="元/MWh",
       height=410,
   )


   st.plotly_chart(
       fig_final_price,
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
           legend_title_text="市场阶段",
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
# Tab 到户电费模拟
# =========================================================
with tab_cost:


   st.markdown(
       '<div class="section-title">到户电费模拟</div>',
       unsafe_allow_html=True,
   )


   with st.expander("⚙️ 电费计算参数设置", expanded=False):


       peak_up = st.number_input(
           "峰上浮比例",
           value=0.80,
           step=0.01,
       )


       valley_down = st.number_input(
           "谷下降比例",
           value=0.65,
           step=0.01,
       )


       line_loss = st.number_input(
           "线损率",
           value=0.0318,
           step=0.001,
           format="%.4f",
       )


       transmission = st.number_input(
           "输配电价 元/kWh",
           value=0.0857,
           step=0.0001,
           format="%.4f",
       )


       fund = st.number_input(
           "基金附加 元/kWh",
           value=0.0294,
           step=0.0001,
           format="%.4f",
       )


       system_detail = {}
       st.markdown("系统运行费用")
       for k, v in DEFAULT_SYSTEM_FEE_DETAIL.items():
           system_detail[k] = st.number_input(
               k,
               value=v,
               step=0.0001,
               format="%.4f",
           )


   params = {
       "peak_up_ratio": peak_up,
       "valley_down_ratio": valley_down,
       "line_loss_rate": line_loss,
       "transmission_price": transmission,
       "fund_surcharge": fund,
       "system_fee": sum(system_detail.values()),
   }


   # 固定使用电能量价格，不含环境权益
   ppg = calculate_period_weighted_prices(
       filtered,
       "energy_amount_yuan",
   )


   delivery = calc_delivery_price(
       ppg.get("平", np.nan),
       params,
   )


   if delivery:


       fpg = filtered.copy()
       fpg["period_type_tmp"] = fpg["datetime_bj"].apply(period_type)


       volumes = (
           fpg.groupby("period_type_tmp")["energy_mwh"]
           .sum()
           .to_dict()
       )


       peak_fee = abs(volumes.get("峰", 0) * delivery["peak"] * 1000)
       flat_fee = abs(volumes.get("平", 0) * delivery["flat"] * 1000)
       valley_fee = abs(volumes.get("谷", 0) * delivery["valley"] * 1000)


       energy_fee = peak_fee + flat_fee + valley_fee


       env_fee = 0
       if "amount_yuan" in fpg.columns and "energy_amount_yuan" in fpg.columns:
           env_fee = (
               fpg["amount_yuan"].sum()
               - fpg["energy_amount_yuan"].sum()
           )
       elif "environmental_value_amount" in fpg.columns:
           env_fee = fpg["environmental_value_amount"].sum()


       final_cost = energy_fee + env_fee


       c1, c2, c3, c4 = st.columns(4)


       c1.metric("峰段电费", f"{peak_fee:,.2f} 元")
       c2.metric("平段电费", f"{flat_fee:,.2f} 元")
       c3.metric("谷段电费", f"{valley_fee:,.2f} 元")
       c4.metric("最终预计成本", f"{final_cost:,.2f} 元")


       st.markdown("#### 计算拆解")


       st.dataframe(
           pd.DataFrame([
               ["平段市场交易基准价", ppg.get("平", np.nan), "元/MWh"],
               ["峰到户电价", delivery["peak"] * 1000, "元/MWh"],
               ["平到户电价", delivery["flat"] * 1000, "元/MWh"],
               ["谷到户电价", delivery["valley"] * 1000, "元/MWh"],
               ["电能量费用", energy_fee, "元"],
               ["环境权益费用", env_fee, "元"],
               ["最终预计成本", final_cost, "元"],
           ], columns=["项目", "数值", "单位"]),
           use_container_width=True,
           hide_index=True,
       )


       with st.expander("查看电费计算算法"):
           st.markdown(
               f"""
**1. 市场交易购电价格**


平段电能量成交加权均价：
{ppg.get("平", np.nan):.2f} 元/MWh


**2. 到户电价计算**


峰：
市场价格 × (1 + 峰上浮比例) + 固定费用


平：
市场价格 + 固定费用


谷：
市场价格 × (1 - 谷下降比例) + 固定费用


**3. 固定费用**


线损 + 输配电价 + 系统运行费用 + 基金附加


**4. 最终成本**


+ 峰段电量 × 峰价
+ 平段电量 × 平价
+ 谷段电量 × 谷价
+ 环境权益费用
"""
           )


       st.caption("暂未考虑深谷尖峰变化。")
   else:
       st.warning("缺少平段电能量成交价格，无法计算到户电费。")






# =========================================================
# Tab 2 日前偏差费用
# =========================================================
with tab_deviation:
   st.markdown(
       '<div class="section-title">日前偏差费用</div>',
       unsafe_allow_html=True,
   )

   st.caption(
       "来源于交易中心结算单“偏差收益回收”字段，直接读取电费结果，不重新计算偏差规则。"
   )

   d1, d2 = st.columns(2)
   with d1:
       render_kpi(
           "日前偏差费用总额",
           f"{deviation_fee_total:,.2f} 元",
           "增加成本",
       )
   with d2:
       render_kpi(
           "对应偏差电量",
           f"{deviation_energy_total:,.2f} MWh",
           "",
       )

   if not deviation_filtered.empty:
       st.dataframe(
           deviation_filtered,
           use_container_width=True,
           hide_index=True,
       )
   else:
       st.info("当前筛选范围无日前偏差费用数据。")


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
