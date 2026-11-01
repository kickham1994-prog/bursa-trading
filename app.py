# -*- coding: utf-8 -*-
"""
====================================================================
Bursa Dual-System Quant Dashboard (Streamlit 网页控制台)
- 查看最新自动生成的 Excel 报告
- 实时交互：筛选买入候选、查看持仓风险、浏览 Signal Memory
====================================================================
"""

import os
import glob
import pandas as pd
import streamlit as st

# 页面配置
st.set_page_config(
    page_title="Bursa 自动化双系统交易控制台",
    page_icon="📈",
    layout="wide"
)

st.title("🇲🇾 Bursa 自动化双系统交易控制台")
st.caption("Automated Dual-System Trading Workflow Dashboard | 主系统: EMA V4.1 + 辅助系统: Pre-Trade Checker 2.3")

# 路径配置
REPORTS_DIR = "Reports/Daily_Report"
SIGNAL_MEMORY_PATH = "History/Signal_Memory.xlsx"
HOLD_LIST_PATH = "Hold list.xlsx"

# 侧边栏导航
st.sidebar.header("📂 报告与数据导航")
menu = st.sidebar.radio(
    "选择功能模块",
    ["📊 最新每日交易报告", "🛡️ 持仓风险监控 (Hold List)", "⏳ 信号记忆追踪 (Signal Memory)"]
)

# --------------------------------------------------------------------
# 模块一：最新每日交易报告
# --------------------------------------------------------------------
if menu == "📊 最新每日交易报告":
    st.header("📊 每日自动化扫描综合报告")
    
    # 查找 Reports/Daily_Report/ 下最新的 Excel 文件
    excel_files = sorted(glob.glob(os.path.join(REPORTS_DIR, "*.xlsx")), reverse=True)
    
    if not excel_files:
        st.warning(f"⚠️ 在 `{REPORTS_DIR}` 目录下暂未找到生成的 Excel 报告。请先运行一次 `daily_master_pipeline.py`。")
    else:
        # 下拉菜单选择报告日期
        selected_file = st.selectbox("选择要查看的报告文件：", excel_files)
        
        try:
            # 读取 Excel 的各个 Sheet
            xls = pd.ExcelFile(selected_file)
            sheet_names = xls.sheet_names
            
            st.success(f"✅ 成功加载报告：{os.path.basename(selected_file)}")
            
            # 选项卡展示不同 Sheet
            tabs = st.tabs(sheet_names)
            for idx, sheet in enumerate(sheet_names):
                with tabs[idx]:
                    st.subheader(f"📑 {sheet}")
                    df_sheet = pd.read_excel(selected_file, sheet_name=sheet)
                    if df_sheet.empty:
                        st.info("该 Sheet 暂无数据。")
                    else:
                        st.dataframe(df_sheet, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"❌ 读取报告出错：{e}")

# --------------------------------------------------------------------
# 模块二：持仓风险监控 (Hold List)
# --------------------------------------------------------------------
elif menu == "🛡️ 持仓风险监控 (Hold List)":
    st.header("🛡️ 持仓风险与卖出预警管理")
    
    if os.path.exists(HOLD_LIST_PATH):
        df_hold_config = pd.read_excel(HOLD_LIST_PATH)
        st.subheader("📋 当前持仓清单配置 (Hold list.xlsx)")
        st.dataframe(df_hold_config, use_container_width=True, hide_index=True)
        
        # 联动展示最新报告里的持仓风险 Sheet（如果存在）
        excel_files = sorted(glob.glob(os.path.join(REPORTS_DIR, "*.xlsx")), reverse=True)
        if excel_files:
            latest_file = excel_files[0]
            try:
                xls = pd.ExcelFile(latest_file)
                if "④持仓风险与卖出预警" in xls.sheet_names:
                    st.divider()
                    st.subheader("🔥 最新持仓风控扫描结果")
                    df_hold_risk = pd.read_excel(latest_file, sheet_name="④持仓风险与卖出预警")
                    st.dataframe(df_hold_risk, use_container_width=True, hide_index=True)
            except Exception:
                pass
    else:
        st.warning(f"⚠️ 未找到持仓文件: `{HOLD_LIST_PATH}`")

# --------------------------------------------------------------------
# 模块三：信号记忆追踪 (Signal Memory)
# --------------------------------------------------------------------
elif menu == "⏳ 信号记忆追踪 (Signal Memory)":
    st.header("⏳ Setup 信号历史记忆与演变 (Signal Memory)")
    
    if os.path.exists(SIGNAL_MEMORY_PATH):
        try:
            df_memory = pd.read_excel(SIGNAL_MEMORY_PATH)
            st.markdown("通过 Signal Memory 追踪潜在买入机会（SETUP -> TRIGGER -> BUY），防止错过早期机会。")
            
            # 交互筛选器
            if "代码" in df_memory.columns:
                selected_code = st.selectbox("按股票代码过滤：", ["全部"] + list(df_memory["代码"].unique()))
                if selected_code != "全部":
                    df_memory = df_memory[df_memory["代码"] == selected_code]
            
            st.dataframe(df_memory, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"❌ 读取 Signal Memory 失败：{e}")
    else:
        st.info("📭 暂无 Signal Memory 历史数据。随着每天自动化扫描运行，这里将自动积累记录。")