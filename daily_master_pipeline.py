# -*- coding: utf-8 -*-
"""
====================================================================
自动化双系统交易总控中心 (Daily Master Pipeline)
结合：
- EMA Trading System (主系统：趋势、阶段、Signal Memory、买卖时机)
- Pre-Trade Checker (辅助系统：交易前风险检查)
- 持仓自动风险监控 (Hold List 风险、ATR 追踪、OBV 资金流向)
- 智能量化排序与多维评分排名 (Smart Scoring & Ranking)
====================================================================
"""

import os
import sys
import importlib.util
import pandas as pd
import numpy as np
from datetime import datetime

# 引入假期库（如未安装请在终端执行: pip install holidays）
try:
    import holidays
except ImportError:
    import subprocess
    subprocess.check_call(["pip", "install", "holidays"])
    import holidays

# ==================== 动态加载核心模块 ====================
def load_module_safely(file_name, module_alias):
    if not os.path.exists(file_name):
        raise FileNotFoundError(f"找不到核心文件: {file_name}，请确保它与 daily_master_pipeline.py 在同一个目录下。")
    spec = importlib.util.spec_from_file_location(module_alias, file_name)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[module_alias] = mod
    spec.loader.exec_module(mod)
    return mod

ema_filename = "ema_trading_system_v4_0_2a_with_master_prompt_2.0.py"
pre_filename = "Pre_Trade_Checker_2.3.py"

if not os.path.exists(ema_filename):
    for f in os.listdir("."):
        if f.startswith("ema_trading_system") and f.endswith(".py"):
            ema_filename = f
            break

if not os.path.exists(pre_filename):
    for f in os.listdir("."):
        if "pre_trade" in f.lower() and f.endswith(".py"):
            pre_filename = f
            break

print(f"🔗 正在加载主系统: {ema_filename}")
ema_sys = load_module_safely(ema_filename, "ema_sys")

print(f"🔗 正在加载辅助系统: {pre_filename}")
pre_sys = load_module_safely(pre_filename, "pre_sys")

# ==================== 配置路径 ====================
FINAL_LIST_PATH = "Final list.xlsx"
HOLD_LIST_PATH = "Hold list.xlsx"
SIGNAL_MEMORY_PATH = "History/Signal_Memory.xlsx"
REPORTS_DIR = "Reports/Daily_Report"

os.makedirs(REPORTS_DIR, exist_ok=True)
os.makedirs("History", exist_ok=True)
os.makedirs("Logs", exist_ok=True)


def load_stock_codes(filepath, col_name_candidates):
    if not os.path.exists(filepath):
        print(f"⚠️ 文件不存在: {filepath}")
        return []
    try:
        df = pd.read_excel(filepath, dtype=str)
        target_col = None
        for c in df.columns:
            if any(cand.lower() in c.lower() for cand in col_name_candidates):
                target_col = c
                break
        if not target_col:
            target_col = df.columns[0]
        
        codes = []
        for val in df[target_col].dropna():
            cleaned = ema_sys.clean_stock_code(val)
            if cleaned and cleaned.lower() not in ['code', 'symbol', 'stock', '代码', '股票']:
                codes.append(cleaned)
        return list(dict.fromkeys(codes))
    except Exception as e:
        print(f"❌ 读取 {filepath} 失败: {e}")
        return []


def run_master_pipeline():
    today = datetime.now().date()
    
    # 1. 检查是否为周末或假期
    is_weekend = today.weekday() >= 5
    my_holidays = holidays.Malaysia(years=today.year)
    is_holiday = today in my_holidays

    if is_weekend or is_holiday:
        reason = f"周末（星期{'六' if today.weekday() == 5 else '日'}）" if is_weekend else f"马来西亚公共假期 [{my_holidays.get(today)}]"
        print(f"📅今天是{reason}。")
        if sys.stdin and sys.stdin.isatty():
            choice = input("👉 当前为非交易日，是否仍要强制执行扫描？(输入 y 确认运行，其他键退出): ").strip().lower()
            if choice != 'y':
                print("已取消运行。")
                return
        else:
            print("当前为自动化任务调度触发，已自动跳过。")
            return

    print("=" * 80)
    print(f"🚀 开始执行自动化双系统交易总控中心 (Master Pipeline) — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 80)

    watchlist_codes = load_stock_codes(FINAL_LIST_PATH, ['code', 'symbol', 'stock', '代码'])
    hold_codes = load_stock_codes(HOLD_LIST_PATH, ['code', 'symbol', 'stock', '代码'])
    
    print(f"📋 扫描候选池 (Final list.xlsx): {len(watchlist_codes)} 只")
    print(f"📋 持仓监控池 (Hold list.xlsx): {len(hold_codes)} 只")

    # 2. 执行扫描
    watchlist_results = []
    signal_memory_records = []
    
    print("\n--------------------------------------------------------------------------------")
    print("📈 Step 1: 执行 EMA Trading System 扫描与智能评分排名...")
    print("--------------------------------------------------------------------------------")
    
    for idx, code in enumerate(watchlist_codes, 1):
        print(f"[{idx}/{len(watchlist_codes)}] 扫描 {code}...", end=" ")
        try:
            df_raw, name = ema_sys.safe_fetch_stock(code)
            if df_raw is None or df_raw.empty:
                print("❌ 无数据")
                continue
            
            data = ema_sys.calculate_engine(df_raw)
            if data is None or data.empty:
                print("❌ 计算失败")
                continue
                
            row = data.iloc[-1]
            result = ema_sys.latest_output_dict(row, name, code)
            
            # Pre-Trade Checker 交叉验证
            pre_checker = pre_sys.PreTradeChecker(code)
            pre_res = pre_checker.check()
            if 'error' not in pre_res:
                result['PreTrade_Score'] = pre_res.get('score', 0)
                result['PreTrade_Verdict'] = pre_res.get('verdict', 'N/A')
                result['PreTrade_OBV_Status'] = pre_res.get('obv_consolidation_status', 'N/A')
            else:
                result['PreTrade_Score'] = 0
                result['PreTrade_Verdict'] = '检查失败'
                result['PreTrade_OBV_Status'] = 'N/A'

            # 综合量化得分计算 (EMA质量分占60%, PreTrade得分占40%)
            ema_score = float(result.get('股票质量分', 0) or 0)
            pre_score = float(result.get('PreTrade_Score', 0) or 0)
            composite_score = round((ema_score * 0.6) + (pre_score * 0.4), 2)
            result['综合量化得分'] = composite_score

            # 信号冲突标记
            ema_signal = result['综合建议']
            pre_verdict = result['PreTrade_Verdict']
            if "买" in ema_signal and "不建议" in pre_verdict:
                result['双系统信号冲突'] = "⚠️ 存在冲突 (EMA看多 vs PreTrade看空)"
            elif "不" in ema_signal and "买入" in pre_verdict:
                result['双系统信号冲突'] = "⚠️ 存在冲突 (EMA看空 vs PreTrade看多)"
            else:
                result['双系统信号冲突'] = "✅ 信号一致"

            watchlist_results.append(result)
            
            signal_memory_records.append({
                "日期": datetime.now().strftime('%Y-%m-%d'),
                "代码": code,
                "名称": name,
                "综合量化得分": composite_score,
                "阶段": result['阶段'],
                "Setup_Type": result['Setup_Type'],
                "综合建议": ema_signal,
                "PreTrade_结论": pre_verdict,
                "双系统信号冲突": result['双系统信号冲突']
            })
            print(f"✅ 完成 (综合得分: {composite_score} | {ema_signal})")
        except Exception as e:
            print(f"❌ 异常: {e}")

    # 3. 对扫描结果进行智能排名 (按综合量化得分从高到低排序)
    if watchlist_results:
        watchlist_results = sorted(watchlist_results, key=lambda x: x.get('综合量化得分', 0), reverse=True)
        # 加上排名序号
        for i, res in enumerate(watchlist_results, 1):
            res['量化排名'] = i
            # 前 3 名且无冲突的打上优选标签
            if i <= 3 and "⚠️" not in res['双系统信号冲突']:
                res['评级标签'] = "⭐ 核心优选 (Top Pick)"
            else:
                res['评级标签'] = "📋 正常观察"

    # 4. 执行持仓风险监控
    print("\n--------------------------------------------------------------------------------")
    print("🛡️ Step 2: 执行持仓自动风险监控 (Hold List)...")
    print("--------------------------------------------------------------------------------")
    
    hold_results = []
    for idx, code in enumerate(hold_codes, 1):
        print(f"[{idx}/{len(hold_codes)}] 检查持仓 {code}...", end=" ")
        try:
            df_raw, name = ema_sys.safe_fetch_stock(code)
            if df_raw is None or df_raw.empty:
                print("❌ 无数据")
                continue
            data = ema_sys.calculate_engine(df_raw)
            row = data.iloc[-1]
            
            exit_engine = row.get("Exit_Engine", "持有")
            top_warn = row.get("Top_Warning_Score", 0)
            top_reasons = row.get("Top_Warning_Reasons", "")
            atr_stop = row.get("ATR_Stop_Dynamic", np.nan)
            close_price = row.get("Close", 0)
            
            risk_status = "🟢 正常持有"
            if exit_engine == "EXIT" or top_warn >= 90:
                risk_status = "🔴 强烈建议清仓/止损"
            elif exit_engine in ("保护利润", "减仓") or top_warn >= 70:
                risk_status = "🟡 建议减仓/设好移动止损"
            elif top_warn >= 45:
                risk_status = "🟠 出现顶部预警，密切关注"

            hold_results.append({
                "代码": code,
                "名称": name,
                "当前收盘价": round(close_price, 4),
                "阶段": row.get("阶段", ""),
                "Exit_Engine 建议": exit_engine,
                "顶部预警分": top_warn,
                "预警原因": top_reasons,
                "ATR动态止损位": round(atr_stop, 4) if pd.notna(atr_stop) else np.nan,
                "持仓风险状态": risk_status
            })
            print(f"✅ 状态: {risk_status}")
        except Exception as e:
            print(f"❌ 异常: {e}")

    # 5. 生成统一综合 Excel 报告（带排名）
    date_str = datetime.now().strftime('%Y%m%d_%H%M')
    report_filename = os.path.join(REPORTS_DIR, f"Daily_Trading_Report_{date_str}.xlsx")
    
    df_watch = pd.DataFrame(watchlist_results) if watchlist_results else pd.DataFrame()
    df_hold = pd.DataFrame(hold_results) if hold_results else pd.DataFrame()
    df_memory = pd.DataFrame(signal_memory_records) if signal_memory_records else pd.DataFrame()

    df_today_buy = df_watch[df_watch['综合建议'].str.contains("今天可买|明日确认", na=False)] if not df_watch.empty else pd.DataFrame()
    df_wait = df_watch[df_watch['综合建议'].str.contains("等|观察|持有", na=False)] if not df_watch.empty else pd.DataFrame()
    df_conflict = df_watch[df_watch['双系统信号冲突'].str.contains("⚠️", na=False)] if not df_watch.empty else pd.DataFrame()

    with pd.ExcelWriter(report_filename, engine="openpyxl") as writer:
        dash_data = [
            {"项目": "📅 报告生成时间", "值": datetime.now().strftime('%Y-%m-%d %H:%M')},
            {"项目": "📊 扫描股票总数", "值": len(watchlist_codes)},
            {"项目": "🟢 今日/明日可买候选", "值": len(df_today_buy)},
            {"项目": "🔵 等待观察股票", "值": len(df_wait)},
            {"项目": "⚠️ 双系统信号冲突", "值": len(df_conflict)},
            {"项目": "🛡️ 监控持仓总数", "值": len(hold_codes)},
        ]
        pd.DataFrame(dash_data).to_excel(writer, sheet_name="⓪仪表盘", index=False)
        
        if not df_today_buy.empty:
            df_today_buy.to_excel(writer, sheet_name="①今日买入候选", index=False)
        if not df_wait.empty:
            df_wait.to_excel(writer, sheet_name="②等待确认股票", index=False)
        if not df_conflict.empty:
            df_conflict.to_excel(writer, sheet_name="③双系统信号冲突", index=False)
        if not df_hold.empty:
            df_hold.to_excel(writer, sheet_name="④持仓风险与卖出预警", index=False)
        if not df_watch.empty:
            df_watch.to_excel(writer, sheet_name="⑤全量扫描总表", index=False)
            
        if not df_memory.empty:
            if os.path.exists(SIGNAL_MEMORY_PATH):
                try:
                    old_mem = pd.read_excel(SIGNAL_MEMORY_PATH)
                    df_memory = pd.concat([old_mem, df_memory]).drop_duplicates(subset=['日期', '代码'], keep='last')
                except Exception:
                    pass
            df_memory.to_excel(SIGNAL_MEMORY_PATH, index=False)
            df_memory.to_excel(writer, sheet_name="⑥Signal_Memory", index=False)

    print("\n--------------------------------------------------------------------------------")
    print(f"✅ 自动化总控执行完毕（已完成智能排名与多维评分）！")
    print(f"📁 统一 Excel 报告已生成: {report_filename}")
    print(f"📁 Signal Memory 已更新: {SIGNAL_MEMORY_PATH}")
    print("--------------------------------------------------------------------------------")


if __name__ == "__main__":
    run_master_pipeline()