#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bursa Stock Analyzer Professional v1.0
下单前快速检查器 Pre-Trade Checker v3.0
【优化版】- 分层权重 + 硬性门槛 + 核心/辅助评分分离
【保留】OBV 盘整评估系统（相对强度 + 波段结构）
【新增】批量分析 – 读取 Final list.xlsx，一键扫描整个股票池
【ADX 对齐】ADX/+DI/-DI 改用 Wilder RMA，与 EMA_trading_system_V4.0.2a.ADX.py 一致
【批量回溯 V2】连续多天改用「天数」，支持多股
【涨跌金额】新增：当日涨跌金额（RM）+ 涨跌% ；多股报告统一按得分排序

使用方法：
python Pre_Trade_Checker_2.3.py
"""

import os
import sys
import logging
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import yfinance as yf
import warnings
warnings.filterwarnings('ignore')

# ==================== 配置 ====================
ERROR_LOG = "PreTrade_Error.log"
DATA_DAYS = 200
BATCH_FILE = "Final list.xlsx"          # 批量分析的默认文件名

# ==================== 日志配置 ====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(ERROR_LOG, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# ==================== 颜色输出 ====================
class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    BOLD = '\033[1m'
    END = '\033[0m'

def cprint(text: str, color: str = Colors.END) -> None:
    print(f"{color}{text}{Colors.END}")

# ==================== 工具函数 ====================

def format_code(code: str) -> str:
    code = str(code).strip().upper()
    if '.' in code:
        return code
    if code.isdigit():
        return code.zfill(4) + '.KL'
    return code + '.KL'

def get_stock_name(symbol: str) -> str:
    try:
        ticker = yf.Ticker(symbol)
        info = ticker.info
        if 'shortName' in info and info['shortName']:
            return info['shortName']
        elif 'longName' in info and info['longName']:
            return info['longName']
        return symbol.replace('.KL', '')
    except:
        return symbol.replace('.KL', '')

def download_data_until(symbol: str, end_date: str = None, days: int = DATA_DAYS):
    """下载数据，如果指定 end_date，则获取到该日期为止的数据（包含当天）"""
    try:
        ticker = yf.Ticker(symbol)
        if end_date:
            end_dt = datetime.strptime(end_date, '%Y-%m-%d')
            start_dt = end_dt - timedelta(days=days)
            df = ticker.history(
                start=start_dt.strftime('%Y-%m-%d'),
                end=(end_dt + timedelta(days=1)).strftime('%Y-%m-%d')
            )
        else:
            df = ticker.history(period=f"{days}d")
        if df.empty:
            return None
        df.columns = [col.capitalize() for col in df.columns]
        if df.index.tz is not None:
            df.index = df.index.tz_localize(None)
        return df
    except Exception as e:
        logger.error(f"下载失败: {e}")
        return None

# ==================== 涨跌格式化 ====================

def fmt_change_amount_pct(amount, pct):
    """
    显示：🔺 +0.0050 (+2.34%)  或  🔻 -0.0030 (-1.10%)
    """
    if amount is None or (isinstance(amount, float) and np.isnan(amount)):
        return "N/A"
    try:
        a = float(amount)
        p = float(pct) if pct is not None and not (isinstance(pct, float) and np.isnan(pct)) else 0.0
        arrow = "🔺" if a > 0 else ("🔻" if a < 0 else "➖")
        return f"{arrow} {a:+.4f} ({p:+.2f}%)"
    except Exception:
        return "N/A"

def fmt_change_amount_pct_short(amount, pct):
    """
    紧凑版（表格用）：+0.0050 (+2.34%)
    """
    if amount is None or (isinstance(amount, float) and np.isnan(amount)):
        return "N/A"
    try:
        a = float(amount)
        p = float(pct) if pct is not None and not (isinstance(pct, float) and np.isnan(pct)) else 0.0
        return f"{a:+.4f} ({p:+.2f}%)"
    except Exception:
        return "N/A"

# ==================== 指标计算 ====================

def calc_ema(df, period):
    return df['Close'].ewm(span=period, adjust=False).mean()

def calc_rsi(df, period=14):
    delta = df['Close'].diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=period).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=period).mean()
    rs = gain / loss
    rsi = 100 - (100 / (1 + rs))
    return rsi

def calc_macd(df):
    exp1 = df['Close'].ewm(span=12, adjust=False).mean()
    exp2 = df['Close'].ewm(span=26, adjust=False).mean()
    macd = exp1 - exp2
    signal = macd.ewm(span=9, adjust=False).mean()
    histogram = macd - signal
    return macd, signal, histogram

def calc_atr(df, period=14):
    high = df['High']
    low = df['Low']
    close = df['Close']
    tr1 = high - low
    tr2 = abs(high - close.shift())
    tr3 = abs(low - close.shift())
    tr = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
    return tr.rolling(period).mean()

def calc_obv(df):
    obv = [0]
    for i in range(1, len(df)):
        if df['Close'].iloc[i] > df['Close'].iloc[i-1]:
            obv.append(obv[-1] + df['Volume'].iloc[i])
        elif df['Close'].iloc[i] < df['Close'].iloc[i-1]:
            obv.append(obv[-1] - df['Volume'].iloc[i])
        else:
            obv.append(obv[-1])
    return pd.Series(obv, index=df.index)

# ==================== ADX Wilder（与 EMA_trading 对齐） ====================

def wilder_rma(series, period):
    """
    Wilder RMA 平滑，与 EMA_trading_system_V4.0.2a.ADX.py 中 wilder_rma 完全一致。
    """
    s = pd.Series(series, dtype="float64")
    result = pd.Series(np.nan, index=s.index, dtype="float64")

    values = s.to_numpy(dtype=float)
    out = np.full(len(values), np.nan)

    valid_positions = np.where(~np.isnan(values))[0]

    if len(valid_positions) < period:
        return result

    first_positions = valid_positions[:period]
    first_idx = first_positions[-1]
    first_mean = np.mean(values[first_positions])

    out[first_idx] = first_mean
    prev = first_mean

    for i in range(first_idx + 1, len(values)):
        x = values[i]
        if np.isnan(x):
            out[i] = np.nan
            continue
        prev = (prev * (period - 1) + x) / period
        out[i] = prev

    result.iloc[:] = out
    return result


def compute_adx_wilder(df, period=14):
    """与 EMA_trading_system_V4.0.2a.ADX.py 中 compute_adx_wilder 一致。"""
    high = pd.to_numeric(df["High"], errors="coerce")
    low = pd.to_numeric(df["Low"], errors="coerce")
    close = pd.to_numeric(df["Close"], errors="coerce")

    up_move = high.diff()
    down_move = -low.diff()

    plus_dm = pd.Series(np.nan, index=df.index, dtype="float64")
    minus_dm = pd.Series(np.nan, index=df.index, dtype="float64")

    valid = up_move.notna() & down_move.notna()

    plus_dm.loc[valid] = np.where(
        (up_move.loc[valid] > down_move.loc[valid]) &
        (up_move.loc[valid] > 0),
        up_move.loc[valid],
        0.0
    )

    minus_dm.loc[valid] = np.where(
        (down_move.loc[valid] > up_move.loc[valid]) &
        (down_move.loc[valid] > 0),
        down_move.loc[valid],
        0.0
    )

    tr = pd.concat(
        [
            high - low,
            (high - close.shift(1)).abs(),
            (low - close.shift(1)).abs(),
        ],
        axis=1
    ).max(axis=1)

    tr.iloc[0] = np.nan

    tr_rma = wilder_rma(tr, period)
    plus_dm_rma = wilder_rma(plus_dm, period)
    minus_dm_rma = wilder_rma(minus_dm, period)

    plus_di = 100.0 * plus_dm_rma / tr_rma.replace(0, np.nan)
    minus_di = 100.0 * minus_dm_rma / tr_rma.replace(0, np.nan)

    denominator = (plus_di + minus_di).replace(0, np.nan)
    dx = 100.0 * (plus_di - minus_di).abs() / denominator

    adx = wilder_rma(dx, period)

    return plus_di, minus_di, dx, adx


def calc_adx_di(df, period=14):
    """保持 Pre_Trade 原来的调用方式。"""
    plus_di, minus_di, dx, adx = compute_adx_wilder(df, period)

    adx_valid = adx.dropna()
    plus_valid = plus_di.dropna()
    minus_valid = minus_di.dropna()

    if adx_valid.empty or plus_valid.empty or minus_valid.empty:
        return np.nan, np.nan, np.nan

    return adx_valid.iloc[-1], plus_valid.iloc[-1], minus_valid.iloc[-1]

# ==================== 主检查器 ====================

class PreTradeChecker:
    def __init__(self, code: str, target_date: str = None):
        self.code = code
        self.symbol = format_code(code)
        self.name = get_stock_name(self.symbol)
        self.target_date = target_date
        self.df = download_data_until(self.symbol, end_date=target_date)
        self.result = {}
        
    def check(self):
        if self.df is None or len(self.df) < 50:
            return {'error': f'❌ 数据不足，无法分析 {self.symbol}'}
        self._calc_indicators()
        self._analyze()
        return self.result
    
    def _calc_indicators(self):
        df = self.df
        latest = df.iloc[-1]
        close = latest['Close']
        
        ema20 = calc_ema(df, 20).iloc[-1]
        ema50 = calc_ema(df, 50).iloc[-1]
        ema200 = calc_ema(df, 200).iloc[-1]
        rsi = calc_rsi(df).iloc[-1]
        macd, signal, _ = calc_macd(df)
        macd_bull = macd.iloc[-1] > signal.iloc[-1]
        atr = calc_atr(df).iloc[-1]
        obv = calc_obv(df)
        obv_trend = self._check_obv_trend(obv)
        obv_analysis = self._analyze_obv_consolidation(df, obv)
        adx, plus_di, minus_di = calc_adx_di(df)
        vol_ratio = latest['Volume'] / df['Volume'].tail(20).mean()
        support = df['Low'].tail(20).min()
        resistance = df['High'].tail(20).max()
        price_vs_ema20 = (close - ema20) / ema20 * 100

        # ===== 当日涨跌（金额 + %）=====
        prev_close_1d = np.nan
        daily_change_amount = np.nan
        daily_change_pct = np.nan
        if len(df) >= 2:
            prev_close_1d = df['Close'].iloc[-2]
            if pd.notna(prev_close_1d) and prev_close_1d > 0:
                daily_change_amount = close - prev_close_1d
                daily_change_pct = daily_change_amount / prev_close_1d * 100

        display_date = self.target_date if self.target_date else datetime.now().strftime('%Y-%m-%d %H:%M')
        
        self.result = {
            'code': self.code,
            'symbol': self.symbol,
            'name': self.name,
            'current_price': round(close, 3),
            'prev_close': round(prev_close_1d, 3) if pd.notna(prev_close_1d) else np.nan,
            'ema20': round(ema20, 3),
            'ema50': round(ema50, 3),
            'ema200': round(ema200, 3),
            'rsi': round(rsi, 1),
            'macd_bull': macd_bull,
            'atr': round(atr, 3),
            'obv_trend': obv_trend,
            'adx': round(adx, 1),
            'plus_di': round(plus_di, 1),
            'minus_di': round(minus_di, 1),
            'vol_ratio': round(vol_ratio, 2),
            'support': round(support, 3),
            'resistance': round(resistance, 3),
            'price_vs_ema20': round(price_vs_ema20, 1),
            'price_vs_ema50': round((close - ema50) / ema50 * 100, 1),
            'date': display_date,

            # 当日涨跌
            'daily_change_amount': round(daily_change_amount, 4) if pd.notna(daily_change_amount) else np.nan,
            'daily_change_pct': round(daily_change_pct, 2) if pd.notna(daily_change_pct) else np.nan,

            'obv_pullback_pct': obv_analysis['pullback_pct'],
            'obv_slope': obv_analysis['slope'],
            'obv_deviation': obv_analysis['deviation'],
            'obv_ma5_vs_ma20': obv_analysis['ma5_vs_ma20'],
            'obv_consolidation_status': obv_analysis['status'],
            'obv_detail_note': obv_analysis['note'],
            'is_backtest': bool(self.target_date),
        }
    
    def _check_obv_trend(self, obv):
        if len(obv) < 10:
            return '持平'
        recent = obv.iloc[-10:]
        inc = sum(1 for i in range(1, len(recent)) if recent.iloc[i] > recent.iloc[i-1])
        if inc >= 7:
            return '上升'
        elif inc <= 3:
            return '下降'
        return '持平'

    def _analyze_obv_consolidation(self, df, obv_series):
        if len(obv_series) < 30:
            return {
                'pullback_pct': 0,
                'slope': 0,
                'deviation': 1.0,
                'ma5_vs_ma20': 0,
                'status': '数据不足',
                'note': 'OBV 数据不足，无法评估盘整状态'
            }
        close_series = df['Close']
        lookback = min(60, len(obv_series))
        recent_obv = obv_series.iloc[-lookback:]
        recent_close = close_series.iloc[-lookback:]
        peak_obv = recent_obv.max()
        peak_idx = recent_obv.idxmax()
        current_obv = obv_series.iloc[-1]
        pre_peak_obv = obv_series.loc[:peak_idx].min() if peak_idx != obv_series.index[0] else peak_obv * 0.9
        wave_range = peak_obv - pre_peak_obv if peak_obv > pre_peak_obv else peak_obv * 0.1
        if wave_range > 0:
            pullback_pct = (peak_obv - current_obv) / wave_range * 100
        else:
            pullback_pct = 0
        pullback_pct = max(0, min(100, pullback_pct))
        if len(obv_series) >= 11:
            recent_10 = obv_series.iloc[-11:]
            slope = (recent_10.iloc[-1] - recent_10.iloc[0]) / 10
            avg_obv = recent_obv.mean()
            slope_pct = (slope / avg_obv * 100) if avg_obv > 0 else 0
        else:
            slope_pct = 0
        peak_price = close_series.loc[peak_idx] if peak_idx in close_series.index else close_series.iloc[-1]
        if peak_price > 0 and current_obv > 0 and peak_obv > 0:
            current_ratio = current_obv / close_series.iloc[-1]
            peak_ratio = peak_obv / peak_price
            deviation = current_ratio / peak_ratio if peak_ratio > 0 else 1.0
        else:
            deviation = 1.0
        deviation = max(0.5, min(1.5, deviation))
        obv_ma5 = obv_series.rolling(5).mean().iloc[-1]
        obv_ma20 = obv_series.rolling(20).mean().iloc[-1]
        if obv_ma20 > 0:
            ma5_vs_ma20 = (obv_ma5 - obv_ma20) / obv_ma20 * 100
        else:
            ma5_vs_ma20 = 0
        if pullback_pct < 38.2 and slope_pct > -3 and deviation > 0.95 and ma5_vs_ma20 > -3:
            status = "健康洗盘"
            note = f"回撤仅{pullback_pct:.1f}%，斜率{slope_pct:.1f}%，相对强度{deviation:.2f}，属正常换手"
        elif pullback_pct < 61.8 and slope_pct > -5 and deviation > 0.88 and ma5_vs_ma20 > -8:
            status = "温和盘整"
            note = f"回撤{pullback_pct:.1f}%，斜率{slope_pct:.1f}%，相对强度{deviation:.2f}，可继续观察"
        elif pullback_pct > 61.8 or slope_pct < -8 or deviation < 0.88 or ma5_vs_ma20 < -10:
            status = "资金出逃"
            note = f"回撤{pullback_pct:.1f}%，斜率{slope_pct:.1f}%，相对强度{deviation:.2f}，警惕主力派发"
        else:
            status = "中性"
            note = f"回撤{pullback_pct:.1f}%，斜率{slope_pct:.1f}%，相对强度{deviation:.2f}，方向不明"
        return {
            'pullback_pct': round(pullback_pct, 1),
            'slope': round(slope_pct, 2),
            'deviation': round(deviation, 3),
            'ma5_vs_ma20': round(ma5_vs_ma20, 2),
            'status': status,
            'note': note
        }
    
    def _reject_result(self, reason: str):
        """生成淘汰结果（硬性条件不满足）"""
        self.result['verdict'] = '🔴 暂不建议买入'
        self.result['verdict_color'] = Colors.RED
        self.result['score'] = 0
        self.result['core_score'] = 0
        self.result['helper_score'] = 0
        self.result['hard_conditions'] = {'passed': False, 'reason': reason}
        self.result['pros'] = []
        self.result['cons'] = [f'❌ 硬性淘汰：{reason}']
        self.result['entry_advice'] = ['🚫 不符合建仓条件，建议放弃']
        return self.result

    def _analyze(self):
        r = self.result
        
        # ============ 第一步：硬性条件 ============
        hard_checks = []
        
        if r['current_price'] <= r['ema50']:
            return self._reject_result(f"价格 RM{r['current_price']:.3f} ≤ EMA50 RM{r['ema50']:.3f}，中期趋势偏空")
        hard_checks.append('✅ 价格 > EMA50')
        
        obv_status = r.get('obv_consolidation_status', '')
        if obv_status == '资金出逃':
            return self._reject_result("OBV资金出逃，主力出货迹象")
        hard_checks.append(f'✅ OBV状态：{obv_status if obv_status else "正常"}')
        
        if r['adx'] <= 12:
            return self._reject_result(f"ADX {r['adx']:.1f} ≤ 12，无明显趋势")
        hard_checks.append(f'✅ ADX {r["adx"]:.1f} > 12')
        
        # ============ 第二步：核心得分 ============
        core_score = 0
        core_details = []
        
        if obv_status == '极健康洗盘':
            core_score += 3
            core_details.append('OBV极健康洗盘 +3')
        elif obv_status == '健康洗盘':
            core_score += 3
            core_details.append('OBV健康洗盘 +3')
        elif obv_status == '温和盘整':
            core_score += 0
            core_details.append('OBV温和盘整 0')
        elif obv_status == '中性':
            core_score += 0
            core_details.append('OBV中性 0')
        else:
            core_details.append('OBV状态未知 0')
        
        if r['adx'] >= 25:
            core_score += 3
            core_details.append(f'ADX {r["adx"]:.1f} ≥ 25 +3')
        elif r['adx'] >= 20:
            core_score += 2
            core_details.append(f'ADX {r["adx"]:.1f} ≥ 20 +2')
        elif r['adx'] >= 15:
            core_score += 1
            core_details.append(f'ADX {r["adx"]:.1f} ≥ 15 +1')
        else:
            core_details.append(f'ADX {r["adx"]:.1f} 偏低 +0')
        
        pct = r['price_vs_ema20']
        if -3 <= pct <= 3:
            core_score += 2
            core_details.append(f'价格接近EMA20 ({pct:+.1f}%) +2')
        elif 3 < pct <= 8:
            core_score += 1
            core_details.append(f'价格略高于EMA20 ({pct:+.1f}%) +1')
        elif 8 < pct <= 15:
            core_score += 0
            core_details.append(f'价格偏高 ({pct:+.1f}%) 0')
        elif pct > 15:
            core_score -= 1
            core_details.append(f'价格严重偏离 ({pct:+.1f}%) -1')
        else:
            core_score -= 2
            core_details.append(f'价格低于EMA20 ({pct:+.1f}%) -2')
        
        vol = r['vol_ratio']
        if vol >= 1.5:
            core_score += 2
            core_details.append(f'放量 {vol:.2f}倍 +2')
        elif vol >= 0.8:
            core_score += 1
            core_details.append(f'成交量正常 {vol:.2f}倍 +1')
        else:
            core_score -= 1
            core_details.append(f'成交量萎缩 {vol:.2f}倍 -1')
        
        core_score = max(-6, min(10, core_score))
        
        # ============ 第三步：辅助微调 ============
        helper_score = 0
        helper_details = []
        
        if r['macd_bull']:
            helper_score += 0.5
            helper_details.append('MACD多头 +0.5')
        else:
            helper_score -= 0.5
            helper_details.append('MACD空头 -0.5')
        
        if 40 <= r['rsi'] <= 70:
            helper_score += 0.5
            helper_details.append(f'RSI {r["rsi"]:.1f} 健康 +0.5')
        elif r['rsi'] > 75 or r['rsi'] < 30:
            helper_score -= 0.5
            helper_details.append(f'RSI {r["rsi"]:.1f} 极端区域 -0.5')
        else:
            helper_details.append(f'RSI {r["rsi"]:.1f} 中性 0')
        
        if r['current_price'] <= r['support'] * 1.05:
            helper_score += 0.5
            helper_details.append('接近支撑位 +0.5')
        elif r['current_price'] >= r['resistance'] * 0.98:
            helper_score -= 0.5
            helper_details.append('接近压力位 -0.5')
        else:
            helper_details.append('位置中性 0')
        
        if r['plus_di'] > r['minus_di']:
            helper_score += 0.5
            helper_details.append('多头占优 +0.5')
        else:
            helper_score -= 0.5
            helper_details.append('空头占优 -0.5')
        
        helper_score = max(-2, min(2, helper_score))
        
        total_score = core_score + helper_score
        
        # ============ 第五步：最终结论 ============
        if total_score >= 8:
            verdict = '🟢 可以买入'
            verdict_color = Colors.GREEN
            position_advice = '60-80%'
        elif total_score >= 5:
            verdict = '🟡 谨慎买入'
            verdict_color = Colors.YELLOW
            position_advice = '30-50%'
        elif total_score >= 2:
            verdict = '🟠 建议观望'
            verdict_color = Colors.YELLOW
            position_advice = '10-20%（轻仓试错）'
        else:
            verdict = '🔴 暂不建议买入'
            verdict_color = Colors.RED
            position_advice = '0%'
        
        # ============ 第六步：建仓建议 ============
        entry_advice = []
        if verdict in ['🟢 可以买入', '🟡 谨慎买入']:
            if r['current_price'] <= r['ema20'] * 1.03:
                entry_advice.append(f'🎯 当前价 RM {r["current_price"]} 接近EMA20，适合建仓')
            elif r['current_price'] <= r['ema20'] * 1.08:
                entry_advice.append(f'🎯 建议建仓价：RM {r["ema20"]:.3f}（EMA20附近）')
            else:
                entry_advice.append(f'🎯 溢价偏高({r["price_vs_ema20"]:.1f}%)，建议等待回踩到 RM {r["ema20"]:.3f} 附近')
        else:
            entry_advice.append('🚫 暂不建议建仓')
        
        if verdict in ['🟢 可以买入', '🟡 谨慎买入']:
            stop_loss = min(r['current_price'] - 1.5 * r['atr'], r['support'] * 0.99)
            entry_advice.append(f'🛑 止损价：RM {stop_loss:.3f}（约 {((r["current_price"]-stop_loss)/r["current_price"]*100):.1f}% 亏损）')
        
        # ============ 第七步：优点/隐忧 ============
        pros = []
        cons = []
        
        for check in hard_checks:
            pros.append(check)
        
        if core_score >= 4:
            pros.append(f'✅ 核心评分 {core_score:.1f}，条件良好')
        elif core_score <= -2:
            cons.append(f'⚠️ 核心评分 {core_score:.1f}，多项核心条件偏弱')
        
        if obv_status in ['极健康洗盘', '健康洗盘']:
            pros.append(f'✅ OBV {obv_status}，资金健康')
        elif obv_status == '资金出逃':
            cons.append(f'⚠️ OBV {obv_status}，警惕出货')
        
        if r['adx'] >= 25:
            pros.append(f'✅ ADX {r["adx"]:.1f}，趋势明确')
        elif r['adx'] < 18:
            cons.append(f'⚠️ ADX {r["adx"]:.1f}，趋势偏弱')
        
        if -3 <= r['price_vs_ema20'] <= 3:
            pros.append(f'✅ 价格接近EMA20，建仓合理')
        elif r['price_vs_ema20'] > 10:
            cons.append(f'⚠️ 价格偏离EMA20达 {r["price_vs_ema20"]:.1f}%，溢价过高')
        elif r['price_vs_ema20'] < -5:
            cons.append(f'⚠️ 价格低于EMA20 {abs(r["price_vs_ema20"]):.1f}%，趋势偏弱')
        
        if r['vol_ratio'] >= 1.2:
            pros.append(f'✅ 成交量放大，有资金关注')
        elif r['vol_ratio'] < 0.5:
            cons.append('⚠️ 成交量明显萎缩，关注度下降')
        
        if r['macd_bull']:
            pros.append('✅ MACD多头，动能向上')
        else:
            cons.append('⚠️ MACD空头，动能向下')
        
        if 40 <= r['rsi'] <= 70:
            pros.append('✅ RSI中性偏强，不超买')
        elif r['rsi'] > 75:
            cons.append(f'⚠️ RSI {r["rsi"]:.1f}，超买区域')
        elif r['rsi'] < 30:
            cons.append(f'⚠️ RSI {r["rsi"]:.1f}，超卖区域')
        
        if r['plus_di'] > r['minus_di']:
            pros.append('✅ 多头力量占优')
        else:
            cons.append(f'⚠️ 空头力量占优')
        
        if r['current_price'] <= r['support'] * 1.05:
            pros.append(f'✅ 价格接近支撑位 RM {r["support"]:.3f}')
        elif r['current_price'] >= r['resistance'] * 0.98:
            cons.append(f'⚠️ 价格接近压力位 RM {r["resistance"]:.3f}')
        
        self.result['hard_checks'] = hard_checks
        self.result['core_score'] = core_score
        self.result['core_details'] = core_details
        self.result['helper_score'] = helper_score
        self.result['helper_details'] = helper_details
        self.result['score'] = total_score
        self.result['verdict'] = verdict
        self.result['verdict_color'] = verdict_color
        self.result['position_advice'] = position_advice
        self.result['entry_advice'] = entry_advice
        self.result['pros'] = pros
        self.result['cons'] = cons
    
    def print_report(self):
        if 'error' in self.result:
            cprint(f"\n{self.result['error']}", Colors.RED)
            return
        r = self.result
        
        print("\n" + "=" * 70)
        title = "  📊 下单前检查报告 v3.0（优化版）"
        if r.get('is_backtest'):
            title += " [历史回溯]"
        cprint(title, Colors.BOLD)
        print("=" * 70)
        if r.get('is_backtest'):
            print(f"  ⏳ 回溯日期：{r['date']}")
        print(f"\n  {r['code']}  {r['name']}")
        print(f"  当前价格：RM {r['current_price']}")
        if pd.notna(r.get('prev_close', np.nan)):
            print(f"  昨日收价：RM {r['prev_close']}")
        print(f"  今日涨跌：{fmt_change_amount_pct(r.get('daily_change_amount'), r.get('daily_change_pct'))}")
        print(f"  分析时间：{r['date']}")
        
        print("\n" + "-" * 70)
        cprint("  【硬性条件检查】", Colors.BOLD)
        print("-" * 70)
        for check in r.get('hard_checks', []):
            print(f"    {check}")
        
        print("\n" + "-" * 70)
        cprint("  【OBV 盘整评估】", Colors.BOLD)
        print("-" * 70)
        print(f"  状态：{r.get('obv_consolidation_status', 'N/A')}")
        print(f"  回撤幅度：{r.get('obv_pullback_pct', 0):.1f}%")
        print(f"  斜率：{r.get('obv_slope', 0):.2f}%")
        print(f"  相对强度：{r.get('obv_deviation', 1):.3f}  (>0.95为健康)")
        print(f"  评估结论：{r.get('obv_detail_note', '')}")
        
        print("\n" + "-" * 70)
        cprint("  【分层评分】", Colors.BOLD)
        print("-" * 70)
        print(f"  核心得分：{r.get('core_score', 0):.1f}  (满分 10 分)")
        for detail in r.get('core_details', []):
            print(f"    {detail}")
        print(f"  辅助微调：{r.get('helper_score', 0):.1f}  (±2 分)")
        for detail in r.get('helper_details', []):
            print(f"    {detail}")
        print(f"\n  📊 总得分：{r.get('score', 0):.1f}  (核心+辅助)")
        
        print("\n" + "-" * 70)
        cprint("  【检查结果】", Colors.BOLD)
        print("-" * 70)
        if r.get('pros'):
            cprint("\n  ✅ 优点：", Colors.GREEN)
            for p in r['pros']:
                print(f"    {p}")
        if r.get('cons'):
            cprint("\n  ⚠️ 隐忧：", Colors.RED)
            for c in r['cons']:
                print(f"    {c}")
        if not r.get('pros') and not r.get('cons'):
            print("  无明显优点或隐忧")
        
        print("\n" + "-" * 70)
        cprint("  【建仓建议】", Colors.BOLD)
        print("-" * 70)
        for a in r.get('entry_advice', []):
            print(f"    {a}")
        
        print("\n" + "-" * 70)
        cprint("  【综合结论】", Colors.BOLD)
        print("-" * 70)
        cprint(f"\n  {r['verdict']}  (总得分: {r.get('score', 0):.1f})", r['verdict_color'])
        print(f"  建议仓位：{r.get('position_advice', '0%')}")
        print("\n" + "=" * 70)

# ==================== 批量分析类 ====================

class PreTradeBatchAnalyzer:
    def __init__(self, filename: str = BATCH_FILE):
        self.filename = filename
        self.stocks = []
        self.results = []
        self._load_stocks()
        
    def _load_stocks(self):
        if not os.path.exists(self.filename):
            print(f"❌ 文件 {self.filename} 不存在")
            return
        try:
            df = pd.read_excel(self.filename, dtype=str)
            cols = df.columns
            target_col = cols[0]
            for col in cols:
                if col.lower() in ['code', 'symbol', 'stock', 'ticker', '代码', '股票']:
                    target_col = col
                    break
            codes = df[target_col].dropna().astype(str).str.strip().tolist()
            codes = [c.replace('.KL', '').replace('.SS', '').replace('.SZ', '') for c in codes]
            codes = [c.zfill(4) if c.isdigit() else c for c in codes]
            codes = [c for c in codes if c and c.lower() not in ['code', 'symbol', 'stock', '股票']]
            self.stocks = list(dict.fromkeys(codes))
            print(f"✅ 从 {self.filename} 加载了 {len(self.stocks)} 只股票")
        except Exception as e:
            print(f"❌ 读取 {self.filename} 失败: {e}")
    
    def run(self):
        if not self.stocks:
            print("📭 没有可分析的股票")
            return
        
        print(f"\n🔍 开始批量分析 {len(self.stocks)} 只股票...\n")
        for idx, code in enumerate(self.stocks, 1):
            print(f"  [{idx}/{len(self.stocks)}] {code}...", end=" ")
            checker = PreTradeChecker(code)
            result = checker.check()
            if 'error' in result:
                print(f"❌ {result['error']}")
                self.results.append({
                    'code': code,
                    'name': 'N/A',
                    'error': result['error'],
                    'score': None,
                    'verdict': '❌ 分析失败'
                })
            else:
                self.results.append({
                    'code': code,
                    'name': result['name'],
                    'price': result['current_price'],
                    'prev_close': result.get('prev_close', np.nan),
                    'daily_change_amount': result.get('daily_change_amount', np.nan),
                    'daily_change_pct': result.get('daily_change_pct', np.nan),
                    'score': result['score'],
                    'verdict': result['verdict'],
                    'obv_status': result.get('obv_consolidation_status', 'N/A'),
                    'adx': result['adx'],
                    'rsi': result['rsi'],
                    'vol_ratio': result['vol_ratio'],
                    'macd_bull': result['macd_bull'],
                    'price_vs_ema20': result['price_vs_ema20'],
                    'support': result['support'],
                    'resistance': result['resistance'],
                    'plus_di': result['plus_di'],
                    'minus_di': result['minus_di'],
                })
                print(f"✅ 得分 {result['score']:.1f}  {result['verdict']}")
        
        self._print_summary()
        self._export_excel()
    
    def _print_summary(self):
        print("\n" + "=" * 110)
        print(f"  📊 批量分析汇总 (共 {len(self.results)} 只) — 按得分由高到低排序")
        print("=" * 110)
        
        sorted_res = sorted([r for r in self.results if r.get('score') is not None], 
                           key=lambda x: x['score'], reverse=True)
        
        print("\n" + "-" * 110)
        print("排名 代码   名称              当前价    昨日价    当日涨跌(RM)           得分  结论")
        print("-" * 110)
        for i, r in enumerate(sorted_res, 1):
            code = r['code']
            name = r.get('name', '')[:12].ljust(12)
            price = f"{r.get('price', 0):.3f}".rjust(7)
            prev_p = r.get('prev_close', np.nan)
            prev_str = f"{prev_p:.3f}".rjust(7) if pd.notna(prev_p) else "N/A".rjust(7)
            d1 = fmt_change_amount_pct_short(r.get('daily_change_amount'), r.get('daily_change_pct')).rjust(20)
            score = f"{r['score']:.1f}".rjust(5)
            verdict = r['verdict'][:12].ljust(12)
            print(f"{i:>3}  {code:<6} {name} {price}  {prev_str}  {d1}  {score}  {verdict}")
        print("-" * 110)
        
        buy = sum(1 for r in self.results if r.get('verdict') == '🟢 可以买入')
        caution = sum(1 for r in self.results if r.get('verdict') == '🟡 谨慎买入')
        watch = sum(1 for r in self.results if r.get('verdict') == '🟠 建议观望')
        avoid = sum(1 for r in self.results if r.get('verdict') == '🔴 暂不建议买入')
        fail = sum(1 for r in self.results if r.get('score') is None)
        print(f"\n📈 统计：")
        print(f"  🟢 可以买入: {buy} 只")
        print(f"  🟡 谨慎买入: {caution} 只")
        print(f"  🟠 建议观望: {watch} 只")
        print(f"  🔴 暂不建议买入: {avoid} 只")
        if fail:
            print(f"  ❌ 分析失败: {fail} 只")
        print("=" * 110)
    
    def _export_excel(self):
        if not self.results:
            return
        # 按得分排序后导出
        sorted_res = sorted(
            [r for r in self.results if r.get('score') is not None],
            key=lambda x: x['score'], reverse=True
        )
        rows = []
        for i, r in enumerate(sorted_res, 1):
            rows.append({
                '排名': i,
                '代码': r.get('code', ''),
                '名称': r.get('name', ''),
                '当前价': r.get('price', 0),
                '昨日价': r.get('prev_close', np.nan),
                '当日涨跌(RM)': r.get('daily_change_amount', np.nan),
                '当日涨跌%': r.get('daily_change_pct', np.nan),
                '得分': r.get('score', 0),
                '结论': r.get('verdict', ''),
                'OBV状态': r.get('obv_status', ''),
                'ADX': r.get('adx', 0),
                'RSI': r.get('rsi', 0),
                '成交量比例': r.get('vol_ratio', 0),
                'MACD多头': r.get('macd_bull', False),
                '价格vsEMA20%': r.get('price_vs_ema20', 0),
                '支撑': r.get('support', 0),
                '压力': r.get('resistance', 0),
                '+DI': r.get('plus_di', 0),
                '-DI': r.get('minus_di', 0),
                '错误信息': r.get('error', '')
            })
        df = pd.DataFrame(rows)
        date_str = datetime.now().strftime("%Y%m%d")
        filename = f"batch_report_{date_str}.xlsx"
        try:
            df.to_excel(filename, index=False, engine='openpyxl')
            print(f"✅ 详细报告已导出（按得分排序）: {filename}")
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_filename = f"batch_report_{date_str}_{timestamp}.xlsx"
            try:
                df.to_excel(new_filename, index=False, engine='openpyxl')
                print(f"⚠️ 原文件被占用，已另存为: {new_filename}")
            except Exception as e:
                print(f"❌ 导出失败: {e}")
        except Exception as e:
            print(f"❌ 导出失败: {e}")

# ==================== 批量历史回溯类（连续多天，按天数输入） ====================

class PreTradeBatchBacktest:
    """
    单只股票的连续多天回溯。
    输入：code + days（最近 N 个交易日）
    """
    def __init__(self, code: str, days: int = 60):
        self.code = code
        self.symbol = format_code(code)
        self.name = get_stock_name(self.symbol)
        self.days = int(days)
        self.results = []

    def run(self):
        print(f"\n⏳ 正在回溯 {self.code} 最近 {self.days} 个交易日...")
        try:
            ticker = yf.Ticker(self.symbol)
            fetch_period_days = (self.days * 2) + DATA_DAYS + 100
            df_full = ticker.history(period=f"{fetch_period_days}d")
            if df_full.empty:
                print(f"❌ {self.code} 无法获取数据")
                return None
            df_full.columns = [col.capitalize() for col in df_full.columns]
            if df_full.index.tz is not None:
                df_full.index = df_full.index.tz_localize(None)
        except Exception as e:
            print(f"❌ {self.code} 数据获取失败: {e}")
            return None

        if len(df_full) < 50:
            print(f"❌ {self.code} 数据不足（<50 天）")
            return None

        n = min(self.days, len(df_full))
        trading_dates = list(df_full.index[-n:])

        print(f"📋 {self.code} 找到 {len(trading_dates)} 个交易日，开始分析...\n")

        for date_idx in trading_dates:
            date_str = date_idx.strftime('%Y-%m-%d')
            df_until = df_full[df_full.index <= date_idx].copy()
            if df_until.empty or len(df_until) < 50:
                continue
            try:
                checker = PreTradeChecker(self.code)
                checker.df = df_until
                checker.target_date = date_str
                result = checker.check()
                if 'error' not in result:
                    result['backtest_date'] = date_str
                    self.results.append(result)
            except Exception:
                continue

        print(f"✅ {self.code} 完成 {len(self.results)} 天回溯")
        return self.results

    def get_summary(self):
        """返回该股票的汇总行（供多股对比用），含回溯期间涨跌金额 + %"""
        if not self.results:
            return None
        sorted_by_date = sorted(self.results, key=lambda x: x.get('backtest_date', ''))
        latest = sorted_by_date[-1]
        first = sorted_by_date[0]
        scores = [r.get('score', 0) for r in self.results]

        first_price = first.get('current_price', np.nan)
        latest_price = latest.get('current_price', np.nan)
        if pd.notna(first_price) and first_price > 0 and pd.notna(latest_price):
            period_change_amount = latest_price - first_price
            period_change_pct = period_change_amount / first_price * 100
        else:
            period_change_amount = np.nan
            period_change_pct = np.nan

        return {
            'code': self.code,
            'name': self.name,
            'days_backtested': len(self.results),
            'first_date': first.get('backtest_date', ''),
            'first_price': first_price,
            'latest_date': latest.get('backtest_date', ''),
            'latest_price': latest_price,
            'latest_prev_close': latest.get('prev_close', np.nan),
            'latest_daily_change_amount': latest.get('daily_change_amount', np.nan),
            'latest_daily_change_pct': latest.get('daily_change_pct', np.nan),
            'period_change_amount': period_change_amount,
            'period_change_pct': period_change_pct,
            'latest_score': float(scores[-1]),
            'avg_score': float(np.mean(scores)),
            'max_score': float(np.max(scores)),
            'min_score': float(np.min(scores)),
        }

    def print_report(self):
        """打印单只股票的逐日明细（含每日涨跌金额 + %）"""
        if not self.results:
            print(f"❌ {self.code} 无有效结果")
            return

        sorted_by_date = sorted(self.results, key=lambda x: x.get('backtest_date', ''))

        print("\n" + "=" * 95)
        print(f"  📊 历史回溯明细 ({self.code} {self.name})")
        print("=" * 95)
        print(f"  回溯天数: 最近 {len(self.results)} 个交易日")

        first_price = sorted_by_date[0].get('current_price', np.nan)
        latest_price = sorted_by_date[-1].get('current_price', np.nan)
        if pd.notna(first_price) and first_price > 0 and pd.notna(latest_price):
            period_amt = latest_price - first_price
            period_pct = period_amt / first_price * 100
            print(f"  期间涨跌: RM {first_price:.3f} → RM {latest_price:.3f}  "
                  f"（{period_amt:+.4f}  RM,  {period_pct:+.2f}%）")

        print("\n" + "-" * 95)
        print("日期        价格      昨日      当日涨跌(RM)           得分  结论")
        print("-" * 95)

        prev_score = None
        for i, r in enumerate(sorted_by_date):
            date = r.get('backtest_date', 'N/A')
            price = r['current_price']
            prev_p = r.get('prev_close', np.nan)
            d1_a = r.get('daily_change_amount', np.nan)
            d1_p = r.get('daily_change_pct', np.nan)
            score = r.get('score', 0)
            verdict = r.get('verdict', 'N/A')

            prev_str = f"{prev_p:.3f}" if pd.notna(prev_p) else "N/A"
            d1_str = fmt_change_amount_pct_short(d1_a, d1_p)

            print(f"{date}  {price:>7.3f}  {prev_str:>7}  {d1_str:>20}  "
                  f"{score:>5.1f}  {verdict}")
            prev_score = score

        print("-" * 95)

        sorted_by_score = sorted(self.results, key=lambda x: x.get('score', 0), reverse=True)
        print(f"\n🏆 {self.code} 最佳得分日 TOP 5：")
        for r in sorted_by_score[:5]:
            date = r.get('backtest_date', 'N/A')
            price = r['current_price']
            d1_a = r.get('daily_change_amount', np.nan)
            d1_p = r.get('daily_change_pct', np.nan)
            d1_str = fmt_change_amount_pct_short(d1_a, d1_p)
            score = r.get('score', 0)
            verdict = r.get('verdict', 'N/A')
            print(f"   {date}  价格 {price:>7.3f}  当日 {d1_str:>20}  得分 {score:>5.1f}  {verdict}")

        scores = [r.get('score', 0) for r in self.results]
        print(f"\n📈 统计：最新={scores[-1]:.1f}，平均={np.mean(scores):.1f}，"
              f"最高={np.max(scores):.1f}，最低={np.min(scores):.1f}")
        print("=" * 95)

# ==================== 交互式主程序 ====================

def main():
    print("\n" + "=" * 70)
    cprint("  🔥 下单前快速检查器 Pre-Trade Checker v2.3 (优化版+批量分析)", Colors.BOLD)
    print("=" * 70)
    print("\n  功能：")
    print("    [1] 当前分析")
    print("    [2] 单日历史回溯")
    print("    [3] 批量历史回溯（连续多天，按天数，多股对比按得分排序）")
    print("    [4] 批量分析（读取 Final list.xlsx）")
    print("    [5] 退出")
    print("=" * 70)
    print("  💡 优化说明：")
    print("    - 硬性条件：价格>EMA50、OBV不逃、ADX>12")
    print("    - 分层评分：核心得分(OBV/ADX/价格/量) + 辅助微调(MACD/RSI/支撑/DI)")
    print("    - 结论更严格，减少过度乐观")
    print("    - ADX 使用 Wilder RMA，与 EMA_trading V4.0.2 对齐")
    print("    - 报告含当日涨跌金额(RM) + 涨跌%")
    print("    - 多股报告统一按得分由高到低排列")
    print("=" * 70)
    
    while True:
        print("\n" + "-" * 50)
        choice = input("请选择 (1/2/3/4/5): ").strip()
        
        if choice == '5':
            print("\n👋 再见！")
            break
            
        elif choice == '1':
            code = input("  请输入股票代码: ").strip()
            if not code:
                print("❌ 请输入有效的股票代码")
                continue
            print("\n⏳ 正在分析...\n")
            checker = PreTradeChecker(code)
            result = checker.check()
            if 'error' in result:
                cprint(f"\n{result['error']}", Colors.RED)
                continue
            checker.print_report()
            input("\n按 Enter 继续...")
            
        elif choice == '2':
            code = input("  请输入股票代码: ").strip()
            if not code:
                print("❌ 请输入有效的股票代码")
                continue
            date_str = input("  请输入回溯日期 (格式 YYYY-MM-DD): ").strip()
            try:
                datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                print("❌ 日期格式错误")
                continue
            print(f"\n⏳ 正在回溯到 {date_str} ...\n")
            checker = PreTradeChecker(code, target_date=date_str)
            result = checker.check()
            if 'error' in result:
                cprint(f"\n{result['error']}", Colors.RED)
                continue
            checker.print_report()
            input("\n按 Enter 继续...")
            
        elif choice == '3':
            # ============ 批量历史回溯（连续多天，按天数） ============
            codes_input = input("  请输入股票代码（可逗号分隔多个）: ").strip()
            if not codes_input:
                print("❌ 请输入有效的股票代码")
                continue
            codes = [c.strip() for c in codes_input.split(',') if c.strip()]
            
            days_input = input("  请输入回溯天数 [60]: ").strip()
            if days_input.isdigit():
                days = int(days_input)
            else:
                days = 60
            if days <= 0:
                days = 60
            print(f"\n📋 将回溯每只股票最近 {days} 个交易日\n")
            
            all_summaries = []
            bt_instances = {}
            
            for code in codes:
                bt = PreTradeBatchBacktest(code, days)
                res = bt.run()
                if res:
                    bt_instances[code] = bt
                    summary = bt.get_summary()
                    if summary:
                        all_summaries.append(summary)
            
            if not all_summaries:
                print("❌ 无有效回溯结果")
                input("\n按 Enter 继续...")
                continue
            
            # ============ 按最新得分从高到低排序 ============
            all_summaries.sort(key=lambda x: x['latest_score'], reverse=True)
            
            print("\n" + "=" * 125)
            cprint("  🏆 批量回溯摘要（股票按最新得分由高到低编排）", Colors.BOLD)
            print("=" * 125)
            print(f"排名  代码   名称              回溯  最新日期     最新价    昨日价    当日涨跌(RM)           期间涨跌(RM)            最新得分  平均  最高")
            print("-" * 125)
            for i, s in enumerate(all_summaries, 1):
                name_trunc = s['name'][:16]
                prev_str = f"{s['latest_prev_close']:.3f}" if pd.notna(s.get('latest_prev_close', np.nan)) else "N/A"
                d1 = fmt_change_amount_pct_short(
                    s.get('latest_daily_change_amount'),
                    s.get('latest_daily_change_pct')
                ).rjust(20)
                pc = fmt_change_amount_pct_short(
                    s.get('period_change_amount'),
                    s.get('period_change_pct')
                ).rjust(20)
                print(f"{i:>3}   {s['code']:<6} {name_trunc:<18} "
                      f"{s['days_backtested']:>4}   "
                      f"{s['latest_date']:<12}  "
                      f"{s['latest_price']:>7.3f}   "
                      f"{prev_str:>7}   "
                      f"{d1}   "
                      f"{pc}   "
                      f"{s['latest_score']:>7.1f}   "
                      f"{s['avg_score']:>6.1f}   "
                      f"{s['max_score']:>6.1f}")
            print("=" * 125)
            
            show_detail = input("\n是否显示每只股票的逐日明细？(y/n) [n]: ").strip().lower()
            if show_detail == 'y':
                for s in all_summaries:
                    bt = bt_instances.get(s['code'])
                    if bt:
                        bt.print_report()
            
            export = input("\n是否导出Excel报告？(y/n): ").strip().lower()
            if export == 'y':
                date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"batch_backtest_{date_str}.xlsx"
                
                summary_rows = []
                for rank, s in enumerate(all_summaries, 1):
                    summary_rows.append({
                        '排名': rank,
                        '代码': s['code'],
                        '名称': s['name'],
                        '回溯天数': s['days_backtested'],
                        '起始日期': s['first_date'],
                        '起始价': s['first_price'],
                        '最新日期': s['latest_date'],
                        '最新价': s['latest_price'],
                        '昨日价': s.get('latest_prev_close', np.nan),
                        '当日涨跌(RM)': s.get('latest_daily_change_amount', np.nan),
                        '当日涨跌%': s.get('latest_daily_change_pct', np.nan),
                        '期间涨跌(RM)': round(s['period_change_amount'], 4) if pd.notna(s['period_change_amount']) else np.nan,
                        '期间涨跌%': round(s['period_change_pct'], 2) if pd.notna(s['period_change_pct']) else np.nan,
                        '最新得分': round(s['latest_score'], 2),
                        '平均得分': round(s['avg_score'], 2),
                        '最高得分': round(s['max_score'], 2),
                        '最低得分': round(s['min_score'], 2),
                    })
                summary_df = pd.DataFrame(summary_rows)
                
                def _safe_sheet(name, max_len=31):
                    name = str(name)
                    for ch in ['\\', '/', '*', '?', ':', '[', ']']:
                        name = name.replace(ch, '_')
                    return name[:max_len]
                
                try:
                    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                        summary_df.to_excel(
                            writer, sheet_name='摘要（按得分排序）', index=False
                        )
                        for s in all_summaries:
                            bt = bt_instances.get(s['code'])
                            if not bt:
                                continue
                            rows = []
                            for r in sorted(bt.results, key=lambda x: x.get('backtest_date', '')):
                                rows.append({
                                    '日期': r.get('backtest_date', 'N/A'),
                                    '价格': r['current_price'],
                                    '昨日价': r.get('prev_close', np.nan),
                                    '当日涨跌(RM)': r.get('daily_change_amount', np.nan),
                                    '当日涨跌%': r.get('daily_change_pct', np.nan),
                                    '总得分': r.get('score', 0),
                                    '核心得分': r.get('core_score', 0),
                                    '辅助得分': r.get('helper_score', 0),
                                    '结论': r.get('verdict', 'N/A'),
                                    'OBV状态': r.get('obv_consolidation_status', 'N/A'),
                                    'ADX': r['adx'],
                                    'RSI': r['rsi'],
                                    '价格vsEMA20%': r['price_vs_ema20'],
                                    '成交量比例': r['vol_ratio'],
                                })
                            detail_df = pd.DataFrame(rows)
                            sheet_name = _safe_sheet(f"{s['code']}_{s['name']}")
                            detail_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"✅ 已导出（摘要按得分排序）: {filename}")
                except PermissionError:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    new_filename = f"batch_backtest_{timestamp}.xlsx"
                    try:
                        with pd.ExcelWriter(new_filename, engine='openpyxl') as writer:
                            summary_df.to_excel(
                                writer, sheet_name='摘要（按得分排序）', index=False
                            )
                            for s in all_summaries:
                                bt = bt_instances.get(s['code'])
                                if not bt:
                                    continue
                                rows = []
                                for r in sorted(bt.results, key=lambda x: x.get('backtest_date', '')):
                                    rows.append({
                                        '日期': r.get('backtest_date', 'N/A'),
                                        '价格': r['current_price'],
                                        '当日涨跌(RM)': r.get('daily_change_amount', np.nan),
                                        '当日涨跌%': r.get('daily_change_pct', np.nan),
                                        '总得分': r.get('score', 0),
                                        '结论': r.get('verdict', 'N/A'),
                                    })
                                detail_df = pd.DataFrame(rows)
                                sheet_name = _safe_sheet(f"{s['code']}_{s['name']}")
                                detail_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"⚠️ 原文件被占用，已另存为: {new_filename}")
                    except Exception as e:
                        print(f"❌ 导出失败: {e}")
                except Exception as e:
                    print(f"❌ 导出失败: {e}")
            
            input("\n按 Enter 继续...")
            
        elif choice == '4':
            analyzer = PreTradeBatchAnalyzer(BATCH_FILE)
            analyzer.run()
            input("\n按 Enter 继续...")
            
        else:
            print("❌ 无效选项")

if __name__ == "__main__":
    main()