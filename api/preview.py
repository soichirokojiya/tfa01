"""
Vercel Serverless Function: 自動取得データのプレビュー
POST /api/preview
銘柄コード・評価基準日・権利行使期間終了日から自動取得される数値を返す
"""

import json
import math
import re
import urllib.request
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from http.server import BaseHTTPRequestHandler
import yfinance as yf
import numpy as np
import warnings
warnings.filterwarnings("ignore")


def fetch_jsda_bond(eval_dt, exercise_end_dt) -> dict:
    """JSDA売買参考統計値から、権利行使期間終了日に最も近い長期国債(超長期除く)を返す。"""
    result = {"name": "", "maturity": "", "yield_value": ""}
    try:
        import xlrd
        yy = eval_dt.year % 100
        fname = f"S{yy:02d}{eval_dt.month:02d}{eval_dt.day:02d}"
        url = f"https://market.jsda.or.jp/shijyo/saiken/baibai/baisanchi/files/{eval_dt.year}/{fname}.xls"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        wb = xlrd.open_workbook(file_contents=data)
        ws = wb.sheet_by_index(0)
        best = None
        for r in range(1, ws.nrows):
            name = str(ws.cell_value(r, 2)).strip()
            if "長期国債" not in name or "超長期国債" in name:
                continue
            maturity_str = str(ws.cell_value(r, 3))
            med_compound = ws.cell_value(r, 11)
            if not med_compound:
                continue
            try:
                mat_dt = datetime.strptime(maturity_str, "%Y/%m/%d")
            except (ValueError, TypeError):
                continue
            diff = abs((mat_dt - exercise_end_dt).days)
            if best is None or diff < best[0]:
                best = (diff, name, mat_dt, float(med_compound))
        if best:
            result["name"] = best[1]
            result["maturity"] = best[2].strftime("%Y-%m-%d")
            result["yield_value"] = str(best[3])
    except Exception:
        pass
    return result


def fetch_yahoo_quote_data(ticker_code: str) -> dict:
    result = {"shares_outstanding": 0, "dividend_per_share": 0, "dividend_yield": 0.0}
    try:
        url = f"https://finance.yahoo.co.jp/quote/{ticker_code}.T"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8")
        m = re.search(
            r'発行済株式数.*?<span[^>]*class="StyledNumber__value[^"]*"[^>]*>([\d,]+)</span>',
            html, re.DOTALL)
        if m:
            result["shares_outstanding"] = int(m.group(1).replace(",", ""))
        m2 = re.search(
            r'配当利回り.*?<span[^>]*class="StyledNumber__value[^"]*"[^>]*>([\d.]+)</span>',
            html, re.DOTALL)
        if m2:
            result["dividend_yield"] = round(float(m2.group(1)), 2)
        m3 = re.search(
            r'1株配当.*?<span[^>]*class="StyledNumber__value[^"]*"[^>]*>([\d.]+)</span>',
            html, re.DOTALL)
        if m3:
            result["dividend_per_share"] = int(float(m3.group(1)))
    except Exception:
        pass
    if result["shares_outstanding"] == 0:
        try:
            ticker = yf.Ticker(f"{ticker_code}.T")
            info = ticker.info
            shares = info.get("sharesOutstanding", 0)
            if shares:
                result["shares_outstanding"] = int(shares)
        except Exception:
            pass
    return result


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(content_length))
            ticker_code = body["ticker_code"]
            eval_date = body["eval_date"]
            exercise_start = body.get("exercise_start", "")
            exercise_end = body.get("exercise_end", "")

            eval_dt = datetime.strptime(eval_date, "%Y-%m-%d")
            ticker = yf.Ticker(f"{ticker_code}.T")

            if exercise_end:
                ex_end_dt = datetime.strptime(exercise_end, "%Y-%m-%d")
            else:
                ex_end_dt = eval_dt + relativedelta(years=5)

            # 行使期間の開始日（未指定なら基準日を代用）
            ex_start_dt = datetime.strptime(exercise_start, "%Y-%m-%d") if exercise_start else eval_dt

            if ex_end_dt <= eval_dt:
                raise ValueError(
                    f"権利行使期間(終了) {ex_end_dt.strftime('%Y-%m-%d')} は "
                    f"評価基準日 {eval_dt.strftime('%Y-%m-%d')} より後の日付にしてください。"
                )
            # 行使期間の長さで変数取得期間を決める
            rd = relativedelta(ex_end_dt, ex_start_dt)
            months_to_maturity = rd.years * 12 + rd.months
            if rd.days > 0:
                months_to_maturity += 1
            days_to_maturity = (ex_end_dt - ex_start_dt).days

            # 株価
            start = (eval_dt - timedelta(days=10)).strftime("%Y-%m-%d")
            end = (eval_dt + timedelta(days=1)).strftime("%Y-%m-%d")
            hist_around = ticker.history(start=start, end=end)
            hist_before = hist_around[hist_around.index.strftime("%Y-%m-%d") <= eval_date]
            if len(hist_before) == 0:
                raise ValueError(f"評価基準日 {eval_date} の株価データが取得できません")
            close_val = hist_before["Close"].iloc[-1]
            if close_val is None or (isinstance(close_val, float) and math.isnan(close_val)):
                raise ValueError(f"評価基準日 {eval_date} の終値が NaN です（銘柄 {ticker_code}）")
            stock_price = int(close_val)

            # ボラティリティ: 基準日の前月末から (月数+1) 本の月次株価を取得
            vol_end_month = eval_dt.replace(day=1) - timedelta(days=1)
            vol_start_month = vol_end_month - relativedelta(months=months_to_maturity)
            vol_start = vol_start_month.replace(day=1).strftime("%Y-%m-%d")
            vol_end = (vol_end_month + timedelta(days=1)).strftime("%Y-%m-%d")
            hist_monthly = ticker.history(start=vol_start, end=vol_end, interval="1mo")
            returns = np.log(hist_monthly["Close"] / hist_monthly["Close"].shift(1)).dropna()
            annual_vol = returns.std(ddof=0) * np.sqrt(12)
            vol_start_label = f"{vol_start_month.year}年{vol_start_month.month}月"
            vol_end_label = f"{vol_end_month.year}年{vol_end_month.month}月"

            # 出来高
            volume_end = eval_dt
            volume_start = eval_dt - timedelta(days=days_to_maturity)
            hist_daily = ticker.history(
                start=volume_start.strftime("%Y-%m-%d"),
                end=(volume_end + timedelta(days=1)).strftime("%Y-%m-%d"),
            )
            if "Volume" not in hist_daily or len(hist_daily["Volume"].dropna()) == 0:
                raise ValueError(
                    f"出来高データが取得できません（銘柄 {ticker_code}, 期間 "
                    f"{volume_start.strftime('%Y-%m-%d')}〜{volume_end.strftime('%Y-%m-%d')}）"
                )
            median_raw = hist_daily["Volume"].median()
            if median_raw is None or (isinstance(median_raw, float) and math.isnan(median_raw)):
                raise ValueError(f"出来高中央値が NaN です（銘柄 {ticker_code}）")
            median_volume = int(median_raw)
            liquidity_shares = math.ceil(median_volume * 0.1)

            # 配当・発行済株式数
            yahoo_data = fetch_yahoo_quote_data(ticker_code)

            # リスクフリーレート (JSDA 長期国債、権利行使期間終了日に最も近い銘柄)
            jsda = fetch_jsda_bond(eval_dt, ex_end_dt) if exercise_end else {"name": "", "maturity": "", "yield_value": ""}

            result = {
                "stock_price": stock_price,
                "volatility": round(annual_vol * 100, 2),
                "vol_start_label": vol_start_label,
                "vol_end_label": vol_end_label,
                "median_volume": median_volume,
                "liquidity_shares": liquidity_shares,
                "volume_start": volume_start.strftime("%Y-%m-%d"),
                "volume_end": volume_end.strftime("%Y-%m-%d"),
                "dividend_yield": yahoo_data["dividend_yield"],
                "dividend_per_share": yahoo_data["dividend_per_share"],
                "shares_outstanding": yahoo_data["shares_outstanding"],
                "bond_name": jsda["name"],
                "bond_maturity": jsda["maturity"],
                "bond_yield": jsda["yield_value"],
            }

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(result).encode())

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())
