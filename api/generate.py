"""
Vercel Serverless Function: 新株予約権評価報告書を生成して返す
POST /api/generate
"""

import json
import os
import re
import math
import urllib.request
import urllib.parse
import tempfile
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from http.server import BaseHTTPRequestHandler
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
import yfinance as yf
import numpy as np
import warnings
warnings.filterwarnings("ignore")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "..", "template.docx")


# ──────────────────────────────────────────────
# データ取得
# ──────────────────────────────────────────────

def fetch_japanese_company_name(ticker_code: str) -> str:
    url = f"https://finance.yahoo.co.jp/quote/{ticker_code}.T"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        html = resp.read().decode("utf-8")
    m = re.search(r"<title>(.*?)【\d+】", html)
    if not m:
        raise ValueError(f"社名を取得できませんでした: {ticker_code}")
    name = m.group(1).strip()
    name = re.sub(r"^\(株\)", "", name)
    name = re.sub(r"\(株\)$", "", name)
    name = re.sub(r"^株式会社", "", name)
    name = re.sub(r"株式会社$", "", name)
    return name.strip()


def fetch_company_profile(ticker_code: str) -> dict:
    url = f"https://finance.yahoo.co.jp/quote/{ticker_code}.T/profile"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        html = resp.read().decode("utf-8")

    def extract(label):
        m = re.search(rf'<th[^>]*>{label}</th>\s*<td[^>]*>(.*?)</td>', html, re.DOTALL)
        if m:
            return re.sub(r'<[^>]+>', '', m.group(1)).strip()
        return ""

    address = ""
    m = re.search(r'〒[\d\-]+\s*(.+?)(?=<|")', html)
    if m:
        address = m.group(1).strip()

    return {
        "representative": extract("代表者名"),
        "address": address,
        "established": extract("設立年月日"),
        "settlement": extract("決算"),
    }


def fetch_stock_data(ticker_code: str, eval_date: str):
    ticker = yf.Ticker(f"{ticker_code}.T")
    eval_dt = datetime.strptime(eval_date, "%Y-%m-%d")

    start = (eval_dt - timedelta(days=10)).strftime("%Y-%m-%d")
    end = (eval_dt + timedelta(days=1)).strftime("%Y-%m-%d")
    hist_around = ticker.history(start=start, end=end)
    hist_before = hist_around[hist_around.index.strftime("%Y-%m-%d") <= eval_date]
    if len(hist_before) == 0:
        raise ValueError(f"評価基準日 {eval_date} の株価データが取得できません")
    stock_price = int(hist_before["Close"].iloc[-1])

    vol_end_month = eval_dt.replace(day=1) - timedelta(days=1)
    vol_start_month = vol_end_month - relativedelta(years=5)
    vol_start = vol_start_month.replace(day=1).strftime("%Y-%m-%d")
    vol_end = (vol_end_month + timedelta(days=1)).strftime("%Y-%m-%d")
    hist_monthly = ticker.history(start=vol_start, end=vol_end, interval="1mo")
    returns = np.log(hist_monthly["Close"] / hist_monthly["Close"].shift(1)).dropna()
    annual_vol = returns.std() * np.sqrt(12)
    vol_start_label = f"{vol_start_month.year}年{vol_start_month.month}月"
    vol_end_label = f"{vol_end_month.year}年{vol_end_month.month}月"

    report_date = eval_dt
    volume_end = eval_dt - timedelta(days=1)
    volume_start = volume_end - relativedelta(years=5)
    hist_daily = ticker.history(
        start=volume_start.strftime("%Y-%m-%d"),
        end=(volume_end + timedelta(days=1)).strftime("%Y-%m-%d"),
    )
    median_volume = int(hist_daily["Volume"].median())
    liquidity_shares = math.ceil(median_volume * 0.1)

    dividends = ticker.dividends
    dividend_per_share = 0
    if len(dividends) > 0:
        divs_before = dividends[dividends.index.strftime("%Y-%m-%d") <= eval_date]
        if len(divs_before) > 0:
            dividend_per_share = int(divs_before.iloc[-1])
    dividend_yield = round((dividend_per_share / stock_price * 100), 2) if stock_price > 0 else 0.0

    shares_outstanding = ticker.info.get("sharesOutstanding", 0)

    return {
        "stock_price": stock_price,
        "volatility": round(annual_vol * 100, 2),
        "vol_start_label": vol_start_label,
        "vol_end_label": vol_end_label,
        "median_daily_volume": median_volume,
        "liquidity_shares": liquidity_shares,
        "volume_start_date": volume_start,
        "volume_end_date": volume_end,
        "dividend_yield": dividend_yield,
        "dividend_per_share": dividend_per_share,
        "report_date": report_date,
        "shares_outstanding": shares_outstanding,
    }


# ──────────────────────────────────────────────
# docx 操作
# ──────────────────────────────────────────────

def insert_paragraph_after(paragraph, text, font_name="ＭＳ Ｐ明朝"):
    new_p = deepcopy(paragraph._element)
    for r in new_p.findall(qn('w:r')):
        new_p.remove(r)
    run_elem = paragraph._element.makeelement(qn('w:r'), {})
    rPr = run_elem.makeelement(qn('w:rPr'), {})
    rFonts = rPr.makeelement(qn('w:rFonts'), {
        qn('w:ascii'): font_name,
        qn('w:eastAsia'): font_name,
        qn('w:hAnsi'): font_name,
    })
    rPr.append(rFonts)
    run_elem.append(rPr)
    t_elem = run_elem.makeelement(qn('w:t'), {})
    t_elem.text = text
    run_elem.append(t_elem)
    new_p.append(run_elem)
    paragraph._element.addnext(new_p)
    return new_p


def replace_in_runs(paragraph, old_text, new_text):
    full_text = "".join(run.text for run in paragraph.runs)
    if old_text not in full_text:
        return False
    new_full = full_text.replace(old_text, new_text)
    runs = paragraph.runs
    if not runs:
        return False
    remaining = new_full
    for run in runs:
        run_len = len(run.text)
        if len(remaining) <= 0:
            run.text = ""
        elif len(remaining) <= run_len:
            run.text = remaining
            remaining = ""
        else:
            run.text = remaining[:run_len]
            remaining = remaining[run_len:]
    if remaining:
        runs[-1].text += remaining
    return True


def replace_in_document(doc, old_text, new_text):
    count = 0
    for para in doc.paragraphs:
        if replace_in_runs(para, old_text, new_text):
            count += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if replace_in_runs(para, old_text, new_text):
                        count += 1
    return count


def fmt_date_jp(dt):
    return f"{dt.year}年{dt.month}月{dt.day}日"


# ──────────────────────────────────────────────
# Handler
# ──────────────────────────────────────────────

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(content_length))
            ticker_code = body["ticker_code"]
            eval_date = body["eval_date"]
            eval_dt = datetime.strptime(eval_date, "%Y-%m-%d")

            # TOPページ入力項目
            exercise_start = body.get("exercise_start", "")
            exercise_end = body.get("exercise_end", "")
            resolution_date = body.get("resolution_date", "")
            warrant_total = body.get("warrant_total", "")
            issuable_shares = body.get("issuable_shares", "")
            fair_value_str = body.get("fair_value_per_share", "")
            special_terms = body.get("special_terms", "")
            market_risk_premium = body.get("market_risk_premium", "")
            beta_val = body.get("beta", "")

            fair_value_per_share = float(fair_value_str) if fair_value_str else None

            # データ取得
            company_name_jp = fetch_japanese_company_name(ticker_code)
            profile = fetch_company_profile(ticker_code)
            data = fetch_stock_data(ticker_code, eval_date)

            # テンプレート読み込み
            doc = Document(TEMPLATE_PATH)

            # 置換
            replacements = [
                ("ジェリービーンズグループ", company_name_jp),
                ("3070", ticker_code),
                ("110円", f"{data['stock_price']}円"),
                ("62.54%", f"{data['volatility']}%"),
                ("2021年2月- 2026年2月",
                 f"{data['vol_start_label']}- {data['vol_end_label']}"),
                ("2021年3月3日から2026年3月2日",
                 f"{fmt_date_jp(data['volume_start_date'])}から{fmt_date_jp(data['volume_end_date'])}"),
                ("1,483,123", f"{data['median_daily_volume']:,}"),
                ("148,313", f"{data['liquidity_shares']:,}"),
                ("0%（0円/株）",
                 f"{data['dividend_yield']}%（{data['dividend_per_share']}円/株）"),
                ("2026年3月2日", fmt_date_jp(eval_dt)),
                ("79,440,000", f"{data['shares_outstanding']:,}"),
                ("宮崎明", profile['representative'].replace("\u3000", "")),
                ("宮崎\u3000明", profile['representative']),
                ("東京都台東区上野1-16-5", profile['address']),
                ("1990年4月", profile['established'].replace("10日", "").rstrip("日")),
                ("1月末", profile['settlement'].replace("日", "")),
            ]

            # CAPM各変数（テーブル内の個別セル）
            if market_risk_premium:
                replacements.append(("9.3%", f"{market_risk_premium}%"))
            if beta_val:
                replacements.append(("0.567", beta_val))

            # 権利行使期間
            if exercise_start and exercise_end:
                ex_start_dt = datetime.strptime(exercise_start, "%Y-%m-%d")
                ex_end_dt = datetime.strptime(exercise_end, "%Y-%m-%d")
                replacements.append(("2026年3月3日-", f"{fmt_date_jp(ex_start_dt)}-"))
                replacements.append(("2026年3月4日", fmt_date_jp(ex_end_dt)))

            # ●プレースホルダー（テーブル1）
            if resolution_date:
                res_dt = datetime.strptime(resolution_date, "%Y-%m-%d")
                replacements.append(("2026年●月●日", fmt_date_jp(res_dt)))
            if warrant_total:
                replacements.append(("●個", f"{warrant_total}個"))
            if issuable_shares:
                replacements.append(("●株", f"{issuable_shares}株"))
            # 行使による払込価額 = 株価と同額
            replacements.append(("●円", f"{data['stock_price']}円"))

            # 公正価値 → 株価比率
            if fair_value_per_share is not None:
                price_ratio = round(fair_value_per_share / data['stock_price'] * 100, 2)
                fair_value_per_unit = round(fair_value_per_share * 100)
                replacements.append(("公正価値113円", f"公正価値{fair_value_per_unit:,}円"))
                replacements.append(("1.13円/株", f"{fair_value_per_share}円/株"))
                replacements.append(("当初株価の1.03%", f"当初株価の{price_ratio:.2f}%"))

            for old, new in replacements:
                replace_in_document(doc, old, new)

            # 一時ファイルに保存して返す
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                doc.save(tmp.name)
                tmp_path = tmp.name

            with open(tmp_path, "rb") as f:
                docx_bytes = f.read()
            os.unlink(tmp_path)

            eval_ym = eval_dt.strftime("%Y%m")
            filename = f"{eval_ym}_新株予約権評価報告書_株式会社{company_name_jp}.docx"

            self.send_response(200)
            self.send_header("Content-Type",
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition",
                             f"attachment; filename*=UTF-8''{urllib.parse.quote(filename)}")
            self.send_header("Content-Length", str(len(docx_bytes)))
            self.end_headers()
            self.wfile.write(docx_bytes)

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())
