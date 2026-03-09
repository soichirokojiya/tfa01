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
from docx.shared import Pt
from docx.text.paragraph import Paragraph
import yfinance as yf
import numpy as np
import warnings
warnings.filterwarnings("ignore")

# Vercel: テンプレートファイルのパスを複数候補から解決
_candidates = [
    os.path.join(os.path.dirname(__file__), "..", "template.docx"),
    os.path.join(os.path.dirname(__file__), "template.docx"),
    "/var/task/template.docx",
]
TEMPLATE_PATH = next((p for p in _candidates if os.path.exists(p)), _candidates[0])


def fetch_yahoo_quote_data(ticker_code: str) -> dict:
    """Yahoo Finance Japan から発行済株式数・配当情報を取得"""
    result = {"shares_outstanding": 0, "dividend_per_share": 0, "dividend_yield": 0.0}
    try:
        url = f"https://finance.yahoo.co.jp/quote/{ticker_code}.T"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8")
        # HTML から直接スクレイピング
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

    # フォールバック: yfinance から発行済株式数を取得
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


# ──────────────────────────────────────────────
# データ取得
# ──────────────────────────────────────────────

def fetch_japanese_company_name(ticker_code: str) -> str:
    # 1) Yahoo Finance Japan Webページ
    try:
        url = f"https://finance.yahoo.co.jp/quote/{ticker_code}.T"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8")
        m = re.search(r"<title>(.*?)【\d+】", html)
        if m:
            name = m.group(1).strip()
            name = re.sub(r"^\(株\)", "", name)
            name = re.sub(r"\(株\)$", "", name)
            name = re.sub(r"^株式会社", "", name)
            name = re.sub(r"株式会社$", "", name)
            return name.strip()
    except Exception:
        pass

    # 2) フォールバック: yfinance API
    try:
        ticker = yf.Ticker(f"{ticker_code}.T")
        info = ticker.info
        name = info.get("shortName", "") or info.get("longName", "")
        name = re.sub(r"^\(株\)", "", name)
        name = re.sub(r"\(株\)$", "", name)
        name = re.sub(r"^株式会社", "", name)
        name = re.sub(r"株式会社$", "", name)
        return name.strip()
    except Exception:
        pass

    raise ValueError(f"社名を取得できませんでした: {ticker_code}")


def fetch_company_profile(ticker_code: str) -> dict:
    profile = {"representative": "", "address": "", "established": "", "settlement": ""}

    # 1) Yahoo Finance Japan Webページ
    try:
        url = f"https://finance.yahoo.co.jp/quote/{ticker_code}.T/profile"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8")

        def extract(label):
            m = re.search(rf'<th[^>]*>{label}</th>\s*<td[^>]*>(.*?)</td>', html, re.DOTALL)
            if m:
                return re.sub(r'<[^>]+>', '', m.group(1)).strip()
            return ""

        m = re.search(r'〒[\d\-]+\s*(.+?)(?=<|")', html)
        profile["representative"] = extract("代表者名")
        profile["address"] = m.group(1).strip() if m else ""
        profile["established"] = extract("設立年月日")
        profile["settlement"] = extract("決算")
    except Exception:
        pass

    return profile


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
    annual_vol = returns.std(ddof=0) * np.sqrt(12)
    vol_start_label = f"{vol_start_month.year}年{vol_start_month.month}月"
    vol_end_label = f"{vol_end_month.year}年{vol_end_month.month}月"

    report_date = eval_dt
    volume_end = eval_dt - timedelta(days=1)
    volume_start = eval_dt - relativedelta(years=5)
    hist_daily = ticker.history(
        start=volume_start.strftime("%Y-%m-%d"),
        end=(volume_end + timedelta(days=1)).strftime("%Y-%m-%d"),
    )
    median_volume = int(hist_daily["Volume"].median())
    liquidity_shares = math.ceil(median_volume * 0.1)

    # Yahoo Finance Japan から配当・発行済株式数を取得
    yahoo_data = fetch_yahoo_quote_data(ticker_code)
    dividend_per_share = yahoo_data["dividend_per_share"]
    dividend_yield = yahoo_data["dividend_yield"]
    shares_outstanding = yahoo_data["shares_outstanding"]

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

def insert_paragraph_after(paragraph, text, font_name="ＭＳ Ｐ明朝", font_size=11):
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
    sz = rPr.makeelement(qn('w:sz'), {qn('w:val'): str(font_size * 2)})
    szCs = rPr.makeelement(qn('w:szCs'), {qn('w:val'): str(font_size * 2)})
    rPr.append(sz)
    rPr.append(szCs)
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
            warrant_total = body.get("warrant_total", "").replace(",", "")
            issuable_shares = body.get("issuable_shares", "").replace(",", "")
            fair_value_str = body.get("fair_value_per_share", "")
            special_terms = body.get("special_terms", "")
            market_risk_premium = body.get("market_risk_premium", "")
            default_rate = body.get("default_rate", "")
            credit_cost_input = body.get("credit_cost", "")
            bond_name = body.get("bond_name", "")
            bond_maturity = body.get("bond_maturity", "")
            bond_yield = body.get("bond_yield", "")
            beta_value = body.get("beta", "")
            volatility_override = body.get("volatility", "")
            vol_start_override = body.get("vol_start_label", "")
            vol_end_override = body.get("vol_end_label", "")
            median_volume_override = body.get("median_volume", "")
            volume_start_override = body.get("volume_start", "")
            volume_end_override = body.get("volume_end", "")

            fair_value_per_share = float(fair_value_str) if fair_value_str else None

            # データ取得
            company_name_jp = fetch_japanese_company_name(ticker_code)
            profile = fetch_company_profile(ticker_code)
            data = fetch_stock_data(ticker_code, eval_date)

            # テンプレート読み込み
            if not os.path.exists(TEMPLATE_PATH):
                raise FileNotFoundError(f"テンプレートが見つかりません: {TEMPLATE_PATH}")
            doc = Document(TEMPLATE_PATH)

            # 置換
            # SPEEDAデータからの上書き
            vol_pct = float(volatility_override) if volatility_override else data['volatility']
            vol_start_lbl = vol_start_override if vol_start_override else data['vol_start_label']
            vol_end_lbl = vol_end_override if vol_end_override else data['vol_end_label']

            # 出来高: SPEEDAデータ優先
            if median_volume_override:
                median_vol = int(median_volume_override)
                liquidity = math.ceil(median_vol * 0.1)
            else:
                median_vol = data['median_daily_volume']
                liquidity = data['liquidity_shares']

            if volume_start_override and volume_end_override:
                vol_s_dt = datetime.strptime(volume_start_override, "%Y-%m-%d")
                vol_e_dt = datetime.strptime(volume_end_override, "%Y-%m-%d")
                volume_period = f"{fmt_date_jp(vol_s_dt)}から{fmt_date_jp(vol_e_dt)}"
            else:
                volume_period = f"{fmt_date_jp(data['volume_start_date'])}から{fmt_date_jp(data['volume_end_date'])}"

            # 権利行使期間（他の日付置換より先に実行）
            replacements = []
            if exercise_start and exercise_end:
                ex_start_dt = datetime.strptime(exercise_start, "%Y-%m-%d")
                ex_end_dt = datetime.strptime(exercise_end, "%Y-%m-%d")
                replacements.append(("2026年3月3日-", f"{fmt_date_jp(ex_start_dt)}-"))
                replacements.append(("2026年3月4日", fmt_date_jp(ex_end_dt)))

            replacements += [
                ("ジェリービーンズグループ", company_name_jp),
                ("3070", ticker_code),
                ("110円", f"{data['stock_price']:,}円"),
                ("62.54%", f"{vol_pct}%"),
                ("2021年2月- 2026年2月",
                 f"{vol_start_lbl}- {vol_end_lbl}"),
                ("2021年3月3日から2026年3月2日", volume_period),
                ("1,483,123", f"{median_vol:,}"),
                ("148,313", f"{liquidity:,}"),
                ("0%（0円/株）",
                 f"{data['dividend_yield']}%（{data['dividend_per_share']}円/株）"),
                ("2026年3月2日", fmt_date_jp(eval_dt)),
                ("79,440,000", f"{data['shares_outstanding']:,}"),
                ("宮崎明", profile['representative'].replace("\u3000", "")),
                ("宮崎\u3000明", profile['representative']),
                ("東京都台東区上野1-16-5", profile['address']),
                ("1990年4月", re.sub(r'\d{1,2}日$', '', profile['established'])),
                ("1月末", profile['settlement'].replace("日", "")),
            ]

            # 売買参考統計値から抽出された国債情報
            if bond_name:
                replacements.append(("長期国債362", bond_name))
            if bond_maturity:
                bm_dt = datetime.strptime(bond_maturity, "%Y-%m-%d")
                replacements.append(("2031年3月20日", fmt_date_jp(bm_dt)))
            if bond_yield:
                replacements.append(("1.591%", f"{bond_yield}%"))

            # CAPM各変数（テーブル内の個別セル）
            if market_risk_premium:
                replacements.append(("9.3%", f"{market_risk_premium}%"))

            # 対指数β
            if beta_value:
                replacements.append(("0.567", str(beta_value)))

            # デフォルト率
            default_rate_num = float(default_rate) if default_rate else 17.92
            if default_rate and default_rate != "17.92":
                replacements.append(("17.92%", f"{default_rate_num}%"))
                # 回収率: max(59.1 - 8.356 * デフォルト率, 0)
                recovery = max(59.1 - 8.356 * default_rate_num, 0)
                recovery_str = f"{recovery:.1f}%" if recovery > 0 else "0%"
                replacements.append(("0%\n", f"{recovery_str}\n"))

            # クレジットコスト
            if credit_cost_input and credit_cost_input != "21.83":
                replacements.append(("21.83%", f"{credit_cost_input}%"))

            # CAPM計算式の自動計算
            rfr = float(bond_yield) if bond_yield else 1.591
            mrp = float(market_risk_premium) if market_risk_premium else 9.3
            beta_num = float(beta_value) if beta_value else 0.567
            credit_cost = float(credit_cost_input) if credit_cost_input else 21.83
            capm_result = round(rfr + mrp * beta_num + credit_cost, 2)
            replacements.append((
                "= 1.591% + 9.3%\u00d7 0.567 + 21.83%",
                f"= {rfr}% + {mrp}%\u00d7 {beta_num} + {credit_cost}%"
            ))
            replacements.append(("= 28.69%", f"= {capm_result}%"))

            # ●プレースホルダー（テーブル1）
            if resolution_date:
                try:
                    res_dt = datetime.strptime(resolution_date, "%Y-%m-%d")
                    replacements.append(("2026年●月●日", fmt_date_jp(res_dt)))
                except ValueError:
                    # 「未定」等のテキストの場合はそのまま置換
                    replacements.append(("2026年●月●日", resolution_date))
            if warrant_total:
                replacements.append(("●個", f"{int(warrant_total):,}個"))
            if issuable_shares:
                replacements.append(("●株", f"{int(issuable_shares):,}株"))
            # 行使による払込価額 = 株価と同額
            replacements.append(("●円", f"{data['stock_price']:,}円"))

            # 公正価値 → 株価比率
            if fair_value_per_share is not None:
                price_ratio = round(fair_value_per_share / data['stock_price'] * 100, 2)
                fair_value_per_unit = round(fair_value_per_share * 100)
                replacements.append(("公正価値113円", f"公正価値{fair_value_per_unit:,}円"))
                replacements.append(("1.13円/株", f"{fair_value_per_share}円/株"))
                replacements.append(("当初株価の1.03%", f"当初株価の{price_ratio:.2f}%"))

            for old, new in replacements:
                replace_in_document(doc, old, new)

            # 査定に関連する特約条項（Table1 R5 C1）
            if special_terms:
                try:
                    cell = doc.tables[1].rows[5].cells[1]
                    # 既存段落のテキストをクリアして新しいテキストを設定
                    for i, para in enumerate(cell.paragraphs):
                        for run in para.runs:
                            run.text = ""
                    # 最初の段落に特約事項テキストを設定（フォントサイズ11pt）
                    lines = special_terms.split("\n")
                    if cell.paragraphs and cell.paragraphs[0].runs:
                        cell.paragraphs[0].runs[0].text = lines[0]
                        cell.paragraphs[0].runs[0].font.size = Pt(11)
                    else:
                        cell.paragraphs[0].text = lines[0]
                    # 残りの行は新規段落追加
                    for line in lines[1:]:
                        insert_paragraph_after(cell.paragraphs[-1], line)
                except Exception:
                    pass

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
