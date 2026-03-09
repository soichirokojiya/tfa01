"""
Vercel Serverless Function: 発行要項テキストからAIで情報抽出
POST /api/extract_issuance
"""

import json
import os
import re
import urllib.request
from http.server import BaseHTTPRequestHandler

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")


def extract_from_text(text: str) -> dict:
    if not ANTHROPIC_API_KEY:
        raise ValueError("ANTHROPIC_API_KEY が設定されていません")

    prompt = (
        "以下は新株予約権の発行要項テキストです。次の項目を抽出してJSON形式で返してください。\n"
        "値が見つからない場合は空文字にしてください。\n\n"
        "抽出項目:\n"
        "- exercise_start: 権利行使期間の開始日（YYYY-MM-DD形式）\n"
        "- exercise_end: 権利行使期間の終了日（YYYY-MM-DD形式）\n"
        "- assignee: 割当先（複数の場合は改行区切り）\n"
        "- resolution_date: 決議日（YYYY-MM-DD形式、未定の場合は「未定」）\n"
        "- warrant_total: 新株予約権の総数（数字のみ、カンマなし）\n"
        "- issuable_shares: 行使による発行株式総数（数字のみ、カンマなし。見つからない場合は空文字）\n"
        "- special_terms: 査定に関連する特約条項（取得条項、行使条件等があれば原文を一字一句そのまま記載。要約・省略しないこと）\n"
        "- company_name: 発行会社名（「株式会社」を含む正式名称）\n\n"
        "JSONのみ返してください。説明は不要です。\n\n"
        "---\n"
        f"{text}\n"
        "---"
    )

    payload = {
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 2048,
        "messages": [{"role": "user", "content": prompt}],
    }

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )

    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode("utf-8"))

    response_text = result["content"][0]["text"].strip()

    # JSONブロックを抽出
    m = re.search(r'\{[\s\S]*\}', response_text)
    if not m:
        raise ValueError(f"JSON抽出失敗: {response_text[:200]}")
    data = json.loads(m.group())

    # 会社名から上場コードを検索
    company = data.get("company_name", "")
    if company:
        ticker_code = lookup_ticker(company)
        if ticker_code:
            data["ticker_code"] = ticker_code

    return data


def lookup_ticker(company_name: str) -> str:
    """Yahoo Finance で会社名を検索して上場コードを返す"""
    try:
        # 「株式会社」を除いた短縮名で検索
        short = company_name.replace("株式会社", "").strip()
        url = f"https://finance.yahoo.co.jp/search/?query={urllib.request.quote(short)}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8")
        # 検索結果から最初の銘柄コードを取得
        m = re.search(r'/quote/(\d{4})\.T', html)
        if m:
            return m.group(1)
    except Exception:
        pass
    return ""


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(content_length))
            text = body.get("text", "")
            if not text:
                raise ValueError("テキストが空です")

            result = extract_from_text(text)

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(result, ensure_ascii=False).encode())

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())
