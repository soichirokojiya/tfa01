"""
Vercel Serverless Function: 上場コードから会社名を取得
GET /api/lookup?code=3070
"""

import json
import re
import urllib.request
from http.server import BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs


def fetch_company_name(ticker_code: str) -> str:
    url = f"https://finance.yahoo.co.jp/quote/{ticker_code}.T"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        html = resp.read().decode("utf-8")
    m = re.search(r"<title>(.*?)【\d+】", html)
    if not m:
        return ""
    name = m.group(1).strip()
    # 「(株)」「株式会社」は残す（表示用）
    return name


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            params = parse_qs(urlparse(self.path).query)
            code = params.get("code", [""])[0].strip()

            if not code:
                self.send_response(400)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "code parameter required"}).encode())
                return

            name = fetch_company_name(code)

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps({"name": name, "code": code}).encode())

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())
