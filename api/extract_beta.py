"""
Vercel Serverless Function: βスクリーンショットからβ値を抽出
POST /api/extract-beta

SPEEDAのβ値画面スクリーンショットを受け取り、
Claude Vision APIでβ値を読み取って返す。
"""

import json
import os
import base64
import cgi
import urllib.request
import urllib.parse
from http.server import BaseHTTPRequestHandler


ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")


def extract_beta_from_image(image_bytes: bytes, content_type: str) -> dict:
    """Claude Vision API を使ってβ値を抽出"""
    if not ANTHROPIC_API_KEY:
        raise ValueError("ANTHROPIC_API_KEY が設定されていません")

    # 画像をbase64エンコード
    b64 = base64.b64encode(image_bytes).decode("utf-8")

    # media typeを判定
    if "png" in content_type:
        media_type = "image/png"
    elif "jpeg" in content_type or "jpg" in content_type:
        media_type = "image/jpeg"
    elif "webp" in content_type:
        media_type = "image/webp"
    elif "gif" in content_type:
        media_type = "image/gif"
    else:
        media_type = "image/png"

    payload = {
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 256,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": (
                            "この画像はSPEEDAのβ値画面のスクリーンショットです。\n"
                            "「Levered β」の数値を正確に読み取ってください。\n"
                            "Unlevered βではなく、Levered βの値です。\n"
                            "回答は数値のみ（例: 1.185）で返してください。余計な説明は不要です。"
                        ),
                    },
                ],
            }
        ],
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

    text = result["content"][0]["text"].strip()

    # 数値部分を抽出
    import re
    m = re.search(r'-?\d+\.?\d*', text)
    if m:
        beta = float(m.group())
        return {"beta": beta, "raw": text}
    else:
        raise ValueError(f"β値を抽出できませんでした: {text}")


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_type = self.headers.get("Content-Type", "")
            content_length = int(self.headers.get("Content-Length", 0))

            if "multipart/form-data" in content_type:
                environ = {
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": content_type,
                    "CONTENT_LENGTH": str(content_length),
                }
                form = cgi.FieldStorage(
                    fp=self.rfile,
                    headers=self.headers,
                    environ=environ,
                )
                file_item = form["file"]
                file_bytes = file_item.file.read()
                file_type = file_item.type or "image/png"
            else:
                file_bytes = self.rfile.read(content_length)
                file_type = content_type

            result = extract_beta_from_image(file_bytes, file_type)

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(result).encode())

        except Exception as e:
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error": str(e)}).encode())
