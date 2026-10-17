#!/usr/bin/env python3
"""Make シナリオ「食事管理：記録・質問」の設計図(blueprint JSON)を作る。

  python3 make_blueprint.py --token-file <トークンのファイル> > blueprint.json

トークン（ショートカットとの合言葉）はリポジトリに置かない。デプロイ時にファイルから読み込む。
流れ: Webhook → 食事DBを全部読む → Python(prepare) → OpenAI → Python(apply) → [書き込み] → 返答
"""
import argparse
import json

import build_make_code

TEAM_HOOK_ID = 3921217                       # Webhook「食事管理（iPhoneショートカット）」
SHEETS_CONN = 11338733                       # Google（Sheets 権限あり）
OPENAI_CONN = 11330538                       # OpenAI（ボイスメモのシナリオと同じ接続）
SPREADSHEET_ID = "1AzcV_y0rbwGPbGAyGLbriAq57k7QWRrDYgrRsxYO-W0"   # iCloud外部脳/食事・健康管理/食事DB
SHEET_RANGES = ["'食事'!A1:Y", "'補正履歴'!A1:G", "'確認候補'!A1:W", "'体重'!A1:E", "'設定'!A1:C"]


def code_module(mid, mode, extra, x):
    inputs = [{"name": "mode", "value": mode}, {"name": "db", "value": "{{2.body}}"},
              {"name": "text", "value": "{{1.text}}"}, {"name": "image", "value": "{{1.image}}"},
              {"name": "source", "value": "{{1.source}}"}] + extra
    return {"id": mid, "module": "code:ExecuteCode", "version": 0, "parameters": {},
            "mapper": {"language": "python", "inputFormat": "editor", "input": inputs,
                       "codeEditorPython": build_make_code.build()},
            "metadata": {"designer": {"x": x, "y": 0}}}


def blueprint(token):
    read_body = json.dumps({"dataFilters": [{"a1Range": r} for r in SHEET_RANGES],
                            "valueRenderOption": "UNFORMATTED_VALUE"}, ensure_ascii=False)
    flow = [
        {"id": 1, "module": "gateway:CustomWebHook", "version": 1,
         "parameters": {"hook": TEAM_HOOK_ID, "maxResults": 1}, "mapper": {},
         "metadata": {"designer": {"x": 0, "y": 0}}},
        {"id": 2, "module": "google-sheets:makeAPICall", "version": 2,
         "parameters": {"__IMTCONN__": SHEETS_CONN},
         "filter": {"name": "合言葉が一致", "conditions": [[{"a": "{{1.token}}", "o": "text:equal", "b": token}]]},
         "mapper": {"url": f"spreadsheets/{SPREADSHEET_ID}/values:batchGetByDataFilter", "method": "POST",
                    "headers": [{"key": "Content-Type", "value": "application/json"}], "qs": [], "body": read_body},
         "metadata": {"designer": {"x": 300, "y": 0}}},
        code_module(3, "prepare", [], 600),
        {"id": 4, "module": "openai-gpt-3:makeApiCall", "version": 1,
         "parameters": {"__IMTCONN__": OPENAI_CONN},
         "mapper": {"url": "/v1/chat/completions", "method": "POST",
                    "headers": [{"key": "Content-Type", "value": "application/json"}], "qs": [],
                    "body": "{{3.result.request_body}}"},
         "metadata": {"designer": {"x": 900, "y": 0}}},
        code_module(5, "apply", [{"name": "openai", "value": "{{4.body}}"}], 1200),
        {"id": 6, "module": "builtin:BasicRouter", "version": 1, "mapper": None,
         "metadata": {"designer": {"x": 1500, "y": 0}},
         "routes": [
             {"flow": [{"id": 7, "module": "google-sheets:makeAPICall", "version": 2,
                        "parameters": {"__IMTCONN__": SHEETS_CONN},
                        "filter": {"name": "書き込みあり", "conditions": [[{"a": "{{5.result.has_writes}}", "o": "boolean:equal", "b": "true"}]]},
                        "mapper": {"url": f"spreadsheets/{SPREADSHEET_ID}/values:batchUpdate", "method": "POST",
                                   "headers": [{"key": "Content-Type", "value": "application/json"}], "qs": [],
                                   "body": "{{5.result.write_body}}"},
                        "metadata": {"designer": {"x": 1800, "y": -150}}}]},
             {"flow": [{"id": 8, "module": "gateway:WebhookRespond", "version": 1, "parameters": {},
                        "mapper": {"status": 200, "body": "{{5.result.reply}}",
                                   "headers": [{"key": "Content-Type", "value": "text/plain; charset=utf-8"}]},
                        "metadata": {"designer": {"x": 1800, "y": 150}}}]},
         ]},
    ]
    return {"name": "食事管理：記録・質問（iPhoneショートカット）", "flow": flow,
            "metadata": {"instant": True, "version": 1,
                         "scenario": {"roundtrips": 1, "maxErrors": 3, "autoCommit": True, "autoCommitTriggerLast": True,
                                      "sequential": True, "confidential": True, "dataloss": False, "dlq": False,
                                      "freshVariables": False},
                         "designer": {"orphans": []}}}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--token-file", required=True)
    a = ap.parse_args()
    with open(a.token_file, encoding="utf-8") as f:
        tok = f.read().strip()
    print(json.dumps(blueprint(tok), ensure_ascii=False))
