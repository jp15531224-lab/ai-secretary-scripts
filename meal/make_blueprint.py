#!/usr/bin/env python3
"""Make シナリオ「食事管理：記録・質問」の設計図(blueprint JSON)を作る。

  python3 make_blueprint.py --token-file <合言葉のファイル> --commit <コミットSHA> > blueprint.json

- 合言葉（ショートカットとの共有秘密）はリポジトリに置かない。デプロイ時にファイルから読む。
- 処理コード(dist/make_code.py)は GitHub の「コミットを固定した URL」から毎回読み込む
  （既存の定期ジョブと同じ方式。後から改ざんされても固定したコミット以外は実行されない）。
  コードを直したら: テスト → dist 再生成 → push → 新しいコミットSHAで設計図を作り直して scenarios_update。

流れ: Webhook →(合言葉確認) コード取得 → 食事DBを全部読む → シートID取得 → Python(prepare) → OpenAI → Python(apply)
      → ルーター［書き込みがあれば Sheets へ（追加は appendCells なので同時送信でも上書きしない）］→［返答］
- 「1件ずつ順番に処理」(sequential)をオンにすると Make が即「Accepted」を返して返答が届かないため、オフにしている。
"""
import argparse
import json

HOOK_ID = 3921217                            # Webhook「食事管理（iPhoneショートカット）」
SHEETS_CONN = 11338733                       # Google（Sheets 権限あり）
OPENAI_CONN = 11330538                       # OpenAI（ボイスメモのシナリオと同じ接続）
SPREADSHEET_ID = "1AzcV_y0rbwGPbGAyGLbriAq57k7QWRrDYgrRsxYO-W0"   # iCloud外部脳/食事・健康管理/食事DB
REPO_RAW = "https://raw.githubusercontent.com/jp15531224-lab/ai-secretary-scripts"
SHEET_RANGES = ["'食事'!A1:Y", "'補正履歴'!A1:G", "'確認候補'!A1:W", "'体重'!A1:E", "'設定'!A1:C"]
JSON_HEADER = [{"key": "Content-Type", "value": "application/json"}]


def pos(x, y=0):
    return {"designer": {"x": x, "y": y}}


def code_module(mid, mode, extra, x):
    inputs = [{"name": "mode", "value": mode}, {"name": "db", "value": "{{3.body}}"},
              {"name": "text", "value": "{{1.text}}"}, {"name": "image", "value": "{{1.image}}"},
              {"name": "source", "value": "{{1.source}}"}] + extra
    return {"id": mid, "module": "code:ExecuteCode", "version": 0, "parameters": {},
            "mapper": {"language": "python", "inputFormat": "string", "input": inputs,
                       "codeStringPython": "{{toString(2.data)}}"},
            "metadata": pos(x)}


def sheets_call(mid, path, body, x, y=0, flt=None):
    url = f"spreadsheets/{SPREADSHEET_ID}" + (path if path.startswith(":") else "/" + path)
    m = {"id": mid, "module": "google-sheets:makeAPICall", "version": 2, "parameters": {"__IMTCONN__": SHEETS_CONN},
         "mapper": {"url": url, "method": "POST", "headers": JSON_HEADER,
                    "qs": [], "body": body},
         "metadata": pos(x, y)}
    if flt:
        m["filter"] = flt
    return m


def blueprint(token, commit):
    read_body = json.dumps({"dataFilters": [{"a1Range": r} for r in SHEET_RANGES],
                            "valueRenderOption": "UNFORMATTED_VALUE"}, ensure_ascii=False)
    flow = [
        {"id": 1, "module": "gateway:CustomWebHook", "version": 1,
         "parameters": {"hook": HOOK_ID, "maxResults": 1}, "mapper": {}, "metadata": pos(0)},
        {"id": 2, "module": "http:ActionSendData", "version": 3, "parameters": {"handleErrors": True},
         "filter": {"name": "合言葉が一致", "conditions": [[{"a": "{{1.token}}", "o": "text:equal", "b": token}]]},
         "mapper": {"url": f"{REPO_RAW}/{commit}/meal/dist/make_code.py", "method": "get", "headers": [], "qs": [],
                    "bodyType": "", "parseResponse": False, "serializeUrl": False, "shareCookies": False,
                    "rejectUnauthorized": True, "followRedirect": True, "followAllRedirects": False,
                    "useQuerystring": False, "gzip": True, "useMtls": False},
         "metadata": pos(300)},
        sheets_call(3, "values:batchGetByDataFilter", read_body, 600),
        {"id": 10, "module": "google-sheets:makeAPICall", "version": 2, "parameters": {"__IMTCONN__": SHEETS_CONN},
         "mapper": {"url": f"spreadsheets/{SPREADSHEET_ID}", "method": "GET", "headers": [],
                    "qs": [{"key": "fields", "value": "sheets.properties(sheetId,title)"}], "body": ""},
         "metadata": pos(750, 150)},
        code_module(4, "prepare", [], 900),
        {"id": 5, "module": "openai-gpt-3:makeApiCall", "version": 1, "parameters": {"__IMTCONN__": OPENAI_CONN},
         "mapper": {"url": "/v1/chat/completions", "method": "POST", "headers": JSON_HEADER, "qs": [],
                    "body": "{{4.result.request_body}}"},
         "metadata": pos(1200)},
        code_module(6, "apply", [{"name": "openai", "value": "{{5.body}}"}, {"name": "meta", "value": "{{10.body}}"}], 1500),
        {"id": 7, "module": "builtin:BasicRouter", "version": 1, "mapper": None, "metadata": pos(1800),
         "routes": [
             {"flow": [sheets_call(8, ":batchUpdate", "{{6.result.write_body}}", 2100, -150,
                                   {"name": "書き込みあり",
                                    "conditions": [[{"a": "{{6.result.has_writes}}", "o": "boolean:equal", "b": "true"}]]})]},
             {"flow": [{"id": 9, "module": "gateway:WebhookRespond", "version": 1, "parameters": {},
                        "mapper": {"status": 200, "body": "{{6.result.reply}}",
                                   "headers": [{"key": "Content-Type", "value": "text/plain; charset=utf-8"}]},
                        "metadata": pos(2100, 150)}]},
         ]},
    ]
    return {"name": "食事管理：記録・質問（iPhoneショートカット）", "flow": flow,
            "metadata": {"instant": True, "version": 1,
                         "scenario": {"roundtrips": 1, "maxErrors": 3, "autoCommit": True, "autoCommitTriggerLast": True,
                                      "sequential": False, "confidential": True, "dataloss": False, "dlq": False,
                                      "freshVariables": False},
                         "designer": {"orphans": []}}}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--token-file", required=True)
    ap.add_argument("--commit", required=True, help="dist/make_code.py を含むコミットの SHA（40桁）")
    a = ap.parse_args()
    if len(a.commit) != 40:
        raise SystemExit("--commit は40桁の SHA を指定")
    with open(a.token_file, encoding="utf-8") as f:
        tok = f.read().strip()
    print(json.dumps(blueprint(tok, a.commit), ensure_ascii=False, indent=1))
