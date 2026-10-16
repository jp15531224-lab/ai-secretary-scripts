#!/usr/bin/env python3
"""AI秘書: 抽出JSON → Obsidian Markdown 生成ツール（標準ライブラリのみ）

使い方:
  python3 render_notes.py render <extraction.json> [<extraction2.json> ...] --out <dir>
      → <dir>/manifest.json と <dir>/<フォルダ>/<ファイル>.md を生成
  python3 render_notes.py validate <extraction.json>
      → 形式チェックのみ（エラーがあれば exit 1）
  python3 render_notes.py jst <plaud start_at>
      → Plaud の start_at(UTC) を JST に変換して表示

設計方針:
  - 既存ファイルは絶対に上書きしない。manifest の mode は常に "create_if_absent"。
    呼び出し側(Routine)は同名ファイルが Drive に既にあれば作成をスキップする。
  - 議事録ノートは録音1件につき1ファイル（ファイル名に出典IDを含めて重複防止）。
  - 人物・案件・会社ノートは初回だけ作る「索引ノート」。以後の情報は議事録側から
    [[リンク]] するので、Obsidian のバックリンク / Dataview で自動集約される。
  - 聞き取りが不確かな固有名詞はリンク化しない（誤った人物ノートを増やさない）。
  - 電話番号・メール・口座番号らしき数字列はマスクする（情報漏えい対策）。
  - 議事録は「事業」「月別」のハブノートにリンクする（グラフビューで事業・時期ごとに束ねる）。
    ハブノートも create_if_absent。source.transcript_note があれば文字起こし全文ノートへもリンクする。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

VERSION = "1.3.0"
JST = timezone(timedelta(hours=9))
WEEKDAYS = "月火水木金土日"

FOLDERS = {
    "minutes": "議事録",
    "person": "人物",
    "case": "案件",
    "company": "会社",
    "digest": "処理ログ",
    "business": "事業",
    "month": "月別",
}
BUSINESSES = {"グリーン興産", "ベルフィーヌ", "AI防犯カメラ", "全厚済", "新規事業", "プライベート", "不明"}
SOURCE_TYPES = {"plaud", "voicememo", "ainote", "drive"}
SOURCE_LABEL = {"plaud": "Plaud", "voicememo": "iPhoneボイスメモ", "ainote": "AINOTE", "drive": "Drive"}
BASIS = {"文字起こし確認済", "要約のみ", "文字起こし一部確認"}
SENSITIVITY = {"通常", "機密"}
PRIORITIES = {"P1", "P2", "P3", "P4"}
SELF_COMPANIES = {"グリーン興産", "ベルフィーヌ", "全厚済", "全国福利厚生共済会"}  # 自社は取引先ノートにしない
SELF_NAMES = {"本人", "石原", "石原さん", "淳平", "石原淳平"}
SCHEDULE_STATUS = {"確定", "要確認"}

# ファイル名に使えない / Obsidian リンクを壊す文字
_BAD_NAME_CHARS = re.compile(r'[\\/:*?"<>|#^\[\]\n\r\t]')
_HONORIFICS = re.compile(r"(さん|様|さま|氏|社長|専務|常務|部長|課長|先生|くん|ちゃん)$")

# マスク対象（誤検知より漏えい防止を優先）
_MASKS = [
    (re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+"), "[メール伏せ字]"),
    (re.compile(r"0\d{1,4}-\d{1,4}-\d{3,4}"), "[電話番号伏せ字]"),
    (re.compile(r"(?<!\d)0[789]0\d{8}(?!\d)"), "[電話番号伏せ字]"),
    (re.compile(r"(口座|口座番号)[^\d]{0,6}\d{6,8}"), r"\1[口座番号伏せ字]"),
    (re.compile(r"(?<!\d)\d{4}[ -]?\d{4}[ -]?\d{4}(?!\d)"), "[番号伏せ字]"),  # マイナンバー/カード等
]


class ValidationError(ValueError):
    pass


# ---------------------------------------------------------------- helpers
def plaud_start_to_jst(start_at: str) -> str:
    """Plaud の start_at（タイムゾーン無し UTC）を JST ISO8601 に変換。"""
    s = start_at.strip().replace("Z", "")
    dt = datetime.fromisoformat(s)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(JST).isoformat(timespec="seconds")


def mask(text: str) -> str:
    if not text:
        return text
    for pat, rep in _MASKS:
        text = pat.sub(rep, text)
    return text


def safe_name(name: str, limit: int = 40) -> str:
    n = _BAD_NAME_CHARS.sub("", (name or "").strip())
    n = re.sub(r"\s+", " ", n).strip(" .")
    return n[:limit]


def person_key(name: str) -> str:
    """人物ノート名: 敬称を外した氏名。"""
    return safe_name(_HONORIFICS.sub("", safe_name(name)))


def short_id(source_type: str, source_id: str) -> str:
    sid = re.sub(r"^of_", "", source_id)
    return f"{source_type}-{sid[:8]}"


def parse_recorded_at(value: str) -> datetime:
    dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        raise ValidationError("source.recorded_at にタイムゾーン(+09:00)がありません")
    return dt.astimezone(JST)


def yaml_str(v) -> str:
    return json.dumps("" if v is None else str(v), ensure_ascii=False)


def yaml_list(items) -> str:
    return "[" + ", ".join(yaml_str(i) for i in items) + "]"


# ---------------------------------------------------------------- validation
def validate(d: dict) -> list[str]:
    errs: list[str] = []
    src = d.get("source") or {}
    for k in ("type", "id", "title", "recorded_at", "basis"):
        if not src.get(k):
            errs.append(f"source.{k} が必要です")
    if src.get("type") and src["type"] not in SOURCE_TYPES:
        errs.append(f"source.type は {sorted(SOURCE_TYPES)} のいずれか")
    if src.get("basis") and src["basis"] not in BASIS:
        errs.append(f"source.basis は {sorted(BASIS)} のいずれか")
    if src.get("recorded_at"):
        try:
            parse_recorded_at(src["recorded_at"])
        except (ValueError, ValidationError) as e:
            errs.append(f"source.recorded_at 不正: {e}")
    if d.get("business", "不明") not in BUSINESSES:
        errs.append(f"business は {sorted(BUSINESSES)} のいずれか")
    if d.get("sensitivity", "通常") not in SENSITIVITY:
        errs.append("sensitivity は 通常/機密")
    if not d.get("summary"):
        errs.append("summary が必要です")
    for i, t in enumerate(d.get("todos", [])):
        if not t.get("task"):
            errs.append(f"todos[{i}].task が必要です")
        if t.get("priority") and t["priority"] not in PRIORITIES:
            errs.append(f"todos[{i}].priority は P1〜P4")
        if t.get("due"):
            try:
                datetime.strptime(t["due"], "%Y-%m-%d")
            except ValueError:
                errs.append(f"todos[{i}].due は YYYY-MM-DD（相対表現は不可）")
    for i, s in enumerate(d.get("schedules", [])):
        if s.get("status") not in SCHEDULE_STATUS:
            errs.append(f"schedules[{i}].status は 確定/要確認")
        if s.get("status") == "確定":
            if not (s.get("date") and s.get("start")):
                errs.append(f"schedules[{i}] 確定には date と start が必要")
            if not s.get("evidence"):
                errs.append(f"schedules[{i}] 確定には evidence（原文引用）が必要")
            if src.get("basis") == "要約のみ":
                errs.append(f"schedules[{i}] 要約のみの記録から『確定』にはできません")
    for i, r in enumerate(d.get("related_notes", [])):
        if not isinstance(r, dict) or not r.get("path"):
            errs.append(f"related_notes[{i}].path が必要です")
        elif not str(r["path"]).startswith("AI秘書/議事録/"):
            errs.append(f"related_notes[{i}].path は AI秘書/議事録/ 配下のみ")
        elif not r.get("reason"):
            errs.append(f"related_notes[{i}].reason（共通する会社・案件など）が必要です")
    for key in ("people", "companies", "cases"):
        for i, p in enumerate(d.get(key, [])):
            if not p.get("name"):
                errs.append(f"{key}[{i}].name が必要です")
    return errs


# ---------------------------------------------------------------- rendering
def _link(folder_key: str, name: str) -> str:
    return f"[[{FOLDERS[folder_key]}/{name}|{name}]]"


def _entity_ref(folder_key: str, ent: dict, honorific: str = "") -> str:
    name = person_key(ent["name"]) if folder_key == "person" else safe_name(ent["name"])
    if folder_key == "company" and name in SELF_COMPANIES:
        return f"{name}（自社）"
    if ent.get("certain", False) and name:
        return _link(folder_key, name) + honorific
    return f"{mask(ent['name'])}（聞き取り要確認）"


def business_hub_link(business: str) -> str:
    return f"[[{FOLDERS['business']}/{safe_name(business)}]]"


def month_hub_link(dt: datetime) -> str:
    return f"[[{FOLDERS['month']}/{dt:%Y-%m}]]"


def safe_note_path(path: str) -> str:
    """文字起こしノートのVault内パス（例 文字起こし/AINOTE/xxx.md）をリンク用に整える。"""
    parts = [safe_name(p, 120) for p in (path or "").replace("\\", "/").split("/") if p.strip()]
    parts = [p for p in parts if p and p not in (".", "..")]
    if not parts:
        return ""
    if parts[-1].endswith(".md"):
        parts[-1] = parts[-1][:-3]
    return "/".join(parts)


def render_business_hub(business: str, generated_at: str) -> str:
    name = safe_name(business)
    return "\n".join([
        "---", "type: 事業", f"name: {yaml_str(name)}", f"generated_at: {generated_at}",
        "tags: [AI秘書, 事業ハブ]", "---", "",
        f"# 事業：{name}", "",
        "この事業に分類された議事録・文字起こしの集まる場所です（分類はAI推定。違っていたら議事録の business を直してください）。",
        "グラフビューでは、このノートを中心に関連ノートが束ねられます。", "",
        "## 議事録（自動一覧）",
        "```base",
        "filters:",
        "  and:",
        '    - file.inFolder("AI秘書/議事録")',
        f'    - business == "{name}"',
        "views:",
        "  - type: table",
        "    name: 新しい順",
        "    order: [file.name, summary, basis]",
        "    sort:",
        "      - property: file.name",
        "        direction: DESC",
        "```", "",
        "## メモ（手入力欄）", "- ", "",
    ]) + "\n"


def render_month_hub(month: str, generated_at: str) -> str:
    return "\n".join([
        "---", "type: 月別", f"month: {yaml_str(month)}", f"generated_at: {generated_at}",
        "tags: [AI秘書, 月別ハブ]", "---", "",
        f"# {month} の記録", "",
        "この月の議事録・文字起こしの集まる場所です。", "",
        "## この月のノート（自動一覧）",
        "```base",
        "filters:",
        "  and:",
        '    - file.inFolder("AI秘書")',
        f'    - file.name.startsWith("{month}")',
        f'    - \'!file.inFolder("AI秘書/月別")\'',
        "views:",
        "  - type: table",
        "    name: 日付順",
        "    order: [file.name, file.folder, business]",
        "    sort:",
        "      - property: file.name",
        "        direction: ASC",
        "```", "",
    ]) + "\n"


def minutes_filename(d: dict) -> str:
    src = d["source"]
    dt = parse_recorded_at(src["recorded_at"])
    title = safe_name(src["title"], 30) or "無題"
    return f"{dt:%Y-%m-%d_%H%M}_{title}__{short_id(src['type'], src['id'])}.md"


def render_minutes(d: dict, generated_at: str) -> str:
    src = d["source"]
    dt = parse_recorded_at(src["recorded_at"])
    people = [p for p in d.get("people", [])]
    companies = d.get("companies", [])
    cases = d.get("cases", [])
    certain_people = [person_key(p["name"]) for p in people
                      if p.get("certain") and person_key(p["name"]) not in SELF_NAMES]
    certain_companies = [safe_name(c["name"]) for c in companies
                         if c.get("certain") and safe_name(c["name"]) not in SELF_COMPANIES]
    certain_cases = [safe_name(c["name"]) for c in cases if c.get("certain", True)]

    fm = [
        "---",
        "type: 議事録",
        f"title: {yaml_str(mask(src['title']))}",
        f"recorded_at: {dt.isoformat(timespec='minutes')}",
        f"source: {SOURCE_LABEL[src['type']]}",
        f"source_id: {yaml_str(src['type'] + ':' + src['id'])}",
        f"source_url: {yaml_str(src.get('url', ''))}",
        f"source_file: {yaml_str(src.get('file_name', ''))}",
        f"duration_min: {int(src.get('duration_min') or 0)}",
        f"basis: {src['basis']}",
        f"business: {d.get('business', '不明')}",
        f"business_hub: {yaml_str(business_hub_link(d.get('business', '不明')))}",
        f"month: {yaml_str(month_hub_link(dt))}",
        f"sensitivity: {d.get('sensitivity', '通常')}",
        "people: " + yaml_list("[[%s/%s]]" % (FOLDERS["person"], n) for n in certain_people),
        "companies: " + yaml_list("[[%s/%s]]" % (FOLDERS["company"], n) for n in certain_companies),
        "cases: " + yaml_list("[[%s/%s]]" % (FOLDERS["case"], n) for n in certain_cases),
        f"summary: {yaml_str(mask(d['summary'])[:120])}",
        f"generated_at: {generated_at}",
        f"generator: ai-secretary/render_notes.py v{VERSION}",
        "tags: [AI秘書, 議事録]",
        "---",
        "",
    ]
    b: list[str] = []
    b.append(f"# {mask(src['title'])}")
    b.append("")
    b.append(f"> 録音日時: {dt:%Y-%m-%d}({WEEKDAYS[dt.weekday()]}) {dt:%H:%M} JST ／ 情報源: {SOURCE_LABEL[src['type']]}"
             f" ／ 根拠: {src['basis']} ／ 事業: {_link('business', d.get('business', '不明'))}"
             f" ／ 月: {_link('month', f'{dt:%Y-%m}')}")
    if src["basis"] == "要約のみ":
        b.append("> ⚠ AI要約のみから作成。原文の文字起こしは未確認です。")
    if d.get("sensitivity") == "機密":
        b.append("> 🔒 機密扱い：外部共有・転記禁止")
    b.append("")
    b.append("## 要約")
    b.append(mask(d["summary"]))
    b.append("")

    def section(title: str, lines: list[str]):
        b.append(f"## {title}")
        b.extend(lines if lines else ["- なし"])
        b.append("")

    if d.get("key_points"):
        section("重要事項", [f"- {mask(x)}" for x in d["key_points"]])
    section("関係者", [f"- {_entity_ref('person', p, p.get('honorific', ''))}"
                     + (f"（{mask(p['company'])}）" if p.get("company") else "")
                     + (f" — {mask(p['role'])}" if p.get("role") else "") for p in people]
            + [f"- 会社: {_entity_ref('company', c)}" for c in companies])
    section("案件", [f"- {_entity_ref('case', {**c, 'certain': c.get('certain', True)})}"
                   + (f" — 状況: {mask(c['status'])}" if c.get("status") else "") for c in cases])
    section("決定事項", [f"- {mask(x)}" for x in d.get("decisions", [])])
    section("約束", [f"- {mask(p.get('who', '本人'))} → {mask(p.get('to', ''))}：{mask(p['what'])}"
                   + (f"（期限 {p['due']}）" if p.get("due") else "") for p in d.get("promises", [])])

    mine, others = [], []
    for t in d.get("todos", []):
        due = f" 📅 {t['due']}" if t.get("due") else ""
        pri = f" [{t.get('priority', 'P3')}]"
        owner = t.get("owner", "未確定")
        line = f"- [ ] {mask(t['task'])}{pri}{due}"
        if owner in SELF_NAMES:
            pass
        elif owner == "未確定":
            line += " （担当者要確認：本人か相手か録音から判別不可）"
        else:
            others.append(f"- [ ] {mask(owner)}：{mask(t['task'])}（相手の対応待ち）{due}")
            continue
        if t.get("due_basis"):
            line += f" （期限根拠: {mask(t['due_basis'])}）"
        mine.append(line)
    section("ToDo（本人・担当要確認を含む）", mine)
    section("確認待ち（相手の対応）", others)

    sch = []
    for s in d.get("schedules", []):
        when = " ".join(x for x in (s.get("date"), s.get("start")) if x) or "日時未定"
        mark = "✅確定" if s["status"] == "確定" else "❓要確認"
        line = f"- {mark} {when} {mask(s.get('title', ''))}"
        if s.get("place"):
            line += f" ＠{mask(s['place'])}"
        if s.get("missing"):
            line += f" — 不足: {mask(s['missing'])}"
        if s.get("evidence"):
            line += f"\n  - 根拠:「{mask(s['evidence'])}」"
        sch.append(line)
    section("予定", sch)
    section("要確認（聞き取り・解釈が曖昧）", [f"- {mask(x)}" for x in d.get("uncertain", [])])

    rel = []
    for r in d.get("related_notes", [])[:5]:
        rp = safe_note_path(r["path"])
        if rp and rp != f"AI秘書/{FOLDERS['minutes']}/{minutes_filename(d)[:-3]}":
            rel.append(f"- [[{rp}|{rp.split('/')[-1].split('__')[0]}]] — 共通: {mask(r['reason'])}")
    if rel:
        b.append("## 関連する議事録（同じ会社・案件が登場）")
        b.extend(rel)
        b.append("")
    b.append("## 原本")
    b.append(f"- 出典ID: `{src['type']}:{src['id']}`")
    if src.get("transcript_note"):
        tn = safe_note_path(src["transcript_note"])
        if tn:
            b.append(f"- 文字起こし全文: [[{tn}|全文を開く]]")
    if src.get("url"):
        b.append(f"- リンク: {src['url']}")
    if src.get("file_name"):
        b.append(f"- 元ファイル名: {mask(src['file_name'])}")
    if src["type"] == "plaud":
        b.append("- Plaudアプリで同じタイトルの録音を開くと音声・全文を確認できます")
    b.append("")
    if d.get("related"):
        b.append("## 関連資料（Drive内で見つかったもの・関連は推定）")
        for r in d["related"]:
            b.append(f"- [{mask(r['title'])}]({r['url']})" + (f" — {mask(r['note'])}" if r.get("note") else ""))
        b.append("")
    b.append("---")
    b.append("*AI秘書が自動生成。内容の修正は自由ですが、ファイル名の出典IDは消さないでください（重複防止に使用）。*")
    return "\n".join(fm + b) + "\n"


def _stub(kind: str, name: str, d: dict, extra: dict, generated_at: str) -> str:
    src = d["source"]
    dt = parse_recorded_at(src["recorded_at"])
    label = {"person": "人物", "company": "会社", "case": "案件"}[kind]
    field = {"person": "people", "company": "companies", "case": "cases"}[kind]
    fm = ["---", f"type: {label}", f"name: {yaml_str(name)}"]
    for k, v in extra.items():
        if v:
            fm.append(f"{k}: {yaml_str(mask(v))}")
    fm += [
        f"business: {d.get('business', '不明')}",
        f"first_seen: {dt.date().isoformat()}",
        f"first_source: {yaml_str(src['type'] + ':' + src['id'])}",
        f"generated_at: {generated_at}",
        f"tags: [AI秘書, {label}]",
        "---",
        "",
    ]
    body = [
        f"# {name}",
        "",
        f"初出: {dt:%Y-%m-%d %H:%M} 「{mask(src['title'])}」（{SOURCE_LABEL[src['type']]}）",
        "",
        "## メモ（手入力欄）",
        "- ",
        "",
        "## 関連する議事録（自動集約）",
        "Dataviewプラグインがあれば下に一覧が出ます。無くても右側の「バックリンク」に全件表示されます。",
        "",
        "```dataview",
        "TABLE recorded_at AS 録音日時, business AS 事業, summary AS 要約, basis AS 根拠",
        f'FROM "AI秘書/{FOLDERS["minutes"]}"',
        f"WHERE contains({field}, this.file.link)",
        "SORT recorded_at DESC",
        "```",
        "",
        "## 未完了ToDo（自動集約）",
        "```dataview",
        "TASK",
        f'FROM "AI秘書/{FOLDERS["minutes"]}"',
        f"WHERE !completed AND contains({field}, this.file.link)",
        "```",
        "",
    ]
    return "\n".join(fm + body) + "\n"


def render_all(extractions: list[dict], generated_at: str | None = None) -> list[dict]:
    generated_at = generated_at or datetime.now(JST).isoformat(timespec="minutes")
    out: list[dict] = []
    seen: set[tuple[str, str]] = set()
    for d in extractions:
        errs = validate(d)
        if errs:
            raise ValidationError(f"{d.get('source', {}).get('id', '?')}: " + " / ".join(errs))
        fn = minutes_filename(d)
        out.append({"folder": FOLDERS["minutes"], "filename": fn, "mode": "create_if_absent",
                    "source_id": f"{d['source']['type']}:{d['source']['id']}",
                    "content": render_minutes(d, generated_at)})
        for kind, field in (("person", "people"), ("company", "companies"), ("case", "cases")):
            for ent in d.get(field, []):
                certain = ent.get("certain", kind == "case")
                name = person_key(ent["name"]) if kind == "person" else safe_name(ent["name"])
                if not certain or not name or (kind, name) in seen:
                    continue
                if kind == "person" and (name in SELF_NAMES or ent["name"] in SELF_NAMES):
                    continue  # 本人の人物ノートは作らない
                if kind == "company" and name in SELF_COMPANIES:
                    continue  # 自社の会社ノートは作らない
                seen.add((kind, name))
                extra = {"company": ent.get("company"), "role": ent.get("role")} if kind == "person" else \
                        {"status": ent.get("status")} if kind == "case" else {}
                out.append({"folder": FOLDERS[kind], "filename": f"{name}.md", "mode": "create_if_absent",
                            "source_id": f"{d['source']['type']}:{d['source']['id']}",
                            "content": _stub(kind, name, d, extra, generated_at)})
        bname = safe_name(d.get("business", "不明"))
        if ("business", bname) not in seen:
            seen.add(("business", bname))
            out.append({"folder": FOLDERS["business"], "filename": f"{bname}.md", "mode": "create_if_absent",
                        "source_id": "", "content": render_business_hub(bname, generated_at)})
        month = f"{parse_recorded_at(d['source']['recorded_at']):%Y-%m}"
        if ("month", month) not in seen:
            seen.add(("month", month))
            out.append({"folder": FOLDERS["month"], "filename": f"{month}.md", "mode": "create_if_absent",
                        "source_id": "", "content": render_month_hub(month, generated_at)})
    return out


def render_digest(extractions: list[dict], files: list[dict], generated_at: str) -> dict:
    lines = ["---", "type: 処理ログ", f"generated_at: {generated_at}", "tags: [AI秘書, 処理ログ]", "---", "",
             f"# 処理ログ {generated_at}", "", "## 今回取り込んだ録音"]
    for d in extractions:
        s = d["source"]
        dt = parse_recorded_at(s["recorded_at"])
        lines.append(f"- {dt:%m/%d %H:%M} [[{FOLDERS['minutes']}/{minutes_filename(d)[:-3]}|{mask(s['title'])}]]"
                     f"（{SOURCE_LABEL[s['type']]}・{s['basis']}）")
    lines += ["", "## 予定（要確認を含む）"]
    for d in extractions:
        for sc in d.get("schedules", []):
            lines.append(f"- {sc['status']} {sc.get('date') or '日付未定'} {sc.get('start') or ''} "
                         f"{mask(sc.get('title', ''))}（{mask(d['source']['title'])}）")
    lines += ["", "## 生成したファイル"] + [f"- {f['folder']}/{f['filename']}" for f in files] + [""]
    stamp = generated_at.replace(":", "").replace("-", "")[:13].replace("T", "-")
    return {"folder": FOLDERS["digest"], "filename": f"処理ログ_{stamp}.md", "mode": "create_if_absent",
            "source_id": "", "content": "\n".join(lines)}


# ---------------------------------------------------------------- CLI
def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("render")
    r.add_argument("inputs", nargs="+")
    r.add_argument("--out", required=True)
    r.add_argument("--generated-at")
    r.add_argument("--no-digest", action="store_true")
    v = sub.add_parser("validate")
    v.add_argument("inputs", nargs="+")
    j = sub.add_parser("jst")
    j.add_argument("start_at")
    a = ap.parse_args(argv)

    if a.cmd == "jst":
        print(plaud_start_to_jst(a.start_at))
        return 0

    exts = [json.loads(Path(p).read_text(encoding="utf-8")) for p in a.inputs]
    if a.cmd == "validate":
        bad = 0
        for p, d in zip(a.inputs, exts):
            errs = validate(d)
            print(json.dumps({"file": p, "ok": not errs, "errors": errs}, ensure_ascii=False))
            bad += bool(errs)
        return 1 if bad else 0

    gen = a.generated_at or datetime.now(JST).isoformat(timespec="minutes")
    try:
        files = render_all(exts, gen)
    except ValidationError as e:
        print(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False))
        return 1
    if not a.no_digest:
        files.append(render_digest(exts, files, gen))
    out = Path(a.out)
    manifest = []
    for f in files:
        p = out / f["folder"] / f["filename"]
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(f["content"], encoding="utf-8")
        manifest.append({k: f[k] for k in ("folder", "filename", "mode", "source_id")} | {"local_path": str(p)})
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "files": len(manifest), "manifest": str(out / "manifest.json")}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
