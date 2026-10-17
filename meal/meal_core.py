#!/usr/bin/env python3
"""AI秘書: 食事・栄養管理の中核ロジック（標準ライブラリのみ）

Make の「Run code（Python）」モジュールにそのまま載せて使う（build_make_code.py が
このファイルの「CORE」部分を切り出して `return handle(input)` を付ける）。
Claude の定期ジョブからも同じ関数を CLI で使う。

処理の流れ（Make シナリオ「食事管理：記録・質問」）:
  1. Webhook（iPhoneショートカット）  {"token","text","image"(base64, 任意),"sent_at"(任意)}
  2. Sheets batchGetByDataFilter で食事DBを全部読む
  3. handle(mode="prepare")  → OpenAI へのリクエスト本文を作る
  4. OpenAI /v1/chat/completions（JSON Schema で構造化出力）
  5. handle(mode="apply")    → Sheets batchUpdate の本文と返答文を作る
  6. Sheets values:batchUpdate（書き込みがある時だけ）
  7. Webhook response（返答文）

設計方針:
  - 1日の区切りは日本時間 午前6時（6:00〜翌5:59:59 を同じ「集計日」とする）。
  - 品目は1行。元の推定値（1人前全部）は書き換えず、「食べた割合」「状態」「日時」だけを補正する。
    補正は必ず「補正履歴」シートに旧値・新値・元の発言を残す（削除はしない）。
  - 栄養は推定値。単一の数字だけでなく下限〜上限を持つ。
  - 同じ写真（内容ハッシュ一致）は二重登録しない。似た記録は自動統合せず確認を促す。
  - PLAUD・ボイスメモ由来は「確認候補」に入れ、本人が「候補1は記録」と言った時だけ確定にする。
"""
import hashlib
import json
import re
import unicodedata
from datetime import datetime, timedelta, timezone

VERSION = "1.0.0"
JST = timezone(timedelta(hours=9))
DAY_START_HOUR = 6
WEEKDAYS = "月火水木金土日"
DEFAULT_MODEL = "gpt-4.1-mini"
DISCLAIMER = "※栄養は推定値です（医療上の判断には使えません）"

# ---------------------------------------------------------------- シート定義（列の順番＝スプレッドシートの列）
MEAL_COLS = ["品目ID", "食事ID", "状態", "食べた日時", "集計日", "区分", "品名", "数量", "食べた割合",
             "kcal目安", "kcal下限", "kcal上限", "食塩g", "たんぱく質g", "脂質g", "炭水化物g", "アルコールg",
             "根拠", "確信度", "情報源", "元の発言", "登録日時", "更新日時", "写真ハッシュ", "補足"]
HISTORY_COLS = ["日時", "品目ID", "項目", "旧値", "新値", "元の発言", "情報源"]
CAND_COLS = ["候補ID", "状態", "情報源", "出典ID", "録音日時", "推定食事日時", "集計日", "品名", "数量",
             "kcal目安", "kcal下限", "kcal上限", "食塩g", "たんぱく質g", "脂質g", "炭水化物g", "アルコールg",
             "原文", "話者", "重複候補", "判定理由", "登録日時", "判断日時"]
WEIGHT_COLS = ["測定日", "体重kg", "情報源", "登録日時", "メモ"]
SETTING_COLS = ["項目", "値", "説明"]
SHEETS = {"食事": MEAL_COLS, "補正履歴": HISTORY_COLS, "確認候補": CAND_COLS, "体重": WEIGHT_COLS, "設定": SETTING_COLS}
DEFAULT_SETTINGS = [  # 個人の値は食事DBの「設定」シートに入れる（公開リポジトリに健康情報を置かない）
    ["目標kcal", 1800, "1日の摂取カロリー目標（暫定値。ここを書き換えると以後の集計に反映）"],
    ["食塩目標g", 7.5, "食塩相当量の1日の目安（未満）"],
    ["たんぱく質目安g", "", "参考値（体重1kgあたり約1g）。医療上の指示ではない"],
    ["1日の開始時刻", 6, "この時刻から翌日のこの時刻の直前までを1日として集計（変更不可・参考表示）"],
    ["身長cm", "", ""],
    ["目標体重kg", "", ""],
    ["年齢", "", ""],
    ["性別", "", ""],
    ["活動量", "", "例: 少なめ（経営・打ち合わせ中心、車移動が多い）"],
]
NUTRIENTS = [("kcal", "kcal目安"), ("kcal_low", "kcal下限"), ("kcal_high", "kcal上限"), ("salt", "食塩g"),
             ("protein", "たんぱく質g"), ("fat", "脂質g"), ("carb", "炭水化物g"), ("alcohol", "アルコールg")]
ACTIVE_STATUS = "確定"
STATUSES = {"確定", "取消", "本人以外"}
SLOTS = ["朝", "昼", "夜", "間食", "飲み", "不明"]
BASES = ["栄養表示", "成分表", "標準料理", "写真推定", "AI推定"]
CONFIDENCES = ["A", "B", "C", "D"]
SALT_ITEM_WARN = 3.0       # 1品でこれ以上なら塩分注意
PHOTO_DUP_DAYS = 3         # 同じ写真をこの日数以内に送ったら二重登録しない
NEAR_DUP_MINUTES = 90      # 同じ料理名がこの時間内にあれば「重複かも」と伝える

# 外食・定番料理の標準値（1人前の目安。kcal 目安・下限・上限、食塩g、P、F、C、アルコールg）
# 出典: 日本食品標準成分表（八訂）の食品値と一般的な外食の分量からの概算。あくまで推定の目安。
DISH_GUIDE = [
    ("ご飯 普通盛150g", 234, 200, 280, 0.0, 4, 0, 56, 0),
    ("ご飯 大盛250g", 390, 340, 450, 0.0, 6, 1, 93, 0),
    ("卵かけご飯", 330, 290, 380, 0.9, 10, 6, 57, 0),
    ("おにぎり1個（コンビニ）", 180, 160, 230, 1.0, 4, 1, 40, 0),
    ("カツ丼", 950, 850, 1100, 4.0, 35, 35, 120, 0),
    ("親子丼", 700, 620, 800, 3.5, 30, 18, 100, 0),
    ("牛丼 並", 650, 600, 750, 2.8, 22, 22, 90, 0),
    ("カレーライス", 800, 700, 950, 3.5, 20, 28, 110, 0),
    ("チャーハン", 700, 600, 850, 3.5, 18, 25, 95, 0),
    ("醤油ラーメン（スープ全部）", 500, 450, 600, 6.0, 20, 10, 70, 0),
    ("とんこつラーメン（スープ全部）", 800, 700, 950, 6.5, 28, 35, 80, 0),
    ("味噌ラーメン（スープ全部）", 700, 600, 850, 7.0, 25, 25, 85, 0),
    ("つけ麺", 750, 650, 900, 6.0, 30, 15, 115, 0),
    ("かけうどん（汁全部）", 350, 300, 400, 5.5, 10, 2, 70, 0),
    ("ざるそば", 350, 300, 420, 2.5, 13, 3, 65, 0),
    ("パスタ ミートソース 1人前", 650, 580, 780, 3.0, 22, 20, 90, 0),
    ("パスタ クリーム系 1人前", 800, 700, 950, 3.0, 25, 35, 90, 0),
    ("パスタ 大盛（1.5人前）", 1000, 870, 1300, 4.5, 33, 30, 135, 0),
    ("ハンバーグ定食", 900, 800, 1050, 4.0, 35, 35, 100, 0),
    ("生姜焼き定食", 850, 750, 950, 4.5, 30, 30, 100, 0),
    ("焼き魚定食", 650, 550, 750, 4.5, 30, 15, 90, 0),
    ("幕の内弁当", 750, 650, 900, 4.0, 25, 22, 105, 0),
    ("駅弁", 800, 650, 1000, 3.5, 25, 25, 110, 0),
    ("焼き鳥 もも 1本", 100, 80, 120, 0.5, 8, 7, 2, 0),
    ("焼き鳥 むね 1本", 70, 60, 90, 0.4, 10, 2, 1, 0),
    ("焼き鳥 皮 1本", 110, 90, 130, 0.5, 4, 10, 1, 0),
    ("焼き鳥 つくね 1本", 100, 80, 120, 0.7, 7, 6, 4, 0),
    ("焼き鳥 ねぎま 1本", 80, 65, 100, 0.5, 7, 5, 2, 0),
    ("唐揚げ 1個30g", 80, 65, 100, 0.4, 5, 5, 3, 0),
    ("餃子 1個", 45, 35, 55, 0.3, 2, 2, 4, 0),
    ("刺身盛り合わせ 1人前（醤油込み）", 200, 150, 280, 1.2, 25, 8, 2, 0),
    ("天ぷら盛り合わせ 1人前（つゆ込み）", 500, 400, 650, 1.5, 12, 30, 40, 0),
    ("寿司 1貫（醤油込み）", 55, 40, 70, 0.4, 2, 1, 9, 0),
    ("焼肉 1人前150g（カルビ等）", 600, 450, 800, 1.5, 25, 50, 5, 0),
    ("ステーキ 200g", 600, 450, 750, 1.5, 40, 45, 2, 0),
    ("枝豆 1皿", 120, 90, 150, 0.8, 11, 6, 8, 0),
    ("冷奴（醤油込み）", 100, 80, 130, 0.8, 9, 5, 3, 0),
    ("味噌汁 1杯", 40, 30, 60, 1.5, 3, 1, 4, 0),
    ("サラダ（ドレッシング込み）", 100, 50, 200, 0.8, 2, 7, 6, 0),
    ("サンドイッチ 1パック", 300, 250, 400, 1.5, 12, 15, 30, 0),
    ("菓子パン 1個", 350, 250, 500, 0.8, 7, 15, 50, 0),
    ("ポテトチップス 1袋60g", 330, 300, 360, 0.6, 3, 21, 32, 0),
    ("ビール 中ジョッキ約400ml", 170, 140, 200, 0.0, 1, 0, 12, 16),
    ("ビール 350ml缶", 140, 130, 150, 0.0, 1, 0, 11, 14),
    ("ハイボール 1杯（ウイスキー30ml）", 75, 65, 110, 0.0, 0, 0, 0, 10),
    ("レモンサワー 1杯", 130, 80, 200, 0.0, 0, 0, 10, 14),
    ("日本酒 1合", 190, 180, 200, 0.0, 1, 0, 8, 22),
    ("焼酎 水割り・ロック 1杯（25度60ml）", 85, 80, 95, 0.0, 0, 0, 0, 12),
    ("ワイン グラス120ml", 90, 80, 100, 0.0, 0, 0, 2, 11),
    ("コーヒー ブラック", 5, 0, 10, 0.0, 0, 0, 1, 0),
]


# ---------------------------------------------------------------- 日付・数値の基本関数
def now_jst():
    return datetime.now(JST)


def parse_dt(value):
    """'YYYY-MM-DD HH:MM' / ISO8601 → JST の datetime。解釈できなければ None。"""
    if not value:
        return None
    s = unicodedata.normalize("NFKC", str(value)).strip().replace("/", "-")
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?\s*([+-]\d{2}:?\d{2})?$", s)
    if not m:
        return None
    y, mo, d, hh, mi, ss, off = m.groups()
    tz = JST
    if off:
        sign = 1 if off[0] == "+" else -1
        tz = timezone(sign * timedelta(hours=int(off[1:3]), minutes=int(off[-2:])))
    try:
        dt = datetime(int(y), int(mo), int(d), int(hh or 0), int(mi or 0), int(ss or 0), tzinfo=tz)
    except ValueError:
        return None
    return dt.astimezone(JST)


def fmt_dt(dt):
    return dt.astimezone(JST).strftime("%Y-%m-%d %H:%M")


def day_key(dt):
    """集計日（午前6時区切り）。10/18 05:59 → 2026-10-17、10/18 06:00 → 2026-10-18。"""
    return (dt.astimezone(JST) - timedelta(hours=DAY_START_HOUR)).strftime("%Y-%m-%d")


def day_label(key):
    d = datetime.strptime(key, "%Y-%m-%d")
    return f"{d.month}/{d.day}({WEEKDAYS[d.weekday()]})"


def shift_day(key, days):
    return (datetime.strptime(key, "%Y-%m-%d") + timedelta(days=days)).strftime("%Y-%m-%d")


def num(value, default=0.0):
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    s = unicodedata.normalize("NFKC", str(value)).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return default


def r1(x):
    return round(float(x) + 0.0, 1)


def norm_name(name):
    s = unicodedata.normalize("NFKC", name or "").lower()
    return re.sub(r"[\s・（）()\[\]【】、。,.\-ー~〜0-9０-９本杯個皿人前gml大中小盛]+", "", s)


def comma(x):
    return f"{int(round(x)):,}"


def jload(value):
    if value is None or value == "":
        return {}
    if isinstance(value, (dict, list)):
        return value
    return json.loads(value)


# ---------------------------------------------------------------- DB の読み込み
class Table:
    """シート1枚分。rows は dict（列名→値）に __row（シートの行番号）を付けたもの。"""

    def __init__(self, name, cols, values):
        self.name = name
        self.cols = cols
        self.rows = []
        values = values or []
        start = 1 if values and values[0] and str(values[0][0]) == cols[0] else 0
        for i, raw in enumerate(values[start:], start=start + 1):
            if not raw or all(str(v).strip() == "" for v in raw):
                continue
            row = {c: (raw[j] if j < len(raw) else "") for j, c in enumerate(cols)}
            row["__row"] = i
            self.rows.append(row)

    def find(self, col, value):
        return [r for r in self.rows if str(r.get(col, "")) == str(value)]


def load_db(batch_body):
    """Sheets の batchGet / batchGetByDataFilter の応答（dict か JSON 文字列）→ {シート名: Table}。"""
    body = jload(batch_body)
    ranges = body.get("valueRanges", []) if isinstance(body, dict) else []
    by_sheet = {}
    for vr in ranges:
        vr = vr.get("valueRange", vr)
        name = str(vr.get("range", "")).split("!")[0].strip("'")
        by_sheet[name] = vr.get("values", [])
    return {name: Table(name, cols, by_sheet.get(name, [])) for name, cols in SHEETS.items()}


def settings(db):
    s = {k: v for k, v, _ in DEFAULT_SETTINGS}
    for r in db["設定"].rows:
        if r["項目"] in s and str(r["値"]).strip() != "":
            s[r["項目"]] = r["値"]
    return {"kcal": num(s["目標kcal"], 1800), "salt": num(s["食塩目標g"], 7.5),
            "protein": num(s["たんぱく質目安g"], 0), "target_weight": num(s["目標体重kg"], 0),
            "profile": {k: s[k] for k in ("年齢", "性別", "身長cm", "目標体重kg", "活動量")}}


def profile_text(db):
    """プロンプト用の本人情報（「設定」シートと最新の体重記録から作る）。"""
    p = settings(db)["profile"]
    parts = []
    if str(p["年齢"]).strip():
        parts.append(f"{p['年齢']}歳")
    if str(p["性別"]).strip():
        parts.append(str(p["性別"]))
    if str(p["身長cm"]).strip():
        parts.append(f"身長{p['身長cm']}cm")
    w = latest_weights(db, 1)
    if w:
        parts.append(f"直近の体重{w[-1]['体重kg']}kg（{w[-1]['測定日']}）")
    if str(p["目標体重kg"]).strip():
        parts.append(f"目標体重{p['目標体重kg']}kg")
    if str(p["活動量"]).strip():
        parts.append(f"活動量{p['活動量']}")
    return "・".join(parts) or "（未設定）"


# ---------------------------------------------------------------- 集計
def effective(row):
    """品目1行の実摂取量（推定値 × 食べた割合）。状態が確定以外は0。"""
    if str(row.get("状態", "")) != ACTIVE_STATUS:
        return {k: 0.0 for k, _ in NUTRIENTS}
    ratio = min(max(num(row.get("食べた割合"), 1.0), 0.0), 1.0)
    return {k: num(row.get(col)) * ratio for k, col in NUTRIENTS}


def day_totals(db, key):
    t = {k: 0.0 for k, _ in NUTRIENTS}
    items = [r for r in db["食事"].rows if r.get("集計日") == key and r.get("状態") == ACTIVE_STATUS]
    for r in items:
        for k, v in effective(r).items():
            t[k] += v
    t["count"] = len(items)
    return t


def week_stats(db, end_key, days=7):
    """end_key を含む直近 days 日。記録が無い日は平均から除く（0kcalとして扱わない）。"""
    per_day = []
    for i in range(days - 1, -1, -1):
        k = shift_day(end_key, -i)
        per_day.append((k, day_totals(db, k)))
    recorded = [t for _, t in per_day if t["count"] > 0]
    avg = {k: (sum(t[k] for t in recorded) / len(recorded) if recorded else 0.0) for k, _ in NUTRIENTS}
    return {"per_day": per_day, "recorded_days": len(recorded), "avg": avg}


def open_candidates(db):
    rows = [r for r in db["確認候補"].rows if r.get("状態") == "保留"]
    rows.sort(key=lambda r: (str(r.get("登録日時", "")), str(r.get("候補ID", ""))))
    return rows


def latest_weights(db, n=4):
    rows = [r for r in db["体重"].rows if num(r.get("体重kg"), 0) > 0]
    rows.sort(key=lambda r: str(r.get("測定日", "")))
    return rows[-n:]


# ---------------------------------------------------------------- OpenAI リクエスト
def _dish_guide_text():
    lines = ["料理 | kcal目安(下限〜上限) | 食塩g | P | F | C | アルコールg"]
    for name, k, lo, hi, salt, p, f, c, alc in DISH_GUIDE:
        lines.append(f"{name} | {k}({lo}〜{hi}) | {salt} | {p} | {f} | {c} | {alc}")
    return "\n".join(lines)


SYSTEM_PROMPT = """あなたは経営者JP（{profile}）の食事記録係です。出張・外食・会食・夜の飲酒が多い。
JPの短い報告（音声入力の文字起こし・写真）から、食べたもの・補正・質問を読み取り、指定のJSONだけを返します。

# 1日の区切り（最重要）
- 集計日は日本時間の午前6時〜翌朝5時59分。0:00〜5:59に食べたものは前の集計日に入る。
- 「今日」「昨日」は集計日で解釈する。例: 現在が10/18 10:00で「昨日の夜中2時にラーメン」→ eaten_at は 2026-10-18 02:00（集計日10/17）。
- 時刻の言及が無ければ eaten_at は空文字（送信時刻を使う）。「昼」「夜」だけなら常識的な時刻（昼12:30、夜19:30、朝7:30）を入れ time_basis を「推定」にする。
- 現在が0:00〜5:59で「昨日」「今日」が2通りに解釈できる時は、より自然な方を選び needs_confirmation に確認文を入れる。

# 新しい食事（new_items）
- 1品＝1要素。「焼き鳥もも3本、胸肉2本」→ 2要素（もも×3、むね×2）。数量込みの合計値を入れる。
- 栄養は「1人前（出された量）全部」の値を kcal/kcal_low/kcal_high/salt/protein/fat/carb/alcohol に入れ、実際に食べた割合は ratio（0〜1）で表す。
  「半分」→0.5、「3分の1」→0.33、「全部・言及なし」→1、「ご飯は残した（量の明言なし）」→0.5、「スープは飲んでない」→ラーメン等は ratio 1 のまま、スープ分を引いた値（食塩は約4割、kcalは約1〜2割減）を入れ note に書く。
- 大盛・小盛は量を調整（大盛≒1.4〜1.5倍、小盛≒0.7倍）。
- 単一の正確そうな数字を作らない。kcal_low〜kcal_high は情報量に応じて幅を取る（写真・外食は±15〜30%）。
- basis: 栄養表示（商品の表示値が分かる）/成分表（食材・量が明確）/標準料理（下の目安表や一般的な1人前）/写真推定/AI推定。
- confidence: A=表示値, B=量が明確, C=外食・写真の一般量から推定, D=情報不足。
- 他人の料理・食べていないもの・予定（「これから食べる」）は new_items に入れない。
- 写真だけで説明が無ければ、写っている料理を本人が全部食べたものとして入れる（basis=写真推定）。
- slot: 朝/昼/夜/間食/飲み/不明。

# 既存記録の補正（amendments）
- 「これ半分食べた」「さっきのカツ丼は取消」「刺身は半分」「これは友達の料理」「これは昨日食べたもの」等は、直近の記録一覧から対象を選んで補正する。新しい食事として二重に入れない。
- target_id は記録一覧の品目ID（例 M20261017-1230-01-02）か食事ID（例 M20261017-1230-01、その食事の全品目）。
- field: 割合（value は 0〜1 の数字の文字列）/ 日時（value は 'YYYY-MM-DD HH:MM'）/ 状態（value は 取消・本人以外・確定）。
- 対象が特定できない時は補正せず needs_confirmation に質問を入れる。
- 補正は「今回のJPの報告」で明示された時だけ出す。直近の記録一覧に既に反映済みの補正（状態が取消の品目など）を繰り返さない。

# 確認候補（candidate_decisions）
- 「候補1は記録」「候補2は不要」等。候補番号は一覧の番号。decision は 記録 か 不要。

# 体重（weight）
- 「体重77.6」等。date は集計日（言及が無ければ空文字）。無ければ kg は null。

# 質問への回答（answer）
- 今日の合計・残り・塩分・栄養素・栄養状態の質問（「今日何カロリー？」「あと何カロリー？」「今日の塩分は？」等）は answer を空文字にする（システムが最新の集計を自動で付ける）。評価が欲しそうな時は advice に1文だけ書く。
- 「昨日は食べすぎた？」「今週の平均は？」等は、渡された日別集計だけを使って2〜4行で答える。単日でなく週平均で判断する。記録の無い日は推測しない。体重記録が無ければ体重変化を断定しない。
- 食事の判断は参考情報であり医療上の診断ではない。

# 飲酒の目安
- 節度ある適度な飲酒は純アルコールで1日平均約20g程度（厚生労働省の目安。ビール中瓶1本≒20g）。超えている日に「適度」と言わない。説教はせず事実と次の工夫だけ伝える。

# 一言アドバイス（advice）
- 今日の記録・残りカロリー・塩分を踏まえ、次の食事で実行できることを1〜2文。不要なら空文字。説教調にしない。

# 外食・定番料理の目安（1人前）
{dish_guide}
"""


def response_schema():
    def obj(props, required=None):
        return {"type": "object", "properties": props, "required": required or list(props), "additionalProperties": False}

    s = {"type": "string"}
    n = {"type": "number"}
    item = obj({
        "meal_ref": {"type": "string", "description": "同じ食事（同じ時刻・同じ写真）の品目は同じ値（例 A, B）"},
        "eaten_at": {"type": "string", "description": "'YYYY-MM-DD HH:MM' または空文字（送信時刻を使う）"},
        "time_basis": {"type": "string", "enum": ["発言", "推定", "送信時刻", "写真"]},
        "slot": {"type": "string", "enum": SLOTS},
        "name": s, "quantity": s,
        "ratio": n, "kcal": n, "kcal_low": n, "kcal_high": n, "salt": n,
        "protein": n, "fat": n, "carb": n, "alcohol": n,
        "basis": {"type": "string", "enum": BASES},
        "confidence": {"type": "string", "enum": CONFIDENCES},
        "note": s,
    })
    amend = obj({"target_id": s, "field": {"type": "string", "enum": ["割合", "日時", "状態"]}, "value": s, "reason": s})
    dec = obj({"candidate_no": {"type": "integer"}, "decision": {"type": "string", "enum": ["記録", "不要"]}})
    schema = obj({
        "intent": {"type": "string", "enum": ["記録", "補正", "質問", "候補判断", "体重", "複合", "不明"]},
        "new_items": {"type": "array", "items": item},
        "amendments": {"type": "array", "items": amend},
        "candidate_decisions": {"type": "array", "items": dec},
        "weight": obj({"kg": {"type": ["number", "null"]}, "date": s}),
        "needs_confirmation": {"type": "array", "items": s},
        "answer": s,
        "advice": s,
    })
    return {"name": "meal_log", "strict": True, "schema": schema}


def recent_items_text(db, now, hours=36, limit=40):
    lines = []
    cutoff = now - timedelta(hours=hours)
    rows = [r for r in db["食事"].rows if (parse_dt(r.get("食べた日時")) or now) >= cutoff
            or (parse_dt(r.get("登録日時")) or now) >= cutoff]
    rows.sort(key=lambda r: (str(r.get("食べた日時")), str(r.get("品目ID"))))
    for r in rows[-limit:]:
        lines.append(f"{r['品目ID']} | 食事ID {r['食事ID']} | {r['状態']} | {r['食べた日時']}（集計日{r['集計日']}） | "
                     f"{r['品名']} {r['数量']} | 割合{r['食べた割合']} | {r['kcal目安']}kcal | 情報源 {r['情報源']}"
                     + (" | 写真あり" if r.get("写真ハッシュ") else ""))
    return "\n".join(lines) or "（直近36時間の記録なし）"


def candidates_text(db):
    lines = []
    for i, r in enumerate(open_candidates(db), start=1):
        lines.append(f"候補{i} | {r['候補ID']} | 推定 {r['推定食事日時'] or '日時不明'}（集計日{r['集計日'] or '不明'}） | "
                     f"{r['品名']} {r['数量']} | 情報源 {r['情報源']} | 原文「{r['原文']}」")
    return "\n".join(lines) or "（保留中の確認候補なし）"


def daily_text(db, today, days=8):
    lines = []
    st = settings(db)
    for i in range(days - 1, -1, -1):
        k = shift_day(today, -i)
        t = day_totals(db, k)
        if t["count"] == 0:
            lines.append(f"{k} {day_label(k)}: 記録なし")
        else:
            lines.append(f"{k} {day_label(k)}: {comma(t['kcal'])}kcal（{comma(t['kcal_low'])}〜{comma(t['kcal_high'])}）"
                         f" 食塩{r1(t['salt'])}g P{round(t['protein'])}g アルコール{round(t['alcohol'])}g 品数{t['count']}")
    ws = week_stats(db, today)
    lines.append(f"直近7日（今日を含む）の平均: {comma(ws['avg']['kcal'])}kcal、食塩{r1(ws['avg']['salt'])}g"
                 f"（記録のある{ws['recorded_days']}日の平均）／目標 {comma(st['kcal'])}kcal・食塩{st['salt']}g未満")
    w = latest_weights(db)
    lines.append("体重記録: " + ("、".join(f"{r['測定日']} {r['体重kg']}kg" for r in w) if w else "なし"))
    return "\n".join(lines)


def build_openai_request(db, text, now, image_b64="", model=DEFAULT_MODEL):
    today = day_key(now)
    user = (f"現在: {fmt_dt(now)}（{WEEKDAYS[now.weekday()]}曜）／今日の集計日: {today}（{day_label(today)} 6:00〜翌5:59）\n\n"
            f"## 直近の記録（品目ID付き）\n{recent_items_text(db, now)}\n\n"
            f"## 確認候補（保留中）\n{candidates_text(db)}\n\n"
            f"## 日別集計\n{daily_text(db, today)}\n\n"
            f"## JPの報告\n{text.strip() if text else '（文章なし）'}"
            + ("\n（写真あり）" if image_b64 else ""))
    content = [{"type": "text", "text": user}]
    if image_b64:
        content.append({"type": "image_url", "image_url": {"url": "data:image/jpeg;base64," + image_b64, "detail": "high"}})
    return {
        "model": model,
        "temperature": 0.2,
        "messages": [{"role": "system", "content": SYSTEM_PROMPT.replace("{dish_guide}", _dish_guide_text())
                                                    .replace("{profile}", profile_text(db))},
                     {"role": "user", "content": content}],
        "response_format": {"type": "json_schema", "json_schema": response_schema()},
    }


def parse_openai_response(body):
    body = jload(body)
    if isinstance(body, dict) and "choices" in body:
        content = body["choices"][0]["message"].get("content") or "{}"
        return json.loads(content)
    return body if isinstance(body, dict) else {}


# ---------------------------------------------------------------- 適用（書き込み内容と返答文）
class Writer:
    """書き込み内容を Sheets の spreadsheets:batchUpdate 本文にまとめる。
    追加は appendCells（シート末尾への追加。ほぼ同時に2件届いても行が上書きされない）、
    既存行の補正は updateCells（行は削除・移動しないので行番号は変わらない）。"""

    def __init__(self, db, sheet_ids):
        self.db = db
        self.sheet_ids = sheet_ids
        self.requests = []

    @staticmethod
    def _cell(value):
        if isinstance(value, bool):
            return {"userEnteredValue": {"stringValue": str(value)}}
        if isinstance(value, (int, float)):
            return {"userEnteredValue": {"numberValue": value}}
        return {"userEnteredValue": {"stringValue": "" if value is None else str(value)}}

    def append(self, sheet, row_dict):
        values = [self._cell(row_dict.get(c, "")) for c in SHEETS[sheet]]
        self.requests.append({"appendCells": {"sheetId": self.sheet_ids[sheet], "rows": [{"values": values}],
                                              "fields": "userEnteredValue"}})
        row = dict(row_dict)
        row["__row"] = None  # 末尾に追加（行番号は書き込み時に決まる）
        self.db[sheet].rows.append(row)  # 返答の集計に反映させる
        return row

    def update(self, sheet, row, changes):
        cols = SHEETS[sheet]
        for col, value in changes.items():
            if row.get("__row"):
                self.requests.append({"updateCells": {
                    "start": {"sheetId": self.sheet_ids[sheet], "rowIndex": row["__row"] - 1,
                              "columnIndex": cols.index(col)},
                    "rows": [{"values": [self._cell(value)]}], "fields": "userEnteredValue"}})
            row[col] = value

    def body(self):
        return {"requests": self.requests} if self.requests else None


def load_sheet_ids(meta):
    """GET spreadsheets/{id}?fields=sheets.properties の応答 → {シート名: sheetId}"""
    meta = jload(meta)
    return {sh["properties"]["title"]: sh["properties"]["sheetId"] for sh in meta.get("sheets", [])}


def _clean_item(it):
    vals = {k: max(num(it.get(k)), 0.0) for k, _ in NUTRIENTS}
    lo, mid, hi = sorted([vals["kcal_low"] or vals["kcal"], vals["kcal"], vals["kcal_high"] or vals["kcal"]])
    vals["kcal_low"], vals["kcal"], vals["kcal_high"] = lo, mid, hi
    ratio = num(it.get("ratio"), 1.0)
    ratio = 1.0 if ratio <= 0 and not it.get("note") else min(max(ratio, 0.0), 1.0)
    return vals, ratio


def image_hash(image_b64):
    if not image_b64:
        return ""
    try:
        import base64
        raw = base64.b64decode(image_b64, validate=False)
    except Exception:
        raw = image_b64.encode()
    return hashlib.sha256(raw).hexdigest()[:16]


def apply(db, result, now, text="", source="ショートカット", photo_hash="", sheet_ids=None):
    """LLM の構造化結果を検証して DB に反映する書き込みを作る。戻り値: (Writer, 報告行のリスト)"""
    w = Writer(db, sheet_ids or {name: i for i, name in enumerate(SHEETS)})
    notes, recorded, warns = [], [], []
    stamp = fmt_dt(now)
    said = (text or "").strip()[:300]
    meals = db["食事"]

    # --- 同じ写真の二重登録防止
    new_items = result.get("new_items") or []
    if photo_hash and new_items:
        dup = [r for r in meals.rows if r.get("写真ハッシュ") == photo_hash
               and (parse_dt(r.get("登録日時")) or now) >= now - timedelta(days=PHOTO_DUP_DAYS)]
        if dup:
            new_items = []
            notes.append(f"この写真は登録済みです（{dup[0]['食べた日時']} {dup[0]['品名']} ほか）。二重には記録していません。")

    # --- 新しい食事
    groups = {}
    for it in new_items:
        name = str(it.get("name", "")).strip()
        if not name:
            continue
        eaten = parse_dt(it.get("eaten_at")) or now
        if eaten > now + timedelta(minutes=10):
            warns.append(f"「{name}」の日時が未来（{fmt_dt(eaten)}）だったため送信時刻で記録しました。")
            eaten = now
        ref = str(it.get("meal_ref") or "A")
        if ref not in groups:
            groups[ref] = f"M{now.strftime('%Y%m%d-%H%M%S')}-{len(groups) + 1:02d}"
        meal_id = groups[ref]
        seq = sum(1 for r in meals.rows if r.get("食事ID") == meal_id) + 1
        vals, ratio = _clean_item(it)
        row = {
            "品目ID": f"{meal_id}-{seq:02d}", "食事ID": meal_id, "状態": ACTIVE_STATUS,
            "食べた日時": fmt_dt(eaten), "集計日": day_key(eaten), "区分": it.get("slot") or "不明",
            "品名": name, "数量": str(it.get("quantity", "")), "食べた割合": round(ratio, 2),
            "kcal目安": round(vals["kcal"]), "kcal下限": round(vals["kcal_low"]), "kcal上限": round(vals["kcal_high"]),
            "食塩g": r1(vals["salt"]), "たんぱく質g": r1(vals["protein"]), "脂質g": r1(vals["fat"]),
            "炭水化物g": r1(vals["carb"]), "アルコールg": r1(vals["alcohol"]),
            "根拠": it.get("basis") or "AI推定", "確信度": it.get("confidence") or "D", "情報源": source,
            "元の発言": said, "登録日時": stamp, "更新日時": stamp, "写真ハッシュ": photo_hash,
            "補足": "時刻: " + str(it.get("time_basis") or "送信時刻") + ("／" + it["note"] if it.get("note") else ""),
        }
        near = [r for r in meals.rows
                if r.get("状態") == ACTIVE_STATUS and r.get("食事ID") != meal_id
                and norm_name(r.get("品名")) == norm_name(name)
                and abs(((parse_dt(r.get("食べた日時")) or now) - eaten).total_seconds()) <= NEAR_DUP_MINUTES * 60]
        w.append("食事", row)
        recorded.append(row)
        if near:
            warns.append(f"「{name}」は {near[0]['食べた日時']} にも記録があります。同じものなら「さっきの{name}は取消」と言ってください。")
        if vals["salt"] * ratio >= SALT_ITEM_WARN:
            warns.append(f"「{name}」は食塩が多め（約{r1(vals['salt'] * ratio)}g）。")

    # --- 補正
    for a in result.get("amendments") or []:
        tid, field, value = str(a.get("target_id", "")).strip(), a.get("field"), str(a.get("value", "")).strip()
        targets = meals.find("品目ID", tid) or meals.find("食事ID", tid)
        if not targets:
            notes.append(f"補正の対象「{tid}」が見つからないため反映していません。")
            continue
        for row in targets:
            if field == "割合":
                new = min(max(num(value, -1), 0.0), 1.0) if num(value, -1) >= 0 else None
                if new is None:
                    notes.append(f"割合「{value}」を数値にできないため反映していません。")
                    break
                changes = {"食べた割合": round(new, 2)}
            elif field == "日時":
                dt = parse_dt(value)
                if not dt or dt > now + timedelta(minutes=10):
                    notes.append(f"日時「{value}」を解釈できないため反映していません。")
                    break
                changes = {"食べた日時": fmt_dt(dt), "集計日": day_key(dt)}
            elif field == "状態" and value in STATUSES:
                changes = {"状態": value}
            else:
                notes.append(f"補正「{field}={value}」は扱えないため反映していません。")
                break
            changes = {col: new for col, new in changes.items() if str(row.get(col)) != str(new)}
            if not changes:
                continue  # 既に同じ値（AIが過去の補正を繰り返した等）→ 何もしない
            for col, new in changes.items():
                w.append("補正履歴", {"日時": stamp, "品目ID": row["品目ID"], "項目": col, "旧値": row.get(col, ""),
                                   "新値": new, "元の発言": said, "情報源": source})
            changes["更新日時"] = stamp
            w.update("食事", row, changes)
            notes.append(f"補正: {row['品名']} → " + "、".join(f"{k}={v}" for k, v in changes.items() if k != "更新日時"))

    # --- 確認候補の判断
    cands = open_candidates(db)
    for d in result.get("candidate_decisions") or []:
        no = int(num(d.get("candidate_no"), 0))
        if not 1 <= no <= len(cands):
            notes.append(f"候補{no} が見つかりません。")
            continue
        c = cands[no - 1]
        if d.get("decision") == "記録":
            eaten = parse_dt(c.get("推定食事日時"))
            if not eaten:
                notes.append(f"候補{no}（{c['品名']}）は食べた日時が不明なため記録できません。日時を教えてください。")
                continue
            meal_id = f"M{now.strftime('%Y%m%d-%H%M%S')}-C{no:02d}"
            row = {"品目ID": meal_id + "-01", "食事ID": meal_id, "状態": ACTIVE_STATUS, "食べた日時": fmt_dt(eaten),
                   "集計日": day_key(eaten), "区分": "不明", "品名": c["品名"], "数量": c["数量"], "食べた割合": 1,
                   "根拠": "標準料理", "確信度": "C", "情報源": f"{c['情報源']}（本人承認）", "元の発言": c["原文"],
                   "登録日時": stamp, "更新日時": stamp, "写真ハッシュ": "", "補足": f"確認候補 {c['候補ID']} から"}
            for k, col in NUTRIENTS:
                row[col] = c.get(col, 0)
            w.append("食事", row)
            recorded.append(row)
            w.update("確認候補", c, {"状態": "記録済", "判断日時": stamp})
        else:
            w.update("確認候補", c, {"状態": "不要", "判断日時": stamp})
            notes.append(f"候補{no}（{c['品名']}）は記録しないことにしました。")

    # --- 体重
    wt = result.get("weight") or {}
    kg = num(wt.get("kg"), 0) if wt.get("kg") is not None else 0
    if 30 <= kg <= 200:
        date = wt.get("date") or day_key(now)
        w.append("体重", {"測定日": date, "体重kg": r1(kg), "情報源": source, "登録日時": stamp, "メモ": said[:100]})
        notes.append(f"体重 {r1(kg)}kg を記録しました（{date}）。")
    elif kg:
        notes.append(f"体重 {kg} は範囲外のため記録していません。")

    return w, recorded, notes, warns


def summary_text(db, now, key=None):
    st = settings(db)
    key = key or day_key(now)
    t = day_totals(db, key)
    rest = st["kcal"] - t["kcal"]
    lines = [f"— {day_label(key)} 6:00〜翌5:59 —"]
    if t["count"] == 0:
        lines.append(f"まだ記録がありません（目標 {comma(st['kcal'])}kcal）")
        return lines
    lines.append(f"摂取 {comma(t['kcal'])}kcal（推定 {comma(t['kcal_low'])}〜{comma(t['kcal_high'])}）／目標 {comma(st['kcal'])}")
    lines.append(f"→ 残り 約{comma(rest)}kcal" if rest >= 0 else f"→ 目標を 約{comma(-rest)}kcal 超過")
    salt_mark = " ⚠" if t["salt"] >= st["salt"] else ""
    lines.append(f"食塩 {r1(t['salt'])}g（目安{st['salt']}g未満）{salt_mark}")
    lines.append(f"たんぱく質 {round(t['protein'])}g／脂質 {round(t['fat'])}g／炭水化物 {round(t['carb'])}g"
                 + (f"／アルコール {round(t['alcohol'])}g" if t["alcohol"] >= 1 else ""))
    return lines


def candidate_lines(db, limit=5):
    cands = open_candidates(db)
    if not cands:
        return []
    lines = [f"📋 確認待ちの候補 {len(cands)}件（話すだけで処理:「候補1は記録」「候補1は不要」）"]
    for i, c in enumerate(cands[:limit], start=1):
        when = c.get("推定食事日時") or "日時不明"
        lines.append(f" 候補{i} {when} {c['品名']} {c['数量']}（{c['情報源']}）" + (" ※重複かも" if c.get("重複候補") else ""))
    return lines


def compose_reply(db, now, result, recorded, notes, warns):
    out = []
    if recorded:
        out.append("✅ 記録しました（推定値）")
        for r in recorded:
            ratio = num(r.get("食べた割合"), 1)
            part = "" if ratio >= 1 else f"（{int(round(ratio * 100))}%）"
            out.append(f"・{r['品名']} {r['数量']}{part} 約{comma(num(r['kcal目安']) * ratio)}kcal"
                       f"（{comma(num(r['kcal下限']) * ratio)}〜{comma(num(r['kcal上限']) * ratio)}）食塩{r1(num(r['食塩g']) * ratio)}g"
                       + (f"［{r['食べた日時'][5:]}・{day_label(r['集計日'])}分］" if r['集計日'] != day_key(now) else ""))
    out += notes
    for q in result.get("needs_confirmation") or []:
        out.append("❓ " + q)
    if result.get("answer"):
        out.append(result["answer"])
    if out:
        out.append("")
    out += summary_text(db, now)
    out += ["⚠ " + x for x in warns]
    st = settings(db)
    if day_totals(db, day_key(now))["salt"] >= st["salt"]:
        out.append("⚠ 今日の食塩が目安を超えています。次の食事は汁物・漬物・タレを控えめに。")
    cl = candidate_lines(db)
    if cl:
        out.append("")
        out += cl
    if result.get("advice"):
        out.append("💬 " + result["advice"])
    out.append(DISCLAIMER)
    return "\n".join(out)


# ---------------------------------------------------------------- Make から呼ぶ入口
def handle(inp):
    """mode=prepare: {db, text, image, sent_at} → {request_body, photo_hash, today}
       mode=apply:   {db, text, image, openai, sent_at} → {has_writes, write_body, reply}"""
    mode = inp.get("mode", "")
    now = parse_dt(inp.get("now")) or now_jst()
    db = load_db(inp.get("db"))
    text = str(inp.get("text") or "")
    image = str(inp.get("image") or "").strip()
    if mode == "prepare":
        req = build_openai_request(db, text, now, image, inp.get("model") or DEFAULT_MODEL)
        return {"request_body": json.dumps(req, ensure_ascii=False), "photo_hash": image_hash(image),
                "today": day_key(now), "version": VERSION}
    if mode == "apply":
        try:
            result = parse_openai_response(inp.get("openai"))
        except Exception as e:
            return {"has_writes": False, "write_body": "",
                    "reply": f"⚠ 解析結果を読めませんでした（{type(e).__name__}）。記録していません。もう一度送ってください。"}
        src = inp.get("source") or ("写真" if image else "ショートカット")
        sheet_ids = load_sheet_ids(inp.get("meta"))
        missing = [name for name in SHEETS if name not in sheet_ids]
        if missing:
            return {"has_writes": False, "write_body": "",
                    "reply": f"⚠ 食事DBのシートが見つかりません（{'、'.join(missing)}）。記録していません。"}
        w, recorded, notes, warns = apply(db, result, now, text, src, image_hash(image), sheet_ids)
        body = w.body()
        reply = compose_reply(db, now, result, recorded, notes, warns)
        return {"has_writes": bool(body), "write_body": json.dumps(body, ensure_ascii=False) if body else "",
                "reply": reply, "reply_json": json.dumps({"reply": reply}, ensure_ascii=False), "version": VERSION}
    if mode == "summary":
        return {"reply": "\n".join(summary_text(db, now) + candidate_lines(db) + [DISCLAIMER])}
    return {"error": f"unknown mode: {mode}"}
# ---- END CORE ----


if __name__ == "__main__":
    import sys
    inp = json.load(sys.stdin)
    print(json.dumps(handle(inp), ensure_ascii=False, indent=1))
