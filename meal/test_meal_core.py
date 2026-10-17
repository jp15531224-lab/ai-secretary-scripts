#!/usr/bin/env python3
"""meal_core.py のテスト（python3 -m unittest meal/test_meal_core.py）"""
import json
import os
import sys
import unittest
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
import meal_core as mc  # noqa: E402

NOW = datetime(2026, 10, 17, 21, 0, tzinfo=mc.JST)


def sheet(name, rows):
    cols = mc.SHEETS[name]
    return {"range": f"'{name}'!A1:Z1000", "values": [cols] + [[r.get(c, "") for c in cols] for r in rows]}


def make_db(meals=(), cands=(), weights=(), as_filter=False):
    vrs = [sheet("食事", meals), sheet("補正履歴", []), sheet("確認候補", cands), sheet("体重", weights),
           sheet("設定", [{"項目": k, "値": v, "説明": d} for k, v, d in mc.DEFAULT_SETTINGS])]
    if as_filter:
        vrs = [{"valueRange": v, "dataFilters": []} for v in vrs]
    return {"valueRanges": vrs}


def meal(item_id, eaten, name, kcal, salt=1.0, ratio=1, status="確定", photo=""):
    dt = mc.parse_dt(eaten)
    return {"品目ID": item_id, "食事ID": item_id.rsplit("-", 1)[0], "状態": status, "食べた日時": eaten,
            "集計日": mc.day_key(dt), "品名": name, "数量": "1", "食べた割合": ratio, "kcal目安": kcal,
            "kcal下限": kcal * 0.8, "kcal上限": kcal * 1.2, "食塩g": salt, "たんぱく質g": 10, "脂質g": 5,
            "炭水化物g": 50, "アルコールg": 0, "情報源": "ショートカット", "登録日時": eaten, "写真ハッシュ": photo}


def item(name, kcal, salt=1.0, ratio=1.0, eaten_at="", ref="A"):
    return {"meal_ref": ref, "eaten_at": eaten_at, "time_basis": "発言" if eaten_at else "送信時刻", "slot": "夜",
            "name": name, "quantity": "1", "ratio": ratio, "kcal": kcal, "kcal_low": kcal * 0.85,
            "kcal_high": kcal * 1.2, "salt": salt, "protein": 20, "fat": 10, "carb": 60, "alcohol": 0,
            "basis": "標準料理", "confidence": "C", "note": ""}


def result(**kw):
    base = {"intent": "記録", "new_items": [], "amendments": [], "candidate_decisions": [],
            "weight": {"kg": None, "date": ""}, "needs_confirmation": [], "answer": "", "advice": ""}
    base.update(kw)
    return base


def openai_body(res):
    return {"choices": [{"message": {"content": json.dumps(res, ensure_ascii=False)}}]}


class DayKey(unittest.TestCase):
    def test_six_am_boundary(self):
        self.assertEqual(mc.day_key(mc.parse_dt("2026-10-17 06:00")), "2026-10-17")
        self.assertEqual(mc.day_key(mc.parse_dt("2026-10-18 05:59:59")), "2026-10-17")
        self.assertEqual(mc.day_key(mc.parse_dt("2026-10-18 02:00")), "2026-10-17")
        self.assertEqual(mc.day_key(mc.parse_dt("2026-10-17 05:59")), "2026-10-16")
        self.assertEqual(mc.day_key(mc.parse_dt("2026-10-18 06:00")), "2026-10-18")

    def test_parse_utc_and_offsets(self):
        self.assertEqual(mc.fmt_dt(mc.parse_dt("2026-10-17T03:00:00Z")), "2026-10-17 12:00")
        self.assertEqual(mc.fmt_dt(mc.parse_dt("2026-10-17T12:00:00+09:00")), "2026-10-17 12:00")
        self.assertEqual(mc.fmt_dt(mc.parse_dt("2026/10/17 7:05")), "2026-10-17 07:05")
        self.assertIsNone(mc.parse_dt("昨日"))
        self.assertIsNone(mc.parse_dt("2026-02-30 10:00"))


class LoadAndTotals(unittest.TestCase):
    def test_load_both_response_shapes(self):
        for shape in (False, True):
            db = mc.load_db(make_db([meal("M1-01", "2026-10-17 12:00", "カツ丼", 950)], as_filter=shape))
            self.assertEqual(len(db["食事"].rows), 1)
            self.assertEqual(db["食事"].rows[0]["__row"], 2)

    def test_ratio_status_and_night(self):
        db = mc.load_db(make_db([
            meal("M1-01", "2026-10-17 12:00", "カツ丼", 1000, salt=4),
            meal("M2-01", "2026-10-17 19:00", "パスタ", 800, ratio=0.5),
            meal("M3-01", "2026-10-17 20:00", "友達のピザ", 900, status="本人以外"),
            meal("M4-01", "2026-10-18 02:00", "ラーメン", 500, salt=6),   # 深夜2時 → 10/17分
            meal("M5-01", "2026-10-17 05:00", "前日分", 300),            # 早朝5時 → 10/16分
        ]))
        t = mc.day_totals(db, "2026-10-17")
        self.assertEqual(t["count"], 3)
        self.assertAlmostEqual(t["kcal"], 1000 + 400 + 500)
        self.assertAlmostEqual(t["salt"], 4 + 0.5 + 6)
        self.assertEqual(mc.day_totals(db, "2026-10-16")["count"], 1)

    def test_week_average_excludes_unrecorded_days(self):
        db = mc.load_db(make_db([meal("M1-01", "2026-10-17 12:00", "A", 2000),
                                 meal("M2-01", "2026-10-15 12:00", "B", 1600)]))
        ws = mc.week_stats(db, "2026-10-17")
        self.assertEqual(ws["recorded_days"], 2)
        self.assertAlmostEqual(ws["avg"]["kcal"], 1800)


class Apply(unittest.TestCase):
    def run_apply(self, db_dict, res, text="", photo=""):
        db = mc.load_db(db_dict)
        w, rec, notes, warns = mc.apply(db, res, NOW, text, "ショートカット", photo)
        return db, w, rec, notes, warns

    def test_new_items_rows_and_ids(self):
        db, w, rec, _, _ = self.run_apply(make_db(), result(new_items=[item("焼き鳥 もも", 300), item("焼き鳥 むね", 140)]),
                                          "焼き鳥もも3本、胸肉2本")
        self.assertEqual(len(rec), 2)
        self.assertEqual(rec[0]["食事ID"], rec[1]["食事ID"])
        self.assertEqual([r["品目ID"][-2:] for r in rec], ["01", "02"])
        reqs = w.body()["requests"]
        self.assertEqual([list(r) for r in reqs], [["appendCells"], ["appendCells"]])
        cells = reqs[0]["appendCells"]["rows"][0]["values"]
        self.assertEqual(len(cells), len(mc.MEAL_COLS))
        self.assertEqual(cells[mc.MEAL_COLS.index("集計日")], {"userEnteredValue": {"stringValue": "2026-10-17"}})
        self.assertEqual(cells[mc.MEAL_COLS.index("kcal目安")], {"userEnteredValue": {"numberValue": 300}})

    def test_late_night_goes_to_previous_day(self):
        now = datetime(2026, 10, 18, 10, 0, tzinfo=mc.JST)
        db = mc.load_db(make_db())
        _, rec, _, _ = mc.apply(db, result(new_items=[item("ラーメン", 500, eaten_at="2026-10-18 02:00")]), now)
        self.assertEqual(rec[0]["集計日"], "2026-10-17")
        self.assertIn("10/17", mc.compose_reply(db, now, result(), rec, [], []))

    def test_future_time_is_clamped(self):
        _, _, rec, _, warns = self.run_apply(make_db(), result(new_items=[item("カツ丼", 950, eaten_at="2026-10-17 23:30")]))
        self.assertEqual(rec[0]["食べた日時"], "2026-10-17 21:00")
        self.assertTrue(any("未来" in x for x in warns))

    def test_bad_numbers_are_sanitized(self):
        it = item("謎", 500)
        it.update({"kcal_low": 900, "kcal_high": 100, "ratio": 5, "salt": -3})
        _, _, rec, _, _ = self.run_apply(make_db(), result(new_items=[it]))
        r = rec[0]
        self.assertTrue(r["kcal下限"] <= r["kcal目安"] <= r["kcal上限"])
        self.assertEqual(r["食べた割合"], 1.0)
        self.assertEqual(r["食塩g"], 0)

    def test_same_photo_not_recorded_twice(self):
        db_dict = make_db([meal("M1-01", "2026-10-17 12:00", "刺身", 200, photo="abc123")])
        _, w, rec, notes, _ = self.run_apply(db_dict, result(new_items=[item("刺身", 200)]), photo="abc123")
        self.assertEqual(rec, [])
        self.assertIsNone(w.body())
        self.assertTrue(any("登録済み" in n for n in notes))

    def test_near_duplicate_warns_but_records(self):
        db_dict = make_db([meal("M1-01", "2026-10-17 20:30", "カツ丼", 950)])
        _, _, rec, _, warns = self.run_apply(db_dict, result(new_items=[item("カツ丼", 950)]))
        self.assertEqual(len(rec), 1)
        self.assertTrue(any("取消" in x for x in warns))

    def test_amend_ratio_meal_and_history(self):
        db_dict = make_db([meal("M1-01", "2026-10-17 19:00", "刺身", 200), meal("M1-02", "2026-10-17 19:00", "天ぷら", 500)])
        db, w, rec, notes, _ = self.run_apply(db_dict, result(intent="補正", amendments=[
            {"target_id": "M1-01", "field": "割合", "value": "0.5", "reason": "刺身は半分"}]), "刺身は半分、天ぷらは全部")
        self.assertEqual(rec, [])
        reqs = w.body()["requests"]
        self.assertTrue(any("appendCells" in r and r["appendCells"]["sheetId"] == 1 for r in reqs))  # 補正履歴
        upd = [r["updateCells"]["start"] for r in reqs if "updateCells" in r]
        self.assertIn({"sheetId": 0, "rowIndex": 1, "columnIndex": mc.MEAL_COLS.index("食べた割合")}, upd)
        self.assertAlmostEqual(mc.day_totals(db, "2026-10-17")["kcal"], 100 + 500)

    def test_amend_whole_meal_status_and_time(self):
        db_dict = make_db([meal("M1-01", "2026-10-17 19:00", "ピザ", 900), meal("M1-02", "2026-10-17 19:00", "サラダ", 100)])
        db, w, _, _, _ = self.run_apply(db_dict, result(amendments=[
            {"target_id": "M1", "field": "状態", "value": "本人以外", "reason": ""}]))
        self.assertEqual(mc.day_totals(db, "2026-10-17")["count"], 0)
        db, w, _, _, _ = self.run_apply(db_dict, result(amendments=[
            {"target_id": "M1-01", "field": "日時", "value": "2026-10-16 20:00", "reason": "昨日食べた"}]))
        self.assertEqual(mc.day_totals(db, "2026-10-16")["count"], 1)
        row = db["食事"].find("品目ID", "M1-01")[0]
        self.assertEqual(row["集計日"], "2026-10-16")

    def test_amend_noop_is_ignored(self):
        db_dict = make_db([meal("M1-01", "2026-10-17 12:00", "おにぎり", 180, status="取消")])
        _, w, _, notes, _ = self.run_apply(db_dict, result(amendments=[
            {"target_id": "M1-01", "field": "状態", "value": "取消", "reason": ""}]))
        self.assertIsNone(w.body())
        self.assertEqual(notes, [])

    def test_amend_unknown_target(self):
        _, w, _, notes, _ = self.run_apply(make_db(), result(amendments=[
            {"target_id": "M9", "field": "割合", "value": "0.5", "reason": ""}]))
        self.assertIsNone(w.body())
        self.assertTrue(any("見つからない" in n for n in notes))

    def test_candidate_record_and_reject(self):
        cands = [
            {"候補ID": "C1", "状態": "保留", "情報源": "PLAUD", "推定食事日時": "2026-10-16 20:00", "品名": "焼き鳥",
             "数量": "5本", "kcal目安": 450, "kcal下限": 350, "kcal上限": 550, "食塩g": 2.5, "原文": "昨日焼き鳥5本食べた",
             "登録日時": "2026-10-17 10:00"},
            {"候補ID": "C2", "状態": "保留", "情報源": "ボイスメモ", "推定食事日時": "", "品名": "ビール", "数量": "2杯",
             "kcal目安": 340, "登録日時": "2026-10-17 11:00"},
            {"候補ID": "C0", "状態": "記録済", "品名": "済み", "登録日時": "2026-10-17 09:00"},
        ]
        db, w, rec, notes, _ = self.run_apply(make_db(cands=cands), result(candidate_decisions=[
            {"candidate_no": 1, "decision": "記録"}, {"candidate_no": 2, "decision": "不要"}]))
        self.assertEqual(len(rec), 1)
        self.assertEqual(rec[0]["集計日"], "2026-10-16")
        self.assertEqual(rec[0]["kcal目安"], 450)
        statuses = {c["候補ID"]: c["状態"] for c in db["確認候補"].rows}
        self.assertEqual(statuses, {"C1": "記録済", "C2": "不要", "C0": "記録済"})

    def test_weight(self):
        db, w, _, notes, _ = self.run_apply(make_db(), result(weight={"kg": 77.6, "date": ""}), "体重77.6")
        self.assertEqual(db["体重"].rows[0]["体重kg"], 77.6)
        self.assertEqual(db["体重"].rows[0]["測定日"], "2026-10-17")
        _, w, _, notes, _ = self.run_apply(make_db(), result(weight={"kg": 7.76, "date": ""}))
        self.assertIsNone(w.body())


class ReplyAndRequest(unittest.TestCase):
    def test_reply_contains_totals_and_remaining(self):
        db = mc.load_db(make_db([meal("M1-01", "2026-10-17 12:00", "カツ丼", 950, salt=4)]))
        _, rec, notes, warns = mc.apply(db, result(new_items=[item("ラーメン", 500, salt=6)]), NOW)
        reply = mc.compose_reply(db, NOW, result(advice="夜は汁を残しましょう"), rec, notes, warns)
        self.assertIn("1,450kcal", reply)
        self.assertIn("残り 約350kcal", reply)
        self.assertIn("食塩 10.0g", reply)
        self.assertIn("⚠", reply)
        self.assertIn(mc.DISCLAIMER, reply)

    def test_request_schema_is_strict_compatible(self):
        req = mc.build_openai_request(mc.load_db(make_db()), "カツ丼食べた", NOW, image_b64="QUJD")
        schema = req["response_format"]["json_schema"]["schema"]

        def check(node):
            if node.get("type") == "object":
                self.assertFalse(node["additionalProperties"])
                self.assertEqual(sorted(node["required"]), sorted(node["properties"]))
                for v in node["properties"].values():
                    check(v)
            if node.get("type") == "array":
                check(node["items"])
        check(schema)
        self.assertEqual(req["messages"][1]["content"][1]["type"], "image_url")
        self.assertIn("今日の集計日: 2026-10-17", req["messages"][1]["content"][0]["text"])

    def test_handle_round_trip(self):
        dbj = json.dumps(make_db(as_filter=True), ensure_ascii=False)
        meta = {"sheets": [{"properties": {"title": n, "sheetId": 100 + i}} for i, n in enumerate(mc.SHEETS)]}
        prep = mc.handle({"mode": "prepare", "db": dbj, "text": "カツ丼食べた", "now": "2026-10-17 21:00"})
        self.assertIn("gpt", json.loads(prep["request_body"])["model"])
        out = mc.handle({"mode": "apply", "db": dbj, "text": "カツ丼食べた", "now": "2026-10-17 21:00", "meta": meta,
                         "openai": openai_body(result(new_items=[item("カツ丼", 950, salt=4)]))})
        self.assertTrue(out["has_writes"])
        self.assertEqual(json.loads(out["write_body"])["requests"][0]["appendCells"]["sheetId"], 100)
        nometa = mc.handle({"mode": "apply", "db": dbj, "openai": openai_body(result(new_items=[item("カツ丼", 950)]))})
        self.assertFalse(nometa["has_writes"])
        self.assertIn("カツ丼", out["reply"])
        bad = mc.handle({"mode": "apply", "db": dbj, "openai": {"choices": [{"message": {"content": "not json"}}]}})
        self.assertFalse(bad["has_writes"])
        q = mc.handle({"mode": "apply", "db": dbj, "now": "2026-10-17 21:00", "meta": meta,
                       "openai": openai_body(result(intent="質問"))})
        self.assertFalse(q["has_writes"])
        self.assertIn("まだ記録がありません", q["reply"])


class Profile(unittest.TestCase):
    def test_profile_from_settings_and_weight(self):
        d = make_db(weights=[{"測定日": "2026-10-10", "体重kg": 78.2}])
        rows = d["valueRanges"][4]["values"]
        for r in rows:
            if r[0] == "年齢":
                r[1] = 43
        req = mc.build_openai_request(mc.load_db(d), "x", NOW)
        sysmsg = req["messages"][0]["content"]
        self.assertIn("43歳", sysmsg)
        self.assertIn("直近の体重78.2kg", sysmsg)
        self.assertNotIn("{profile}", sysmsg)


class MakeCodeBuild(unittest.TestCase):
    def test_dist_is_in_sync(self):
        import build_make_code as b
        with open(os.path.join(os.path.dirname(__file__), "dist", "make_code.py"), encoding="utf-8") as f:
            self.assertEqual(f.read(), b.build(), "dist/make_code.py が古い。python3 build_make_code.py > dist/make_code.py")

    def test_core_runs_as_function_body(self):
        import build_make_code as b
        code = b.build()
        self.assertNotIn("__main__", code)
        fn_src = "def _make(input):\n" + "\n".join("    " + l for l in code.splitlines())
        ns = {}
        exec(fn_src, ns)
        dbj = json.dumps(make_db(), ensure_ascii=False)
        out = ns["_make"]({"mode": "summary", "db": dbj, "now": "2026-10-17 21:00"})
        self.assertIn("まだ記録がありません", out["reply"])


if __name__ == "__main__":
    unittest.main()
