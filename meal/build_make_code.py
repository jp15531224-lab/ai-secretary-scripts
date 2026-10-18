#!/usr/bin/env python3
"""meal_core.py の CORE 部分を Make「Run code（Python）」用のコード文字列にする。

Make の Python モジュールはコード全体を関数本体として実行し、`input` に入力、`return` で出力を返す。
  python3 build_make_code.py            → コードを表示
  python3 build_make_code.py --sha      → コードの SHA-256（Make に載せた版の確認用）
"""
import hashlib
import os
import sys

MARKER = "# ---- END CORE ----"


def build():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "meal_core.py")
    with open(path, encoding="utf-8") as f:
        src = f.read()
    core = src.split(MARKER)[0].rstrip()
    return core + "\n\nreturn handle(dict(input))\n"


if __name__ == "__main__":
    code = build()
    if "--sha" in sys.argv:
        print(hashlib.sha256(code.encode("utf-8")).hexdigest())
    else:
        sys.stdout.write(code)
