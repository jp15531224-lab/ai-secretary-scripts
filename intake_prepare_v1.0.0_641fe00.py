#!/usr/bin/env python3
"""AI秘書: 画像・PDF取り込みの前処理ツール（AIが画像を「見て」読むための準備）

使い方:
  python3 intake_prepare.py decode <download_file_content の保存先.txt> --out <dir>
      → Drive コネクタの download_file_content が大きすぎてファイル保存された結果(JSON)から
        元ファイルを復元する。{"id","title","mimeType","bytes","sha256","path"} を表示。
  python3 intake_prepare.py decode-b64 <base64テキストのファイル> --id <fileId> --title <名前> --mime <type> --out <dir>
      → 小さいファイルで結果が会話内に直接返った場合用（base64 をファイルに書いてから使う）。
  python3 intake_prepare.py prepare <復元したファイル> --out <dir> [--max-tiles 8] [--max-pages 6]
      → AIが読む用の JPEG（縦長スクショは分割）と PDF のテキストを作る。
        {"kind","sha256","hash16","width","height","exif_datetime","tiles","pdf_pages","pdf_text_path","pdf_text_chars"}
  python3 intake_prepare.py selftest

設計方針:
  - 元ファイルは変更しない（復元・変換は作業フォルダ内だけ）。
  - 重複検出は中身の SHA-256（先頭16桁 = hash16）。同じスクショを2回保存しても同じ値になる。
  - 撮影日時(EXIF)は「元資料の日付」の手がかりにだけ使う。ファイルの保存日時とは区別する。
  - LINE 等の縦長スクリーンショットは、文字が潰れないよう重なり付きで縦に分割する。
  - HEIC は pillow-heif、無ければ ImageMagick(convert) で変換する。
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

VERSION = "1.0.0"
MAX_EDGE = 1568          # 画像認識に渡す長辺の上限(px)
TILE_RATIO = 1.25        # 縦長画像は 幅×1.25 の高さごとに分割
TILE_OVERLAP = 0.08      # 分割の重なり（行の切れ目で文字を落とさない）
EXT = {"image/png": ".png", "image/jpeg": ".jpg", "image/jpg": ".jpg", "image/heic": ".heic",
       "image/heif": ".heic", "application/pdf": ".pdf", "image/webp": ".webp", "image/gif": ".gif",
       "image/tiff": ".tif"}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def sniff(data: bytes) -> str:
    """先頭バイトで実際の形式を判定（拡張子やmimeTypeが違っていても正しく扱う）。"""
    if data.startswith(b"%PDF"):
        return "pdf"
    if data.startswith(b"\x89PNG"):
        return "png"
    if data.startswith(b"\xff\xd8"):
        return "jpeg"
    if data[4:12] in (b"ftypheic", b"ftypheix", b"ftypmif1", b"ftypmsf1", b"ftyphevc", b"ftypheim", b"ftypheis"):
        return "heic"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    if data[:3] == b"GIF":
        return "gif"
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return "tiff"
    return "unknown"


def _safe_id(file_id: str) -> str:
    return "".join(c for c in file_id if c.isalnum() or c in "-_")[:80] or "file"


def write_original(raw: bytes, file_id: str, title: str, mime: str, out: Path) -> dict:
    out.mkdir(parents=True, exist_ok=True)
    ext = EXT.get(mime) or Path(title).suffix.lower() or ".bin"
    p = out / f"{_safe_id(file_id)}{ext}"
    p.write_bytes(raw)
    return {"id": file_id, "title": title, "mimeType": mime, "bytes": len(raw),
            "sha256": sha256_file(p), "path": str(p), "kind": sniff(raw[:32])}


def decode_tool_result(path: Path, out: Path) -> dict:
    d = json.loads(path.read_text(encoding="utf-8"))
    raw = base64.b64decode(d["content"], validate=False)
    return write_original(raw, d.get("id", "file"), d.get("title", ""), d.get("mimeType", ""), out)


def decode_b64(path: Path, file_id: str, title: str, mime: str, out: Path) -> dict:
    text = "".join(path.read_text(encoding="ascii").split())
    return write_original(base64.b64decode(text, validate=True), file_id, title, mime, out)


# ---------------------------------------------------------------- images
def _open_image(path: Path, kind: str, work: Path):
    from PIL import Image  # Pillow はクラウド環境に標準で入っている
    if kind == "heic":
        try:
            import pillow_heif  # noqa: F401
            pillow_heif.register_heif_opener()
        except ImportError:
            conv = work / (path.stem + "_conv.jpg")
            cmd = shutil.which("magick") or shutil.which("convert")
            if not cmd:
                raise RuntimeError("HEIC を変換できません（pillow-heif も ImageMagick も無い）")
            subprocess.run([cmd, str(path), str(conv)], check=True, capture_output=True, timeout=120)
            path = conv
    return Image.open(path)


def _exif_datetime(im) -> str:
    try:
        exif = im.getexif()
        sub = exif.get_ifd(0x8769) if hasattr(exif, "get_ifd") else {}
        v = sub.get(0x9003) or exif.get(0x0132)  # DateTimeOriginal → DateTime
        if v:
            s = str(v).strip()
            return f"{s[0:4]}-{s[5:7]}-{s[8:10]}T{s[11:19]}"  # タイムゾーン情報なし（端末の現地時刻）
    except Exception:
        pass
    return ""


def _fit(im):
    from PIL import Image
    w, h = im.size
    scale = min(1.0, MAX_EDGE / max(w, h))
    if scale < 1.0:
        im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
    return im


def tile_boxes(w: int, h: int, max_tiles: int) -> list[tuple[int, int, int, int]]:
    """縦長画像を重なり付きで分割する範囲。横長・普通の比率は1枚のまま。"""
    tile_h = int(w * TILE_RATIO)
    if h <= tile_h * 1.15:
        return [(0, 0, w, h)]
    step = max(1, int(tile_h * (1 - TILE_OVERLAP)))
    boxes, top = [], 0
    while top < h and len(boxes) < max_tiles:
        bottom = min(h, top + tile_h)
        boxes.append((0, top, w, bottom))
        if bottom >= h:
            break
        top += step
    return boxes


def prepare_image(path: Path, kind: str, out: Path, max_tiles: int) -> dict:
    from PIL import ImageOps
    out.mkdir(parents=True, exist_ok=True)
    im = _open_image(path, kind, out)
    exif_dt = _exif_datetime(im)
    im = ImageOps.exif_transpose(im)
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    w, h = im.size
    boxes = tile_boxes(w, h, max_tiles)
    tiles = []
    for i, box in enumerate(boxes, 1):
        t = _fit(im.crop(box))
        p = out / f"{path.stem}_p{i:02d}.jpg"
        t.save(p, "JPEG", quality=85, optimize=True)
        tiles.append(str(p))
    covered = boxes[-1][3] >= h
    return {"width": w, "height": h, "exif_datetime": exif_dt, "tiles": tiles,
            "truncated": not covered}


# ---------------------------------------------------------------- pdf
def prepare_pdf(path: Path, out: Path, max_pages: int) -> dict:
    out.mkdir(parents=True, exist_ok=True)
    pages = 0
    info = subprocess.run(["pdfinfo", str(path)], capture_output=True, text=True, timeout=60)
    for line in info.stdout.splitlines():
        if line.startswith("Pages:"):
            pages = int(line.split()[1])
    txt = out / f"{path.stem}.txt"
    subprocess.run(["pdftotext", "-layout", "-l", "30", str(path), str(txt)],
                   capture_output=True, timeout=120)
    text = txt.read_text(encoding="utf-8", errors="replace") if txt.exists() else ""
    n = min(pages or 1, max_pages)
    prefix = out / f"{path.stem}_p"
    subprocess.run(["pdftoppm", "-jpeg", "-r", "110", "-f", "1", "-l", str(n), str(path), str(prefix)],
                   capture_output=True, timeout=300)
    from PIL import Image
    tiles = []
    for p in sorted(out.glob(f"{path.stem}_p*.jpg")):
        im = _fit(Image.open(p).convert("RGB"))
        im.save(p, "JPEG", quality=85, optimize=True)
        tiles.append(str(p))
    return {"pdf_pages": pages, "pdf_text_path": str(txt) if text.strip() else "",
            "pdf_text_chars": len("".join(text.split())), "tiles": tiles,
            "truncated": pages > n, "width": 0, "height": 0, "exif_datetime": ""}


def prepare(path: Path, out: Path, max_tiles: int = 8, max_pages: int = 6) -> dict:
    head = path.read_bytes()[:32]
    kind = sniff(head)
    digest = sha256_file(path)
    base = {"file": str(path), "kind": kind, "sha256": digest, "hash16": digest[:16],
            "bytes": path.stat().st_size, "prepare_version": VERSION}
    if kind == "pdf":
        return base | prepare_pdf(path, out, max_pages)
    if kind in ("png", "jpeg", "heic", "webp", "gif", "tiff"):
        return base | prepare_image(path, kind, out, max_tiles)
    raise RuntimeError(f"未対応の形式です（先頭バイト判定: {kind}）")


# ---------------------------------------------------------------- selftest
def selftest() -> None:
    import tempfile
    from PIL import Image, ImageDraw
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        # 縦長スクショ（iPhone 相当 1179x2556）→ 分割される
        im = Image.new("RGB", (1179, 2556), "white")
        ImageDraw.Draw(im).rectangle((0, 2400, 1179, 2556), fill="black")
        src = td / "shot.png"
        im.save(src)
        r = prepare(src, td / "out")
        assert r["kind"] == "png" and len(r["tiles"]) >= 2 and not r["truncated"], r
        for t in r["tiles"]:
            w, h = Image.open(t).size
            assert max(w, h) <= MAX_EDGE, (w, h)
        # 同じ内容は同じハッシュ
        dup = td / "dup.png"
        dup.write_bytes(src.read_bytes())
        assert prepare(dup, td / "out2")["hash16"] == r["hash16"]
        # 普通の写真比率は1枚
        Image.new("RGB", (1600, 1200), "white").save(td / "photo.jpg")
        assert len(prepare(td / "photo.jpg", td / "out3")["tiles"]) == 1
        # decode: ツール結果JSON → 元ファイル（バイト一致）
        tr = td / "tool.txt"
        tr.write_text(json.dumps({"content": base64.b64encode(src.read_bytes()).decode(), "id": "abc_DEF-123",
                                  "mimeType": "image/png", "title": "shot.png"}), encoding="utf-8")
        d = decode_tool_result(tr, td / "dec")
        assert d["sha256"] == sha256_file(src) and d["bytes"] == src.stat().st_size, d
        assert sniff(b"%PDF-1.7") == "pdf" and sniff(b"\x00\x00\x00\x18ftypheic") == "heic"
        # PDF（テキスト層あり）
        if shutil.which("pdftotext") and shutil.which("pdftoppm"):
            pdf = td / "doc.pdf"
            Image.new("RGB", (800, 1100), "white").save(pdf)
            rp = prepare(pdf, td / "out4")
            assert rp["kind"] == "pdf" and rp["pdf_pages"] == 1 and len(rp["tiles"]) == 1, rp
    print("selftest OK", VERSION)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    d = sub.add_parser("decode")
    d.add_argument("tool_result")
    d.add_argument("--out", required=True)
    b = sub.add_parser("decode-b64")
    b.add_argument("b64file")
    b.add_argument("--id", required=True)
    b.add_argument("--title", default="")
    b.add_argument("--mime", default="")
    b.add_argument("--out", required=True)
    p = sub.add_parser("prepare")
    p.add_argument("file")
    p.add_argument("--out", required=True)
    p.add_argument("--max-tiles", type=int, default=8)
    p.add_argument("--max-pages", type=int, default=6)
    sub.add_parser("selftest")
    a = ap.parse_args(argv)
    try:
        if a.cmd == "selftest":
            selftest()
            return 0
        if a.cmd == "decode":
            r = decode_tool_result(Path(a.tool_result), Path(a.out))
        elif a.cmd == "decode-b64":
            r = decode_b64(Path(a.b64file), a.id, a.title, a.mime, Path(a.out))
        else:
            r = prepare(Path(a.file), Path(a.out), a.max_tiles, a.max_pages)
    except Exception as e:  # 失敗理由を JSON で返す（Routine が処理ログに書く）
        print(json.dumps({"ok": False, "error": f"{type(e).__name__}: {e}"}, ensure_ascii=False))
        return 1
    print(json.dumps({"ok": True} | r, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
