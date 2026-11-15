#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
引合検索ブック 生成スクリプト

毎月DB の Excel ファイル群を読み込み、引合検索用の1ブックに統合する。
設計書「毎月DB(Excel) → Dataverse 引合検索システム」2-3 の11列に対応。

使い方:
    python build-search-book.py <入力フォルダ> [出力フォルダ]

    例:
        python build-search-book.py data\\sample out
        python build-search-book.py data\\archive out

出力:
    <出力フォルダ>\\引合検索.xlsx    検索用ブック
    <出力フォルダ>\\品質レポート.md   取り込み時に検出した問題の一覧

必要なパッケージ:
    pip install openpyxl XlsxWriter
"""

import sys
import os
import re
import glob
from datetime import datetime, timedelta, date
from collections import defaultdict

try:
    from openpyxl import load_workbook
    import xlsxwriter
except ImportError:
    sys.exit("必要なパッケージがありません。次を実行してください:\n"
             "    pip install openpyxl XlsxWriter")

# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------

# (正式な列名, 別名リスト, 種別)
#   code = コード。先行ゼロを保持したい文字列
#   text = 自由文字列
#   date = 日付
#   num  = 数値
#
# 別名は、ファイルによって列名が揺れている場合の受け皿。
# 「摘要」は手配票No への改称前のファイルを読むために入れてある。
COLUMN_SPEC = [
    ("受注No",       [],                        "code"),
    ("手配票No",     ["摘要"],                  "code"),
    ("受注日",       [],                        "date"),
    ("納期",         [],                        "date"),
    ("得意先コード", [],                        "code"),
    ("得意先名１",   ["得意先名1"],             "text"),
    ("商品コード",   [],                        "code"),
    ("商品名",       [],                        "text"),
    ("受注数",       [],                        "num"),
    ("単価",         [],                        "num"),
    ("受注金額",     [],                        "num"),
]

COLUMNS   = [c[0] for c in COLUMN_SPEC]
KIND_OF   = {c[0]: c[2] for c in COLUMN_SPEC}
# 別名 → 正式名
ALIAS_MAP = {}
for name, aliases, _ in COLUMN_SPEC:
    ALIAS_MAP[name] = name
    for a in aliases:
        ALIAS_MAP[a] = name

KEY_COLUMNS      = ["受注No", "手配票No"]   # 同期キーの構成要素
MAX_LINES_PER_GROUP = 99                    # sub_no は2桁

# 目に見えない文字。Web からの貼り付けなどで混入する
INVISIBLE = re.compile(r"[​-‍﻿]")
# 同期キーに使えない文字。
# 当初は英数字とハイフンのみに限っていたが、実データの手配票No には
# 「R08-0073-預-08」のように日本語が入る。日本語は JSON のキーにも
# Dataverse の代替キーにも使えるため、本当に危険な文字だけを弾く。
KEY_UNSAFE = re.compile(r"[\"\\/<>*%&:?|#\x00-\x1f]")

EXCEL_EPOCH = datetime(1899, 12, 30)


# ---------------------------------------------------------------------------
# 値の正規化
# ---------------------------------------------------------------------------

def norm_text(v):
    """文字列の正規化。全角スペースと不可視文字を落とす。"""
    if v is None:
        return None
    if isinstance(v, str):
        s = INVISIBLE.sub("", v).strip()
    else:
        s = str(v).strip()
    return s or None


def norm_code(v):
    """
    コード列の正規化。
    数値として格納されていた場合は先行ゼロが失われているため、
    その事実を呼び出し側に返す。
    """
    if v is None:
        return None, False
    numeric_origin = isinstance(v, (int, float)) and not isinstance(v, bool)
    if numeric_origin and float(v).is_integer():
        s = str(int(v))
    else:
        s = norm_text(v)
    return (s or None), numeric_origin


def norm_date(v):
    """
    日付の正規化。実データは YYYYMMDD の8桁（数値または文字列）で入っている。
    Excel のシリアル値・datetime・区切り付き文字列にも一応対応しておく。
    """
    if v is None or v == "":
        return None, None
    if isinstance(v, datetime):
        return v.date(), None
    if isinstance(v, date):
        return v, None

    is_num = isinstance(v, (int, float)) and not isinstance(v, bool)

    # 数値は整数文字列に直してから判定する（20200127.0 のような値に備える）
    if is_num:
        if float(v).is_integer():
            s = str(int(v))
        else:
            s = None
    else:
        s = norm_text(v)

    # YYYYMMDD（8桁）。実データの主形式
    if s and re.fullmatch(r"\d{8}", s):
        try:
            return date(int(s[:4]), int(s[4:6]), int(s[6:8])), None
        except ValueError:
            return None, s

    # Excel のシリアル値（通常5桁程度）
    if is_num:
        if v <= 0:
            return None, None
        if 0 < float(v) < 1000000:
            try:
                return (EXCEL_EPOCH + timedelta(days=float(v))).date(), None
            except (OverflowError, ValueError):
                return None, repr(v)
        return None, repr(v)

    if s is None:
        return None, None

    # 2026-11-02 / 2026/11/02 など
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3))), None
        except ValueError:
            return None, s
    return None, s        # 「未定」「-」など、日付として読めない値


def norm_num(v):
    """数値の正規化。"""
    if v is None or v == "":
        return None, None
    if isinstance(v, bool):
        return None, repr(v)
    if isinstance(v, (int, float)):
        return float(v), None
    s = norm_text(v)
    if s is None:
        return None, None
    s2 = s.replace(",", "").replace("　", "").strip()
    try:
        return float(s2), None
    except ValueError:
        return None, s


# ---------------------------------------------------------------------------
# 読み込み
# ---------------------------------------------------------------------------

def target_files(folder):
    """対象の xlsx を列挙する。一時ファイルと競合コピーは除外。"""
    out = []
    for p in sorted(glob.glob(os.path.join(folder, "*.xlsx"))):
        name = os.path.basename(p)
        if name.startswith("~$"):
            continue
        if re.search(r"-(DESKTOP|LAPTOP|PC)-", name, re.IGNORECASE):
            continue
        out.append(p)
    return out


def map_header(header_cells, path, issues):
    """
    ヘッダー行から、必要な11列の位置を特定する。
    見つからない列があればファイル単位で報告し、そのファイルを飛ばす。
    """
    pos = {}
    for i, raw in enumerate(header_cells):
        h = norm_text(raw)
        if h is None:
            continue
        canon = ALIAS_MAP.get(h)
        if canon and canon not in pos:
            pos[canon] = i
    missing = [c for c in COLUMNS if c not in pos]
    if missing:
        issues.append({
            "kind": "列が見つからない",
            "file": os.path.basename(path),
            "detail": "、".join(missing),
        })
        return None
    return pos


def read_file(path, issues, numeric_codes):
    """1ファイルを読み、正規化済みの行を返す。"""
    rows = []
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]        # 先頭シート
        it = ws.iter_rows(values_only=True)
        try:
            header = next(it)
        except StopIteration:
            issues.append({"kind": "空ファイル", "file": os.path.basename(path), "detail": ""})
            return rows, ws.title

        pos = map_header(header, path, issues)
        if pos is None:
            return rows, ws.title

        fname = os.path.basename(path)
        for rn, raw in enumerate(it, start=2):
            if raw is None or all(c is None or str(c).strip() == "" for c in raw):
                continue                                   # 空行

            rec = {"_file": fname, "_row": rn}
            for col in COLUMNS:
                i = pos[col]
                v = raw[i] if i < len(raw) else None
                kind = KIND_OF[col]

                if kind == "date":
                    val, bad = norm_date(v)
                    if bad is not None:
                        issues.append({"kind": "日付として読めない",
                                       "file": fname, "row": rn,
                                       "detail": f"{col} = {bad}"})
                elif kind == "num":
                    val, bad = norm_num(v)
                    if bad is not None:
                        issues.append({"kind": "数値として読めない",
                                       "file": fname, "row": rn,
                                       "detail": f"{col} = {bad}"})
                elif kind == "code":
                    val, numeric_origin = norm_code(v)
                    if numeric_origin:
                        # 行ごとに出すと件数が膨れるため、列単位で数えて後でまとめる
                        numeric_codes[(fname, col)] += 1
                else:
                    val = norm_text(v)

                rec[col] = val
            rows.append(rec)
    finally:
        wb.close()
    return rows, ws.title


# ---------------------------------------------------------------------------
# キーの採番と検証
# ---------------------------------------------------------------------------

def assign_keys(rows, issues):
    """受注No + 手配票No ごとに出現順で sub_no を採番し、同期キーを作る。"""
    counter = defaultdict(int)
    seen = {}
    for rec in rows:
        order_no = rec["受注No"]
        tehai_no = rec["手配票No"]

        for col, val in (("受注No", order_no), ("手配票No", tehai_no)):
            if val is None:
                issues.append({"kind": "キーが空",
                               "file": rec["_file"], "row": rec["_row"],
                               "detail": f"{col} が空"})
            elif KEY_UNSAFE.search(val):
                issues.append({"kind": "キーに使えない文字",
                               "file": rec["_file"], "row": rec["_row"],
                               "detail": f"{col} = {val}"})

        if order_no is None or tehai_no is None:
            rec["_sync_key"] = None
            continue

        gk = f"{order_no}\x1f{tehai_no}"
        counter[gk] += 1
        seq = counter[gk]
        if seq > MAX_LINES_PER_GROUP:
            issues.append({"kind": "明細数が上限超過",
                           "file": rec["_file"], "row": rec["_row"],
                           "detail": f"受注No {order_no} / 手配票No {tehai_no} が "
                                     f"{MAX_LINES_PER_GROUP} 行を超えた"})
        rec["_sub_no"] = seq
        key = f"{order_no}{tehai_no}{seq:02d}"
        rec["_sync_key"] = key

        if key in seen:
            prev = seen[key]
            issues.append({"kind": "同期キーの重複",
                           "file": rec["_file"], "row": rec["_row"],
                           "detail": f"{key} が {prev['_file']} {prev['_row']}行目 と衝突"})
        else:
            seen[key] = rec


# ---------------------------------------------------------------------------
# 出力
# ---------------------------------------------------------------------------

def write_workbook(rows, out_path):
    """検索用ブックを書き出す。"""
    # constant_memory は add_table() と併用できないため使わない。
    # 10万行 × 11列で数百MB程度のメモリで収まる。
    wb = xlsxwriter.Workbook(out_path, {"default_date_format": "yyyy/mm/dd"})
    ws = wb.add_worksheet("検索")

    f_header = wb.add_format({"bold": True, "bg_color": "#EDF0F7",
                              "border": 1, "border_color": "#C6CEDF"})
    f_date   = wb.add_format({"num_format": "yyyy/mm/dd"})
    f_qty    = wb.add_format({"num_format": "#,##0.###"})
    f_price  = wb.add_format({"num_format": "#,##0.####"})
    f_amt    = wb.add_format({"num_format": "#,##0.##"})
    f_text   = wb.add_format({})

    fmt_of = {"受注日": f_date, "納期": f_date, "受注数": f_qty,
              "単価": f_price, "受注金額": f_amt}

    widths = {"受注No": 12, "手配票No": 14, "受注日": 12, "納期": 12,
              "得意先コード": 13, "得意先名１": 28, "商品コード": 14,
              "商品名": 34, "受注数": 10, "単価": 12, "受注金額": 14}
    for i, col in enumerate(COLUMNS):
        ws.set_column(i, i, widths.get(col, 14), fmt_of.get(col, f_text))

    ws.write_row(0, 0, COLUMNS, f_header)
    for r, rec in enumerate(rows, start=1):
        for c, col in enumerate(COLUMNS):
            v = rec.get(col)
            if v is None:
                continue
            if isinstance(v, date):
                ws.write_datetime(r, c, datetime(v.year, v.month, v.day), f_date)
            else:
                ws.write(r, c, v, fmt_of.get(col, f_text))

    ws.add_table(0, 0, max(len(rows), 1), len(COLUMNS) - 1, {
        "name": "T_検索",
        "style": "Table Style Light 9",
        "columns": [{"header": c} for c in COLUMNS],
    })
    ws.freeze_panes(1, 0)
    wb.close()


def write_report(rows, files, sheet_names, issues, out_path, src_folder):
    """品質レポートを書き出す。"""
    by_kind = defaultdict(list)
    for it in issues:
        by_kind[it["kind"]].append(it)

    L = []
    L.append("# 品質レポート — 引合検索ブック\n")
    L.append(f"- 生成日時: {datetime.now():%Y-%m-%d %H:%M}")
    L.append(f"- 入力フォルダ: `{src_folder}`")
    L.append(f"- 対象ファイル数: {len(files)}")
    L.append(f"- 取り込み行数: {len(rows):,}\n")

    L.append("## ファイル別の行数\n")
    per_file = defaultdict(int)
    for rec in rows:
        per_file[rec["_file"]] += 1
    L.append("| ファイル | シート名 | 行数 |")
    L.append("|---|---|---|")
    for p in files:
        n = os.path.basename(p)
        L.append(f"| {n} | {sheet_names.get(n, '—')} | {per_file.get(n, 0):,} |")
    L.append("")

    L.append("## 列ごとの充足率\n")
    L.append("検索に使う11列に、実際どれだけ値が入っているか。")
    L.append("充足率が低い列は、検索結果に出しても空欄が並ぶ。\n")
    L.append("| 列 | 非空 | 充足率 |")
    L.append("|---|---|---|")
    n = max(len(rows), 1)
    for col in COLUMNS:
        filled = sum(1 for r in rows if r.get(col) is not None)
        pct = filled / n * 100
        mark = " ←空" if filled == 0 else (" ←低" if pct < 30 else "")
        L.append(f"| {col} | {filled:,} | {pct:.1f}%{mark} |")
    L.append("")

    if not issues:
        L.append("## 検出された問題\n")
        L.append("**問題は検出されませんでした。**\n")
    else:
        L.append(f"## 検出された問題（合計 {len(issues):,} 件）\n")
        L.append("| 種別 | 件数 |")
        L.append("|---|---|")
        for kind in sorted(by_kind, key=lambda k: -len(by_kind[k])):
            L.append(f"| {kind} | {len(by_kind[kind]):,} |")
        L.append("")

        for kind in sorted(by_kind, key=lambda k: -len(by_kind[k])):
            items = by_kind[kind]
            L.append(f"### {kind}（{len(items):,} 件）\n")
            L.append("| ファイル | 行 | 内容 |")
            L.append("|---|---|---|")
            for it in items[:50]:
                L.append(f"| {it.get('file','')} | {it.get('row','')} | {it.get('detail','')} |")
            if len(items) > 50:
                L.append(f"\n> 先頭50件のみ表示。残り {len(items)-50:,} 件は省略。")
            L.append("")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(L))


# ---------------------------------------------------------------------------

def main(argv):
    if len(argv) < 2:
        sys.exit(__doc__)
    src = argv[1]
    out_dir = argv[2] if len(argv) > 2 else "out"

    if not os.path.isdir(src):
        sys.exit(f"入力フォルダがありません: {src}")
    os.makedirs(out_dir, exist_ok=True)

    files = target_files(src)
    if not files:
        sys.exit(f"対象の .xlsx がありません: {src}")

    print(f"対象ファイル: {len(files)} 件")
    issues, rows, sheet_names = [], [], {}
    numeric_codes = defaultdict(int)
    for i, p in enumerate(files, 1):
        got, sheet = read_file(p, issues, numeric_codes)
        sheet_names[os.path.basename(p)] = sheet
        rows.extend(got)
        print(f"  [{i}/{len(files)}] {os.path.basename(p)}  {len(got):,} 行  (シート: {sheet})")

    if not rows:
        sys.exit("取り込めた行がありません。品質レポートを確認してください。")

    # コード列が数値として格納されていたファイルを、列単位でまとめて報告する
    for (fname, col), cnt in sorted(numeric_codes.items()):
        issues.append({"kind": "コード列が数値で格納されている",
                       "file": fname, "row": "",
                       "detail": f"{col} … {cnt:,} 行。先行ゼロがある値なら失われている"})

    assign_keys(rows, issues)

    # 受注日 降順。日付が無い行は末尾へ
    rows.sort(key=lambda r: (r["受注日"] is None,
                             -(r["受注日"].toordinal() if r["受注日"] else 0)))

    book   = os.path.join(out_dir, "引合検索.xlsx")
    report = os.path.join(out_dir, "品質レポート.md")
    write_workbook(rows, book)
    write_report(rows, files, sheet_names, issues, report, os.path.abspath(src))

    print(f"\n合計 {len(rows):,} 行")
    print(f"検出した問題: {len(issues):,} 件")
    print(f"\n出力:\n  {book}\n  {report}")
    if issues:
        print("\n問題が検出されています。品質レポートを確認してください。")


if __name__ == "__main__":
    main(sys.argv)
