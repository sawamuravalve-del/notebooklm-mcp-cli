/**
 * 毎月DB → Dataverse 同期用 読み取りスクリプト（Office Scripts）
 *
 * Power Automate の「スクリプトの実行」(Excel Online (Business)) から呼び出す。
 * 1リクエストで以下をすべて処理し、SyncRow[] を JSON で返す。
 *
 *   1. テーブル T_DB が無ければ作成する（冪等）
 *   2. 半角カナ・日本語のヘッダーを ASCII 論理名へ正規化する
 *   3. 日付シリアル値を "YYYY-MM-DD" に変換する
 *   4. 空白セル・不可視文字を null に正規化する
 *   5. 受注No + 手配票No ごとに出現順で sub_no を採番し、同期キーを生成する
 *   6. 業務列59本から行ハッシュ（FNV-1a 16桁hex）を生成する
 *
 * 同期キー:
 *   受注No + 手配票No + sub_no を区切り無しで連結する。
 *     13524 + R08-0234 + 01  →  "13524R08-023401"
 *   sub_no は「受注No + 手配票No」単位の通し番号。
 *
 * 前提（確認済み）:
 *   - ファイルの行順は変わらない  → 出現順採番で安定する
 *   - 同一の受注No が複数の月ファイルに出現しない → 同期キーに年月は不要
 *   - 明細数は99以下 → sub_no は2桁で足りる
 *
 * 注意:
 *   COL_MAP のキーの並び順が行ハッシュの計算順を決める。
 *   並び替えると全行のハッシュが変わり、現在月の全行が更新対象になる。
 *   列の追加・削除も同様。変更する場合は影響を理解した上で行うこと。
 *
 *   本スクリプトは Dataverse / SharePoint リストのどちらを同期先にしても
 *   そのまま使える。出力は論理名のみで、接頭辞はフロー側で補う。
 */

interface SyncRow { [key: string]: string | number | null }

const TABLE_NAME = "T_DB";
const SUB_NO_DIGITS = 2;
const MAX_LINES_PER_GROUP = 99;   // 受注No + 手配票No あたりの明細上限

/** Excel列名 → Dataverse論理名（接頭辞なし。フロー側で crfff_ を補う） */
const COL_MAP: { [k: string]: string } = {
  "ｻｲｽﾞ": "size",
  "データ区分": "data_kubun",
  "ﾌﾟﾛｼﾞｪｸﾄ名コード": "project_code",
  "外税額": "tax_excl_amount",
  "規格・型番": "spec_model",
  "原価": "cost",
  "原価税込区分": "cost_taxincl_kubun",
  "原価税率": "cost_tax_rate",
  "受注No": "order_no",
  "受注金額": "order_amount",
  "受注数": "order_qty",
  "受注日": "order_date",
  "商品コード": "item_code",
  "商品項目１": "item_attr1",
  "商品項目２": "item_attr2",
  "商品項目３": "item_attr3",
  "商品名": "item_name",
  "商品名２": "item_name2",
  "色": "color",
  "税込区分": "taxincl_kubun",
  "税率": "tax_rate",
  "粗利益": "gross_profit",
  "単位": "unit",
  "単位区分": "unit_kubun",
  "単価": "unit_price",
  "担当者コード": "staff_code",
  "直送先コード": "dropship_code",
  "直送先名１": "dropship_name1",
  // 旧「摘要」。手配票No へ改称し、同期キーの構成要素に昇格した。
  // Dataverse 側の論理名は crfff_RNo（実際の論理名は自動的に小文字化され crfff_rno になる）
  "手配票No": "RNo",
  "得意先コード": "customer_code",
  "得意先名１": "customer_name1",
  "内税額": "tax_incl_amount",
  "入数": "pack_qty",
  "納期": "delivery_date",
  "売価金額": "sales_amount",
  "売上項目１": "sales_attr1",
  "売上項目２": "sales_attr2",
  "売上項目３": "sales_attr3",
  "売単価": "sales_unit_price",
  "箱数": "box_qty",
  "番船": "shipment_no",
  "備考": "note",
  "標準価格": "standard_price",
  "部門コード": "dept_code",
  "スポット得意先名１": "spot_customer_name1",
  "スポット得意先名２": "spot_customer_name2",
  "ﾌﾟﾛｼﾞｪｸﾄ名名": "project_name",
  "軽減税率マーク": "reduced_tax_mark",
  "見積No": "quote_no",
  "原価軽減税率マーク": "cost_reduced_tax_mark",
  "受注残高": "backlog_amount",
  "受注残数": "backlog_qty",
  "出荷マーク": "ship_mark",
  "単位区分１受注残数": "unit1_backlog_qty",
  "単位区分１受注数": "unit1_order_qty",
  "単位区分１単位": "unit1_unit",
  "直送先名２": "dropship_name2",
  "得意先名２": "customer_name2",
  "利益率": "profit_rate",
};

/** 日付として解釈する列 */
const DATE_COLS = ["受注日", "納期"];

/** 数値として解釈する列 */
const NUM_COLS = [
  "外税額", "原価", "原価税率", "受注金額", "受注数", "税率", "粗利益",
  "単価", "内税額", "入数", "売価金額", "売単価", "箱数", "標準価格",
  "受注残高", "受注残数", "単位区分１受注残数", "単位区分１受注数", "利益率",
];

/** 先行ゼロを保持するため、数値に見えても文字列として扱う列 */
const FORCE_TEXT_COLS = [
  "受注No", "手配票No", "商品コード", "得意先コード", "担当者コード", "部門コード",
  "直送先コード", "見積No", "番船", "ﾌﾟﾛｼﾞｪｸﾄ名コード",
];

// ---------------------------------------------------------------------------

function main(workbook: ExcelScript.Workbook): SyncRow[] {
  const table = ensureTable(workbook);
  const values = table.getRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(v => String(v).trim());
  const idx: { [k: string]: number } = {};
  for (let i = 0; i < header.length; i++) idx[header[i]] = i;

  // 列の欠落は無言で null を生むため、ここで必ず止める
  const missing: string[] = [];
  for (const jp of Object.keys(COL_MAP)) {
    if (idx[jp] === undefined) missing.push(jp);
  }
  if (missing.length > 0) {
    throw new Error("列が見つかりません: " + missing.join(", "));
  }

  const rows: SyncRow[] = [];
  const counter: { [groupKey: string]: number } = {};
  const seenSyncKeys: { [syncKey: string]: number } = {};

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const excelRow = r + 1;

    // ---- キー列（受注No / 手配票No）----
    const orderNoRaw = normText(row[idx["受注No"]]);
    if (orderNoRaw === null) {
      if (isEmptyRow(row)) continue;                 // 完全な空行はスキップ
      throw new Error("受注Noが空です: Excel " + excelRow + " 行目");
    }
    const orderNo = toKeyPart(orderNoRaw, "受注No", excelRow);
    const tehaiNo = toKeyPart(normText(row[idx["手配票No"]]), "手配票No", excelRow);

    // ---- 業務列59本の正規化とハッシュ材料の収集 ----
    const rec: SyncRow = {};
    const parts: string[] = [];

    for (const jp of Object.keys(COL_MAP)) {
      const raw = row[idx[jp]];
      let out: string | number | null;

      if (DATE_COLS.indexOf(jp) >= 0) {
        out = toIsoDate(raw, jp, excelRow);
      } else if (FORCE_TEXT_COLS.indexOf(jp) >= 0) {
        out = normText(raw);
      } else if (NUM_COLS.indexOf(jp) >= 0) {
        out = toNum(raw, jp, excelRow);
      } else {
        out = normText(raw);
      }

      rec[COL_MAP[jp]] = out;
      parts.push(out === null ? "" : String(out));
    }

    // ---- 出現順採番（受注No + 手配票No 単位）----
    // 採番の単位を同期キーの構成と揃えている。受注No 単位にしたい場合は
    // groupKey を orderNo だけにすればよい。
    const groupKey = orderNo + "\u001f" + tehaiNo;
    const prev = counter[groupKey] === undefined ? 0 : counter[groupKey];
    const seq = prev + 1;
    counter[groupKey] = seq;
    if (seq > MAX_LINES_PER_GROUP) {
      throw new Error(
        "受注No " + orderNo + " / 手配票No " + tehaiNo + " の明細が " + MAX_LINES_PER_GROUP +
        " 行を超えました: Excel " + excelRow + " 行目。SUB_NO_DIGITS の見直しが必要です。");
    }

    // 同期キーは3要素を区切り無しで連結する（例: 13524 + R08-0234 + 01 → 13524R08-023401）。
    // 区切りが無いため、桁数の組み合わせ次第では別の行と同じ文字列になりうる。
    // Dataverse 側の代替キー制約でも検知できるが、ここで先に止めたほうが原因が分かりやすい。
    const syncKey = orderNo + tehaiNo + padLeft(String(seq), SUB_NO_DIGITS);
    if (seenSyncKeys[syncKey] !== undefined) {
      throw new Error(
        "同期キーが重複しました: " + syncKey +
        " (Excel " + seenSyncKeys[syncKey] + " 行目 と " + excelRow + " 行目)。" +
        "受注No と 手配票No を区切り無しで連結しているため、桁数の組み合わせにより衝突しています。");
    }
    seenSyncKeys[syncKey] = excelRow;

    rec["order_no"] = orderNo;
    rec["sub_no"] = seq;                                          // Dataverse は整数列
    rec["sync_key"] = syncKey;
    // 区切りは業務データに現れない制御文字(UNIT SEPARATOR)。区切り無しで連結すると
    // "a|bc" と "ab|c" が同じハッシュになるため必須。
    rec["row_hash"] = hash64(parts.join("\u001f"));   // sub_no は含めない

    rows.push(rec);
  }

  return rows;
}

// --- ヘルパー ---------------------------------------------------------------

/** テーブル T_DB が無ければ先頭シートの使用範囲から作成する（冪等） */
function ensureTable(workbook: ExcelScript.Workbook): ExcelScript.Table {
  const existing = workbook.getTable(TABLE_NAME);
  if (existing) return existing;

  const sheet = workbook.getWorksheets()[0];
  const used = sheet.getUsedRange();
  if (!used) throw new Error("データがありません: " + sheet.getName());

  const table = workbook.addTable(used, true);   // true = 先頭行をヘッダーとして扱う
  table.setName(TABLE_NAME);
  return table;
}

function isEmptyRow(row: (string | number | boolean)[]): boolean {
  for (const v of row) {
    if (v !== "" && v !== null && String(v).trim() !== "") return false;
  }
  return true;
}

/**
 * 文字列の正規化。
 * trim() は全角スペース(U+3000)とNBSP(U+00A0)を除去するが、
 * ZWSP(U+200B)などの書式カテゴリの不可視文字は残るため明示的に落とす。
 */
function normText(v: string | number | boolean): string | null {
  if (v === null || v === undefined) return null;
  const s = String(v).replace(/[\u200b-\u200d\ufeff]/g, "").trim();
  return s === "" ? null : s;
}

/** 数値の正規化。解釈できない値は null に丸めず例外にする */
function toNum(v: string | number | boolean, col: string, excelRow: number): number | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v;

  const s = String(v).replace(/[,\s]/g, "").trim();
  if (s === "") return null;

  const n = Number(s);
  if (isNaN(n)) {
    throw new Error("数値として解釈できません: " + col + " / Excel " + excelRow + " 行目 / 値=" + v);
  }
  return n;
}

/**
 * Excelシリアル値 → "YYYY-MM-DD"。25569 = 1970-01-01 のシリアル値。
 * 「未定」「-」などの文字列を無言で null にすると Excel と同期先が
 * 静かに乖離するため、解釈できない値は例外にする。
 */
function toIsoDate(v: string | number | boolean, col: string, excelRow: number): string | null {
  if (v === null || v === undefined || v === "") return null;

  if (typeof v === "number") {
    if (v <= 0) return null;                        // 0 は 1899-12-30 になるので弾く
    return new Date(Math.round((v - 25569) * 86400000)).toISOString().substring(0, 10);
  }

  const s = String(v).trim().replace(/\//g, "-");
  if (s === "") return null;
  if (!/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) {
    throw new Error("日付として解釈できません: " + col + " / Excel " + excelRow + " 行目 / 値=" + s);
  }
  const p = s.split("-");
  return p[0] + "-" + padLeft(p[1], 2) + "-" + padLeft(p[2], 2);
}

/** 同期キーは Power Automate 側で JSON 辞書のキーになるため文字種を限定する */
function sanitizeKey(s: string): string {
  return s.replace(/[^0-9A-Za-z\-_]/g, "");
}

/**
 * 同期キーの構成要素（受注No / 手配票No）を検証して返す。
 * 空を許すとキーの形が静かに変わって別の行と衝突しうるため、必ず例外にする。
 */
function toKeyPart(raw: string | null, colName: string, excelRow: number): string {
  if (raw === null) {
    throw new Error(
      colName + "が空です: Excel " + excelRow + " 行目。同期キーの構成要素は空にできません。");
  }
  const s = sanitizeKey(raw);
  if (s !== raw) {
    throw new Error(
      colName + "にキーとして使えない文字が含まれます: Excel " + excelRow + " 行目 / 値=" + raw +
      "。使用できるのは英数字・ハイフン・アンダースコアのみです。");
  }
  return s;
}

function padLeft(s: string, n: number): string {
  let out = s;
  while (out.length < n) out = "0" + out;
  return out;
}

/** FNV-1a を順方向・逆方向で2回かけ、16桁hex(≒64bit相当)にする */
function hash64(s: string): string {
  return fnv1a(s) + fnv1a(reverseStr(s));
}

function fnv1a(s: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return ("00000000" + h.toString(16)).slice(-8);
}

function reverseStr(s: string): string {
  return s.split("").reverse().join("");
}
