/**
 * Excel差分比較ユーティリティ
 * Design: "Precision Lab" — Swiss Typography × Industrial Precision
 *
 * 機能:
 * - 列マッピング: ファイルA・Bの列を任意にペアリング
 * - キーマッピング: キー列もA・B間でマッピング指定
 * - 正規化比較: 半角/全角同一視、スペース（半角・全角）無視
 */

import * as XLSX from "xlsx";

// ============================
// 型定義
// ============================

export interface SheetData {
  headers: string[];
  rows: Record<string, string>[];
  sheetNames: string[];
}

/**
 * 列マッピング: colA（ファイルAの列名）と colB（ファイルBの列名）のペア
 */
export interface ColumnMapping {
  colA: string;
  colB: string;
  /** 表示ラベル（省略時は colA を使用） */
  label?: string;
}

/**
 * キーマッピング: キーとなる列のA・Bペア
 */
export interface KeyMapping {
  colA: string;
  colB: string;
}

export interface DiffCell {
  /** 表示列名（ColumnMappingのlabel or colA） */
  col: string;
  valueA: string;
  valueB: string;
  /** 正規化後の値（デバッグ用） */
  normalizedA: string;
  normalizedB: string;
}

export interface DiffRow {
  key: string;
  rowA: Record<string, string> | null;
  rowB: Record<string, string> | null;
  diffs: DiffCell[];
  status: "modified" | "added" | "deleted" | "same";
}

export interface DiffResult {
  /** 表示用ヘッダー（ColumnMappingのlabel or colA） */
  headers: string[];
  rows: DiffRow[];
  stats: {
    total: number;
    modified: number;
    added: number;
    deleted: number;
    same: number;
  };
  /** 使用したキーマッピング */
  keyMappings: KeyMapping[];
  /** 使用した列マッピング */
  columnMappings: ColumnMapping[];
}

// ============================
// 正規化処理
// ============================

/**
 * 全角→半角変換テーブル（英数字・記号・スペース）
 */
function toHalfWidth(str: string): string {
  return str
    // 全角英数字・記号 → 半角
    .replace(/[\uFF01-\uFF5E]/g, (ch) =>
      String.fromCharCode(ch.charCodeAt(0) - 0xFF01 + 0x21)
    )
    // 全角スペース → 半角スペース
    .replace(/\u3000/g, " ");
}

/**
 * 比較用の正規化:
 * 1. 全角→半角変換
 * 2. スペース（半角・全角）を除去
 * 3. 大文字→小文字（オプション: 現在は保持）
 */
export function normalize(value: string): string {
  if (!value) return "";
  const half = toHalfWidth(value.trim());
  // スペースをすべて除去（半角スペースは toHalfWidth 後に除去）
  return half.replace(/\s+/g, "");
}

// ============================
// Excel読み込み
// ============================

export function parseExcel(buffer: ArrayBuffer, sheetIndex = 0): SheetData {
  const workbook = XLSX.read(buffer, { type: "array", cellText: true, cellDates: true });
  const sheetNames = workbook.SheetNames;
  const sheetName = sheetNames[sheetIndex] ?? sheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    return { headers: [], rows: [], sheetNames };
  }

  const rawData = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
  }) as string[][];

  if (rawData.length === 0) {
    return { headers: [], rows: [], sheetNames };
  }

  const headers = rawData[0].map((h) => String(h ?? "").trim());
  const rows: Record<string, string>[] = [];

  for (let i = 1; i < rawData.length; i++) {
    const rawRow = rawData[i] as string[];
    if (rawRow.every((v) => v === "" || v === null || v === undefined)) continue;
    const row: Record<string, string> = {};
    headers.forEach((h, idx) => {
      row[h] = String(rawRow[idx] ?? "").trim();
    });
    rows.push(row);
  }

  return { headers, rows, sheetNames };
}

// ============================
// 列マッピングのデフォルト生成
// ============================

/**
 * 2つのヘッダーリストから、名前が一致する列のデフォルトマッピングを生成する
 * （正規化した名前で比較）
 */
export function buildDefaultColumnMappings(
  headersA: string[],
  headersB: string[]
): ColumnMapping[] {
  const mappings: ColumnMapping[] = [];
  const usedB = new Set<string>();

  for (const colA of headersA) {
    const normA = normalize(colA);
    // Bの中で正規化後が一致する列を探す
    const matchB = headersB.find(
      (b) => !usedB.has(b) && normalize(b) === normA
    );
    if (matchB) {
      mappings.push({ colA, colB: matchB, label: colA });
      usedB.add(matchB);
    }
  }
  return mappings;
}

/**
 * デフォルトのキーマッピングを生成（最初の共通列）
 */
export function buildDefaultKeyMappings(
  headersA: string[],
  headersB: string[]
): KeyMapping[] {
  const defaultMappings = buildDefaultColumnMappings(headersA, headersB);
  return defaultMappings.length > 0
    ? [{ colA: defaultMappings[0].colA, colB: defaultMappings[0].colB }]
    : [];
}

// ============================
// 差分比較
// ============================

/**
 * 2つのシートデータを列マッピング・キーマッピングを使って比較する
 */
export function compareSheets(
  dataA: SheetData,
  dataB: SheetData,
  keyMappings: KeyMapping[],
  columnMappings: ColumnMapping[]
): DiffResult {
  // 表示ヘッダー（キー列 + 比較列）
  const keyLabels = keyMappings.map((km) => km.colA);
  const compareLabels = columnMappings
    .filter((cm) => !keyMappings.some((km) => km.colA === cm.colA))
    .map((cm) => cm.label ?? cm.colA);
  const allHeaders = [...keyLabels, ...compareLabels];

  // キー生成関数
  const makeKeyA = (row: Record<string, string>) =>
    keyMappings.map((km) => normalize(row[km.colA] ?? "")).join("\0");
  const makeKeyB = (row: Record<string, string>) =>
    keyMappings.map((km) => normalize(row[km.colB] ?? "")).join("\0");

  // マップ構築
  const mapA = new Map<string, Record<string, string>>();
  for (const row of dataA.rows) {
    const key = makeKeyA(row);
    if (key.replace(/\0/g, "") !== "") mapA.set(key, row);
  }

  const mapB = new Map<string, Record<string, string>>();
  for (const row of dataB.rows) {
    const key = makeKeyB(row);
    if (key.replace(/\0/g, "") !== "") mapB.set(key, row);
  }

  const allKeys = new Set<string>([
    ...Array.from(mapA.keys()),
    ...Array.from(mapB.keys()),
  ]);

  const diffRows: DiffRow[] = [];
  const stats = { total: 0, modified: 0, added: 0, deleted: 0, same: 0 };

  // 比較列マッピング（キー列を除く）
  const compareMappings = columnMappings.filter(
    (cm) => !keyMappings.some((km) => km.colA === cm.colA)
  );

  for (const key of Array.from(allKeys)) {
    const rowA = mapA.get(key) ?? null;
    const rowB = mapB.get(key) ?? null;
    stats.total++;

    if (rowA === null) {
      diffRows.push({ key, rowA: null, rowB, diffs: [], status: "added" });
      stats.added++;
    } else if (rowB === null) {
      diffRows.push({ key, rowA, rowB: null, diffs: [], status: "deleted" });
      stats.deleted++;
    } else {
      const diffs: DiffCell[] = [];
      for (const cm of compareMappings) {
        const va = rowA[cm.colA] ?? "";
        const vb = rowB[cm.colB] ?? "";
        const na = normalize(va);
        const nb = normalize(vb);
        if (na !== nb) {
          diffs.push({
            col: cm.label ?? cm.colA,
            valueA: va,
            valueB: vb,
            normalizedA: na,
            normalizedB: nb,
          });
        }
      }
      if (diffs.length > 0) {
        diffRows.push({ key, rowA, rowB, diffs, status: "modified" });
        stats.modified++;
      } else {
        diffRows.push({ key, rowA, rowB, diffs: [], status: "same" });
        stats.same++;
      }
    }
  }

  // 元の順序でソート（A優先）
  const keyOrder = new Map<string, number>();
  dataA.rows.forEach((row, i) => keyOrder.set(makeKeyA(row), i));
  dataB.rows.forEach((row, i) => {
    const k = makeKeyB(row);
    if (!keyOrder.has(k)) keyOrder.set(k, dataA.rows.length + i);
  });
  diffRows.sort((a, b) => (keyOrder.get(a.key) ?? 0) - (keyOrder.get(b.key) ?? 0));

  return {
    headers: allHeaders,
    rows: diffRows,
    stats,
    keyMappings,
    columnMappings,
  };
}

// ============================
// Excelエクスポート
// ============================

/** ARGB形式のカラー文字列 */
const COLOR = {
  headerBg: "FF3D4A6B",       // インディゴ（ヘッダー背景）
  headerBgLight: "FF5C6B99",  // 薄インディゴ（比較列グループ）
  headerBgA: "FF7B8EC0",      // A列サブヘッダー
  headerBgB: "FFD63384",      // B列サブヘッダー（ピンク）
  headerFg: "FFFFFFFF",       // ヘッダー文字（白）
  diffBg: "FFFCE4EC",         // 差分セル背景（ライトピンク）
  addedBg: "FFE8F5E9",        // 追加行背景（ライトグリーン）
  deletedBg: "FFFCE4EC",      // 削除行背景（ライトピンク）
  keyBg: "FF3D4A6B",          // キー列ヘッダー背景
  statusBg: "FF3D4A6B",       // 状態列ヘッダー背景
  white: "FFFFFFFF",
  black: "FF000000",
};

function cellStyle(opts: {
  bgColor?: string;
  fgColor?: string;
  bold?: boolean;
  strike?: boolean;
  wrapText?: boolean;
  hAlign?: "left" | "center" | "right";
}) {
  return {
    fill: opts.bgColor ? { patternType: "solid", fgColor: { rgb: opts.bgColor } } : undefined,
    font: {
      color: opts.fgColor ? { rgb: opts.fgColor } : undefined,
      bold: opts.bold ?? false,
      strike: opts.strike ?? false,
    },
    alignment: {
      horizontal: opts.hAlign ?? "left",
      vertical: "center",
      wrapText: opts.wrapText ?? false,
    },
    border: {
      top:    { style: "thin", color: { rgb: "FFD0D0D0" } },
      bottom: { style: "thin", color: { rgb: "FFD0D0D0" } },
      left:   { style: "thin", color: { rgb: "FFD0D0D0" } },
      right:  { style: "thin", color: { rgb: "FFD0D0D0" } },
    },
  };
}

function setCellAt(
  ws: XLSX.WorkSheet,
  r: number,
  c: number,
  value: string | number,
  style: ReturnType<typeof cellStyle>
) {
  const addr = XLSX.utils.encode_cell({ r, c });
  ws[addr] = { v: value, t: typeof value === "number" ? "n" : "s", s: style };
}

function buildDiffSheet(
  result: DiffResult,
  fileAName: string,
  fileBName: string,
  onlyDiffs: boolean
): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {};

  const keyLabels = result.keyMappings.map((km) => km.colA);
  const compareMappings = result.columnMappings.filter(
    (cm) => !result.keyMappings.some((km) => km.colA === cm.colA)
  );

  // ---- 列インデックス計算 ----
  // 状態(1) + キー列(n) + 比較列×2(m*2)
  const statusCol = 0;
  const keyStartCol = 1;
  const compareStartCol = keyStartCol + keyLabels.length;

  // ---- 行1: グループヘッダー ----
  // 状態列（rowSpan=2 → 結合）
  setCellAt(ws, 0, statusCol, "状態", cellStyle({ bgColor: COLOR.headerBg, fgColor: COLOR.headerFg, bold: true, hAlign: "center" }));
  // キー列（rowSpan=2 → 結合）
  keyLabels.forEach((h, i) => {
    setCellAt(ws, 0, keyStartCol + i, `${h} [KEY]`, cellStyle({ bgColor: COLOR.headerBg, fgColor: COLOR.headerFg, bold: true }));
  });
  // 比較列グループ（列名 + colSpan=2）
  compareMappings.forEach((cm, i) => {
    const col = compareStartCol + i * 2;
    const label = cm.label ?? cm.colA;
    setCellAt(ws, 0, col, label, cellStyle({ bgColor: COLOR.headerBgLight, fgColor: COLOR.headerFg, bold: true, hAlign: "center" }));
    setCellAt(ws, 0, col + 1, "", cellStyle({ bgColor: COLOR.headerBgLight, fgColor: COLOR.headerFg, bold: true }));
  });

  // ---- 行2: ファイル名サブヘッダー ----
  setCellAt(ws, 1, statusCol, "", cellStyle({ bgColor: COLOR.headerBg, fgColor: COLOR.headerFg }));
  keyLabels.forEach((_, i) => {
    setCellAt(ws, 1, keyStartCol + i, "", cellStyle({ bgColor: COLOR.headerBg, fgColor: COLOR.headerFg }));
  });
  compareMappings.forEach((_, i) => {
    const col = compareStartCol + i * 2;
    setCellAt(ws, 1, col,     fileAName, cellStyle({ bgColor: COLOR.headerBgA, fgColor: COLOR.headerFg, bold: true, hAlign: "center" }));
    setCellAt(ws, 1, col + 1, fileBName, cellStyle({ bgColor: COLOR.headerBgB, fgColor: COLOR.headerFg, bold: true, hAlign: "center" }));
  });

  // ---- セル結合 ----
  const merges: XLSX.Range[] = [];
  // 状態列: 行1-2結合
  merges.push({ s: { r: 0, c: statusCol }, e: { r: 1, c: statusCol } });
  // キー列: 行1-2結合
  keyLabels.forEach((_, i) => {
    merges.push({ s: { r: 0, c: keyStartCol + i }, e: { r: 1, c: keyStartCol + i } });
  });
  // 比較列グループ: 列2つ結合
  compareMappings.forEach((_, i) => {
    const col = compareStartCol + i * 2;
    merges.push({ s: { r: 0, c: col }, e: { r: 0, c: col + 1 } });
  });
  ws["!merges"] = merges;

  // ---- データ行 ----
  let rowIdx = 2;
  const rows = onlyDiffs
    ? result.rows.filter((r) => r.status !== "same")
    : result.rows;

  for (const row of rows) {
    const isAdded = row.status === "added";
    const isDeleted = row.status === "deleted";
    const isModified = row.status === "modified";
    const rowBg = isAdded ? COLOR.addedBg : isDeleted ? COLOR.deletedBg : undefined;
    const diffCols = new Set(row.diffs.map((d) => d.col));

    // 状態
    const statusLabel = isAdded ? "追加" : isDeleted ? "削除" : isModified ? "変更" : "同一";
    const statusBg = isAdded ? "FF4CAF50" : isDeleted ? "FFE91E63" : isModified ? "FFE91E63" : "FF9E9E9E";
    setCellAt(ws, rowIdx, statusCol, statusLabel, cellStyle({ bgColor: statusBg, fgColor: COLOR.white, bold: true, hAlign: "center" }));

    // キー列
    keyLabels.forEach((h, i) => {
      const km = result.keyMappings.find((k) => k.colA === h);
      const val = km ? (row.rowA?.[km.colA] ?? row.rowB?.[km.colB] ?? "") : "";
      setCellAt(ws, rowIdx, keyStartCol + i, val, cellStyle({ bgColor: rowBg, bold: true }));
    });

    // 比較列（A列・B列横並び）
    compareMappings.forEach((cm, i) => {
      const col = compareStartCol + i * 2;
      const label = cm.label ?? cm.colA;
      const isDiff = diffCols.has(label);
      const valA = row.rowA?.[cm.colA] ?? "";
      const valB = row.rowB?.[cm.colB] ?? "";
      const cellBgA = isDiff ? COLOR.diffBg : rowBg;
      const cellBgB = isDiff ? COLOR.diffBg : rowBg;

      if (isAdded) {
        setCellAt(ws, rowIdx, col,     "–", cellStyle({ bgColor: rowBg }));
        setCellAt(ws, rowIdx, col + 1, valB, cellStyle({ bgColor: rowBg }));
      } else if (isDeleted) {
        setCellAt(ws, rowIdx, col,     valA, cellStyle({ bgColor: cellBgA, strike: true }));
        setCellAt(ws, rowIdx, col + 1, "–",  cellStyle({ bgColor: rowBg }));
      } else {
        // same or modified
        setCellAt(ws, rowIdx, col,     valA, cellStyle({ bgColor: cellBgA, strike: isDiff }));
        setCellAt(ws, rowIdx, col + 1, valB, cellStyle({ bgColor: cellBgB }));
      }
    });

    rowIdx++;
  }

  // ---- 列幅設定 ----
  const totalCols = compareStartCol + compareMappings.length * 2;
  const colWidths: { wch: number }[] = [];
  colWidths.push({ wch: 8 }); // 状態
  keyLabels.forEach(() => colWidths.push({ wch: 18 }));
  compareMappings.forEach(() => {
    colWidths.push({ wch: 24 });
    colWidths.push({ wch: 24 });
  });
  ws["!cols"] = colWidths;

  // ---- シート範囲 ----
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: rowIdx - 1, c: totalCols - 1 },
  });

  return ws;
}

export function exportDiffToExcel(
  result: DiffResult,
  fileAName: string,
  fileBName: string
): void {
  const workbook = XLSX.utils.book_new();

  // 差分行のみシート
  const wsDiff = buildDiffSheet(result, fileAName, fileBName, true);
  XLSX.utils.book_append_sheet(workbook, wsDiff, "差分のみ");

  // 全行シート
  const wsAll = buildDiffSheet(result, fileAName, fileBName, false);
  XLSX.utils.book_append_sheet(workbook, wsAll, "全行");

  // サマリーシート
  const keyMappingStr = result.keyMappings
    .map((km) => `${km.colA} ↔ ${km.colB}`)
    .join(", ");
  const colMappingStr = result.columnMappings
    .map((cm) => `${cm.colA} ↔ ${cm.colB}`)
    .join(", ");

  const summaryData: unknown[][] = [
    ["比較結果サマリー"],
    [],
    ["ファイルA", fileAName],
    ["ファイルB", fileBName],
    ["キーマッピング", keyMappingStr],
    ["列マッピング", colMappingStr],
    [],
    ["ステータス", "件数"],
    ["変更あり", result.stats.modified],
    ["Bのみに存在（追加）", result.stats.added],
    ["Aのみに存在（削除）", result.stats.deleted],
    ["変更なし", result.stats.same],
    ["合計", result.stats.total],
    [],
    ["※ 半角・全角は同一視、スペースは無視して比較しています"],
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(workbook, wsSummary, "比較サマリー");

  XLSX.writeFile(workbook, "差分結果.xlsx");
}
