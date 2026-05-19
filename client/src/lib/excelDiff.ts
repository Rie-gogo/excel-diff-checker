/**
 * Excel差分比較ユーティリティ
 * Design: Precision Lab — Swiss Typography × Industrial Precision
 */

import * as XLSX from "xlsx";

export interface SheetData {
  headers: string[];
  rows: Record<string, string>[];
  sheetNames: string[];
}

export interface DiffCell {
  col: string;
  valueA: string;
  valueB: string;
}

export interface DiffRow {
  key: string;
  rowA: Record<string, string> | null;
  rowB: Record<string, string> | null;
  diffs: DiffCell[];
  status: "modified" | "added" | "deleted" | "same";
}

export interface DiffResult {
  headers: string[];
  rows: DiffRow[];
  stats: {
    total: number;
    modified: number;
    added: number;
    deleted: number;
    same: number;
  };
}

/**
 * ArrayBufferからExcelデータを読み込む
 */
export function parseExcel(buffer: ArrayBuffer, sheetIndex = 0): SheetData {
  const workbook = XLSX.read(buffer, { type: "array", cellText: true, cellDates: true });
  const sheetNames = workbook.SheetNames;
  const sheetName = sheetNames[sheetIndex] ?? sheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    return { headers: [], rows: [], sheetNames };
  }

  // ヘッダー行を含む全データをJSON化
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
    // 全列が空の行はスキップ
    if (rawRow.every((v) => v === "" || v === null || v === undefined)) continue;
    const row: Record<string, string> = {};
    headers.forEach((h, idx) => {
      row[h] = String(rawRow[idx] ?? "").trim();
    });
    rows.push(row);
  }

  return { headers, rows, sheetNames };
}

/**
 * 2つのシートデータを比較してDiffResultを返す
 */
export function compareSheets(
  dataA: SheetData,
  dataB: SheetData,
  keyColumns: string[]
): DiffResult {
  // 全ヘッダーを結合（A優先、Bにしかないものを追加）
  const allHeaders = [...dataA.headers];
  for (const h of dataB.headers) {
    if (!allHeaders.includes(h)) allHeaders.push(h);
  }

  // キーを生成する関数
  const makeKey = (row: Record<string, string>) =>
    keyColumns.map((k) => row[k] ?? "").join("\0");

  // Aのマップ
  const mapA = new Map<string, Record<string, string>>();
  for (const row of dataA.rows) {
    const key = makeKey(row);
    mapA.set(key, row);
  }

  // Bのマップ
  const mapB = new Map<string, Record<string, string>>();
  for (const row of dataB.rows) {
    const key = makeKey(row);
    mapB.set(key, row);
  }

  const allKeys = new Set<string>([...Array.from(mapA.keys()), ...Array.from(mapB.keys())]);
  const diffRows: DiffRow[] = [];

  const stats = { total: 0, modified: 0, added: 0, deleted: 0, same: 0 };

  for (const key of Array.from(allKeys)) {
    const rowA = mapA.get(key) ?? null;
    const rowB = mapB.get(key) ?? null;
    stats.total++;

    if (rowA === null) {
      // Bにしかない行（追加）
      diffRows.push({ key, rowA: null, rowB, diffs: [], status: "added" });
      stats.added++;
    } else if (rowB === null) {
      // Aにしかない行（削除）
      diffRows.push({ key, rowA, rowB: null, diffs: [], status: "deleted" });
      stats.deleted++;
    } else {
      // 両方にある行 → セルレベルで比較
      const diffs: DiffCell[] = [];
      for (const col of allHeaders) {
        if (keyColumns.includes(col)) continue; // キー列は比較しない
        const va = rowA[col] ?? "";
        const vb = rowB[col] ?? "";
        if (va !== vb) {
          diffs.push({ col, valueA: va, valueB: vb });
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

  // キー順にソート（元の順序を維持するためAの順序を優先）
  const keyOrder = new Map<string, number>();
  dataA.rows.forEach((row, i) => keyOrder.set(makeKey(row), i));
  dataB.rows.forEach((row, i) => {
    const k = makeKey(row);
    if (!keyOrder.has(k)) keyOrder.set(k, dataA.rows.length + i);
  });
  diffRows.sort((a, b) => (keyOrder.get(a.key) ?? 0) - (keyOrder.get(b.key) ?? 0));

  return { headers: allHeaders, rows: diffRows, stats };
}

/**
 * DiffResultをExcelファイルとしてエクスポート（ピンク背景付き）
 */
export function exportDiffToExcel(
  result: DiffResult,
  keyColumns: string[],
  fileAName: string,
  fileBName: string
): void {
  const workbook = XLSX.utils.book_new();

  // ===== シート1: ファイルAベースの差分表示 =====
  const wsDataA: unknown[][] = [];
  wsDataA.push(result.headers); // ヘッダー行

  for (const row of result.rows) {
    if (row.status === "added") continue; // Aにない行はスキップ
    const rowData = result.headers.map((h) => row.rowA?.[h] ?? "");
    wsDataA.push(rowData);
  }

  const wsA = XLSX.utils.aoa_to_sheet(wsDataA);

  // ===== シート2: ファイルBベースの差分表示 =====
  const wsDataB: unknown[][] = [];
  wsDataB.push(result.headers);

  for (const row of result.rows) {
    if (row.status === "deleted") continue; // Bにない行はスキップ
    const rowData = result.headers.map((h) => row.rowB?.[h] ?? "");
    wsDataB.push(rowData);
  }

  const wsB = XLSX.utils.aoa_to_sheet(wsDataB);

  // ===== シート3: 差分サマリー =====
  const summaryData: unknown[][] = [
    ["比較結果サマリー"],
    [],
    ["ファイルA", fileAName],
    ["ファイルB", fileBName],
    ["キー列", keyColumns.join(", ")],
    [],
    ["ステータス", "件数"],
    ["変更あり", result.stats.modified],
    ["Bのみに存在（追加）", result.stats.added],
    ["Aのみに存在（削除）", result.stats.deleted],
    ["変更なし", result.stats.same],
    ["合計", result.stats.total],
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);

  XLSX.utils.book_append_sheet(workbook, wsA, `${fileAName.slice(0, 20)}_差分`);
  XLSX.utils.book_append_sheet(workbook, wsB, `${fileBName.slice(0, 20)}_差分`);
  XLSX.utils.book_append_sheet(workbook, wsSummary, "比較サマリー");

  XLSX.writeFile(workbook, "差分結果.xlsx");
}
