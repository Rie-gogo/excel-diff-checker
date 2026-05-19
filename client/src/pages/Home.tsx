/**
 * Home.tsx — Excel差分チェッカー メインページ
 * Design: "Precision Lab" — Swiss Typography × Industrial Precision
 * Layout: Left sidebar (step nav) + Main content area (3-step flow)
 */

import { useState, useCallback, useRef } from "react";
import { parseExcel, compareSheets, exportDiffToExcel } from "@/lib/excelDiff";
import type { SheetData, DiffResult } from "@/lib/excelDiff";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { toast } from "sonner";
import {
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  Circle,
  ArrowRight,
  RotateCcw,
  Download,
  AlertTriangle,
  Info,
  X,
  ChevronDown,
} from "lucide-react";

// ============================
// Types
// ============================
type Step = 1 | 2 | 3;

interface FileState {
  file: File | null;
  data: SheetData | null;
  selectedSheet: number;
  loading: boolean;
  error: string | null;
}

const initFileState = (): FileState => ({
  file: null,
  data: null,
  selectedSheet: 0,
  loading: false,
  error: null,
});

// ============================
// Sub-components
// ============================

/** ステップインジケーター（サイドバー用） */
function StepIndicator({
  step,
  current,
  label,
  sublabel,
}: {
  step: Step;
  current: Step;
  label: string;
  sublabel: string;
}) {
  const done = current > step;
  const active = current === step;

  return (
    <div className="flex items-start gap-3">
      <div className="flex flex-col items-center">
        <div
          className={`w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold transition-all duration-300 ${
            done
              ? "bg-emerald-500 text-white"
              : active
              ? "bg-white text-[#3B3F8C] shadow-lg shadow-white/20"
              : "bg-white/10 text-white/40"
          }`}
        >
          {done ? <CheckCircle2 className="w-4 h-4" /> : step}
        </div>
        {step < 3 && (
          <div
            className={`w-0.5 h-10 mt-1 transition-all duration-500 ${
              done ? "bg-emerald-500/60" : "bg-white/10"
            }`}
          />
        )}
      </div>
      <div className="pt-1">
        <p
          className={`text-sm font-semibold transition-colors duration-300 ${
            active ? "text-white" : done ? "text-emerald-400" : "text-white/40"
          }`}
        >
          {label}
        </p>
        <p
          className={`text-xs mt-0.5 transition-colors duration-300 ${
            active ? "text-white/70" : "text-white/30"
          }`}
        >
          {sublabel}
        </p>
      </div>
    </div>
  );
}

/** ファイルドロップゾーン */
function FileDropZone({
  label,
  fileState,
  onFile,
  onSheetChange,
  accent,
}: {
  label: string;
  fileState: FileState;
  onFile: (file: File) => void;
  onSheetChange: (idx: number) => void;
  accent: string;
}) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const f = e.dataTransfer.files[0];
      if (f) onFile(f);
    },
    [onFile]
  );

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0];
      if (f) onFile(f);
    },
    [onFile]
  );

  return (
    <div className="flex-1 min-w-0">
      <div className="flex items-center gap-2 mb-2">
        <div
          className={`w-3 h-3 rounded-full ${accent}`}
        />
        <span className="text-sm font-semibold text-foreground/80">{label}</span>
      </div>

      {!fileState.file ? (
        <div
          className={`drop-zone rounded-lg p-6 text-center cursor-pointer ${
            dragging ? "drag-over" : ""
          }`}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
          onClick={() => inputRef.current?.click()}
        >
          <Upload className="w-8 h-8 mx-auto mb-3 text-primary/50" />
          <p className="text-sm font-medium text-foreground/60">
            クリックまたはドラッグ＆ドロップ
          </p>
          <p className="text-xs text-foreground/40 mt-1">.xlsx / .xls / .csv</p>
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={handleChange}
          />
        </div>
      ) : (
        <div className="border border-border rounded-lg p-4 bg-card">
          {fileState.loading ? (
            <div className="flex items-center gap-3">
              <div className="w-5 h-5 border-2 border-primary border-t-transparent rounded-full animate-spin" />
              <span className="text-sm text-muted-foreground">読み込み中...</span>
            </div>
          ) : fileState.error ? (
            <div className="flex items-start gap-2 text-destructive">
              <AlertTriangle className="w-4 h-4 mt-0.5 shrink-0" />
              <span className="text-sm">{fileState.error}</span>
            </div>
          ) : (
            <>
              <div className="flex items-start gap-3">
                <FileSpreadsheet className="w-8 h-8 text-primary shrink-0 mt-0.5" />
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-semibold truncate">{fileState.file.name}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">
                    {fileState.data?.rows.length ?? 0} 行 ×{" "}
                    {fileState.data?.headers.length ?? 0} 列
                  </p>
                </div>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    onFile(null as unknown as File);
                  }}
                  className="text-muted-foreground hover:text-foreground transition-colors"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              {/* シート選択 */}
              {(fileState.data?.sheetNames.length ?? 0) > 1 && (
                <div className="mt-3 pt-3 border-t border-border">
                  <label className="text-xs font-medium text-muted-foreground block mb-1">
                    シート選択
                  </label>
                  <div className="relative">
                    <select
                      value={fileState.selectedSheet}
                      onChange={(e) => onSheetChange(Number(e.target.value))}
                      className="w-full text-sm border border-input rounded-md px-3 py-1.5 bg-background appearance-none pr-8"
                    >
                      {fileState.data?.sheetNames.map((name, i) => (
                        <option key={i} value={i}>
                          {name}
                        </option>
                      ))}
                    </select>
                    <ChevronDown className="w-4 h-4 absolute right-2 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none" />
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
}

/** キー列選択コンポーネント */
function KeyColumnSelector({
  headers,
  selected,
  onChange,
}: {
  headers: string[];
  selected: string[];
  onChange: (cols: string[]) => void;
}) {
  const toggle = (col: string) => {
    if (selected.includes(col)) {
      onChange(selected.filter((c) => c !== col));
    } else {
      onChange([...selected, col]);
    }
  };

  return (
    <div className="flex flex-wrap gap-2">
      {headers.map((h) => {
        const active = selected.includes(h);
        return (
          <button
            key={h}
            onClick={() => toggle(h)}
            className={`px-3 py-1.5 rounded-md text-sm font-medium border transition-all duration-150 ${
              active
                ? "bg-primary text-primary-foreground border-primary shadow-sm"
                : "bg-card text-foreground/70 border-border hover:border-primary/50 hover:text-foreground"
            }`}
          >
            {active && <span className="mr-1.5">✓</span>}
            {h}
          </button>
        );
      })}
    </div>
  );
}

/** 差分結果テーブル */
function DiffTable({
  result,
  keyColumns,
  showOnlyDiffs,
}: {
  result: DiffResult;
  keyColumns: string[];
  showOnlyDiffs: boolean;
}) {
  const displayRows = showOnlyDiffs
    ? result.rows.filter((r) => r.status !== "same")
    : result.rows;

  if (displayRows.length === 0) {
    return (
      <div className="text-center py-16 text-muted-foreground">
        <CheckCircle2 className="w-12 h-12 mx-auto mb-3 text-emerald-500" />
        <p className="font-semibold">差分はありません</p>
        <p className="text-sm mt-1">2つのファイルは完全に一致しています</p>
      </div>
    );
  }

  return (
    <div className="overflow-auto">
      <table className="data-table w-full border-collapse text-sm">
        <thead>
          <tr>
            <th className="sticky left-0 z-10 bg-primary text-primary-foreground px-3 py-2 text-left text-xs font-semibold whitespace-nowrap border-r border-primary/30">
              状態
            </th>
            {result.headers.map((h) => (
              <th
                key={h}
                className={`px-3 py-2 text-left text-xs font-semibold whitespace-nowrap border-r border-border/50 last:border-r-0 ${
                  keyColumns.includes(h)
                    ? "bg-primary text-primary-foreground"
                    : "bg-primary/80 text-primary-foreground"
                }`}
              >
                {h}
                {keyColumns.includes(h) && (
                  <span className="ml-1 text-[10px] opacity-70">[KEY]</span>
                )}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {displayRows.map((row, ri) => {
            const diffCols = new Set(row.diffs.map((d) => d.col));

            if (row.status === "added") {
              return (
                <tr key={ri} className="row-added border-b border-border/30">
                  <td className="sticky left-0 z-10 px-3 py-2 whitespace-nowrap bg-emerald-100 border-r border-border/30">
                    <Badge className="bg-emerald-500 text-white text-[10px] px-1.5 py-0">追加</Badge>
                  </td>
                  {result.headers.map((h) => (
                    <td key={h} className="px-3 py-2 whitespace-nowrap border-r border-border/30 last:border-r-0">
                      {row.rowB?.[h] ?? ""}
                    </td>
                  ))}
                </tr>
              );
            }

            if (row.status === "deleted") {
              return (
                <tr key={ri} className="row-deleted border-b border-border/30">
                  <td className="sticky left-0 z-10 px-3 py-2 whitespace-nowrap bg-rose-100 border-r border-border/30">
                    <Badge className="bg-rose-500 text-white text-[10px] px-1.5 py-0">削除</Badge>
                  </td>
                  {result.headers.map((h) => (
                    <td key={h} className="px-3 py-2 whitespace-nowrap border-r border-border/30 last:border-r-0">
                      {row.rowA?.[h] ?? ""}
                    </td>
                  ))}
                </tr>
              );
            }

            // modified or same
            const hasDiff = row.status === "modified";
            return (
              <tr
                key={ri}
                className={`border-b border-border/30 transition-colors ${
                  hasDiff ? "hover:bg-pink-50/50" : "hover:bg-muted/30"
                }`}
              >
                <td className="sticky left-0 z-10 px-3 py-2 whitespace-nowrap bg-background border-r border-border/30">
                  {hasDiff ? (
                    <Badge className="bg-pink-500 text-white text-[10px] px-1.5 py-0">変更</Badge>
                  ) : (
                    <Badge variant="secondary" className="text-[10px] px-1.5 py-0">同一</Badge>
                  )}
                </td>
                {result.headers.map((h) => {
                  const isDiff = diffCols.has(h);
                  const valA = row.rowA?.[h] ?? "";
                  const valB = row.rowB?.[h] ?? "";
                  return (
                    <td
                      key={h}
                      className={`px-3 py-2 whitespace-nowrap border-r border-border/30 last:border-r-0 ${
                        isDiff ? "diff-cell" : ""
                      }`}
                      title={isDiff ? `A: ${valA} → B: ${valB}` : undefined}
                    >
                      {isDiff ? (
                        <span className="flex flex-col gap-0.5">
                          <span className="line-through opacity-50 text-xs">{valA || "(空)"}</span>
                          <span>{valB || "(空)"}</span>
                        </span>
                      ) : (
                        valA
                      )}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

// ============================
// Main Page
// ============================
export default function Home() {
  const [step, setStep] = useState<Step>(1);
  const [fileA, setFileA] = useState<FileState>(initFileState());
  const [fileB, setFileB] = useState<FileState>(initFileState());
  const [keyColumns, setKeyColumns] = useState<string[]>([]);
  const [diffResult, setDiffResult] = useState<DiffResult | null>(null);
  const [showOnlyDiffs, setShowOnlyDiffs] = useState(false);
  const [comparing, setComparing] = useState(false);

  // ファイル読み込み処理
  const loadFile = useCallback(
    async (
      file: File | null,
      setter: React.Dispatch<React.SetStateAction<FileState>>,
      sheetIndex = 0
    ) => {
      if (!file) {
        setter(initFileState());
        return;
      }

      // 拡張子チェック
      const ext = file.name.split(".").pop()?.toLowerCase();
      if (!["xlsx", "xls", "csv"].includes(ext ?? "")) {
        setter((s) => ({ ...s, file, error: ".xlsx / .xls / .csv ファイルのみ対応しています" }));
        return;
      }

      setter((s) => ({ ...s, file, loading: true, error: null, data: null }));

      try {
        const buffer = await file.arrayBuffer();
        const data = parseExcel(buffer, sheetIndex);
        if (data.headers.length === 0) {
          setter((s) => ({
            ...s,
            loading: false,
            error: "ヘッダー行が見つかりませんでした。ファイルを確認してください。",
          }));
          return;
        }
        setter((s) => ({ ...s, loading: false, data, selectedSheet: sheetIndex }));
      } catch (e) {
        setter((s) => ({
          ...s,
          loading: false,
          error: "ファイルの読み込みに失敗しました。",
        }));
      }
    },
    []
  );

  const handleFileA = useCallback(
    (file: File) => loadFile(file, setFileA, 0),
    [loadFile]
  );
  const handleFileB = useCallback(
    (file: File) => loadFile(file, setFileB, 0),
    [loadFile]
  );

  const handleSheetChangeA = useCallback(
    (idx: number) => {
      if (fileA.file) loadFile(fileA.file, setFileA, idx);
    },
    [fileA.file, loadFile]
  );
  const handleSheetChangeB = useCallback(
    (idx: number) => {
      if (fileB.file) loadFile(fileB.file, setFileB, idx);
    },
    [fileB.file, loadFile]
  );

  // ステップ1 → 2
  const goToStep2 = () => {
    if (!fileA.data || !fileB.data) {
      toast.error("両方のファイルをアップロードしてください");
      return;
    }
    // デフォルトキー列：両方に共通する最初の列
    const commonHeaders = fileA.data.headers.filter((h) =>
      fileB.data!.headers.includes(h)
    );
    setKeyColumns(commonHeaders.length > 0 ? [commonHeaders[0]] : []);
    setStep(2);
  };

  // ステップ2 → 3（比較実行）
  const runComparison = async () => {
    if (keyColumns.length === 0) {
      toast.error("キーとなる列を1つ以上選択してください");
      return;
    }
    setComparing(true);
    // 少し遅延を入れてUIの更新を確実にする
    await new Promise((r) => setTimeout(r, 50));
    try {
      const result = compareSheets(fileA.data!, fileB.data!, keyColumns);
      setDiffResult(result);
      setShowOnlyDiffs(result.stats.modified + result.stats.added + result.stats.deleted > 0);
      setStep(3);
      if (result.stats.modified + result.stats.added + result.stats.deleted === 0) {
        toast.success("差分はありません。2つのファイルは完全に一致しています。");
      } else {
        toast.info(
          `差分を検出しました: 変更 ${result.stats.modified}件 / 追加 ${result.stats.added}件 / 削除 ${result.stats.deleted}件`
        );
      }
    } catch (e) {
      toast.error("比較中にエラーが発生しました");
    } finally {
      setComparing(false);
    }
  };

  // リセット
  const reset = () => {
    setStep(1);
    setFileA(initFileState());
    setFileB(initFileState());
    setKeyColumns([]);
    setDiffResult(null);
    setShowOnlyDiffs(false);
  };

  // Excelエクスポート
  const handleExport = () => {
    if (!diffResult) return;
    try {
      exportDiffToExcel(
        diffResult,
        keyColumns,
        fileA.file?.name ?? "ファイルA",
        fileB.file?.name ?? "ファイルB"
      );
      toast.success("差分結果をExcelファイルとしてダウンロードしました");
    } catch (e) {
      toast.error("エクスポートに失敗しました");
    }
  };

  // 共通ヘッダー（キー列選択用）
  const commonHeaders =
    fileA.data && fileB.data
      ? fileA.data.headers.filter((h) => fileB.data!.headers.includes(h))
      : [];

  return (
    <div className="min-h-screen flex" style={{ fontFamily: "'DM Sans', 'Noto Sans JP', sans-serif" }}>
      {/* ===== LEFT SIDEBAR ===== */}
      <aside
        className="w-64 shrink-0 flex flex-col"
        style={{ background: "linear-gradient(160deg, #2D3180 0%, #1a1d4a 100%)" }}
      >
        {/* Logo area */}
        <div className="px-6 pt-8 pb-6 border-b border-white/10">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-white/15 flex items-center justify-center">
              <FileSpreadsheet className="w-4 h-4 text-white" />
            </div>
            <div>
              <p className="text-white font-bold text-sm leading-tight">Excel</p>
              <p className="text-white/60 text-xs">差分チェッカー</p>
            </div>
          </div>
        </div>

        {/* Steps */}
        <nav className="flex-1 px-6 py-8 space-y-1">
          <StepIndicator
            step={1}
            current={step}
            label="ファイルアップロード"
            sublabel="2つのExcelを選択"
          />
          <StepIndicator
            step={2}
            current={step}
            label="キー列の指定"
            sublabel="比較の基準となる列"
          />
          <StepIndicator
            step={3}
            current={step}
            label="差分結果"
            sublabel="相違箇所をハイライト"
          />
        </nav>

        {/* Info */}
        <div className="px-6 pb-8">
          <div className="rounded-lg bg-white/5 border border-white/10 p-3">
            <div className="flex items-start gap-2">
              <Info className="w-3.5 h-3.5 text-white/50 shrink-0 mt-0.5" />
              <p className="text-white/40 text-[11px] leading-relaxed">
                ファイルはブラウザ内で処理されます。サーバーへの送信はありません。
              </p>
            </div>
          </div>
        </div>
      </aside>

      {/* ===== MAIN CONTENT ===== */}
      <main className="flex-1 min-w-0 overflow-auto bg-background">
        <div className="max-w-4xl mx-auto px-8 py-10">

          {/* ===== STEP 1: FILE UPLOAD ===== */}
          {step === 1 && (
            <div
              className="animate-in fade-in slide-in-from-right-4 duration-300"
            >
              <div className="mb-8">
                <div className="flex items-center gap-2 mb-1">
                  <span className="text-xs font-semibold text-primary/60 uppercase tracking-widest">
                    Step 01
                  </span>
                </div>
                <h1 className="text-2xl font-bold text-foreground">
                  ファイルをアップロード
                </h1>
                <p className="text-muted-foreground mt-1 text-sm">
                  比較したい2つのExcelファイル（.xlsx / .xls / .csv）を選択してください
                </p>
              </div>

              <div className="flex gap-6 mb-8">
                <FileDropZone
                  label="ファイル A（比較元）"
                  fileState={fileA}
                  onFile={handleFileA}
                  onSheetChange={handleSheetChangeA}
                  accent="bg-primary"
                />
                <FileDropZone
                  label="ファイル B（比較先）"
                  fileState={fileB}
                  onFile={handleFileB}
                  onSheetChange={handleSheetChangeB}
                  accent="bg-pink-500"
                />
              </div>

              {/* ヘッダープレビュー */}
              {(fileA.data || fileB.data) && (
                <div className="mb-8 p-4 rounded-lg border border-border bg-card">
                  <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider mb-3">
                    ヘッダープレビュー
                  </p>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-xs font-medium text-foreground/60 mb-2">ファイル A</p>
                      <div className="flex flex-wrap gap-1.5">
                        {fileA.data?.headers.map((h) => (
                          <span key={h} className="px-2 py-0.5 rounded bg-primary/10 text-primary text-xs font-mono">
                            {h}
                          </span>
                        )) ?? <span className="text-xs text-muted-foreground">—</span>}
                      </div>
                    </div>
                    <div>
                      <p className="text-xs font-medium text-foreground/60 mb-2">ファイル B</p>
                      <div className="flex flex-wrap gap-1.5">
                        {fileB.data?.headers.map((h) => (
                          <span key={h} className="px-2 py-0.5 rounded bg-pink-500/10 text-pink-700 text-xs font-mono">
                            {h}
                          </span>
                        )) ?? <span className="text-xs text-muted-foreground">—</span>}
                      </div>
                    </div>
                  </div>
                </div>
              )}

              <div className="flex justify-end">
                <Button
                  onClick={goToStep2}
                  disabled={!fileA.data || !fileB.data}
                  className="gap-2 px-6"
                  size="lg"
                >
                  次へ：キー列を指定
                  <ArrowRight className="w-4 h-4" />
                </Button>
              </div>
            </div>
          )}

          {/* ===== STEP 2: KEY COLUMN SELECTION ===== */}
          {step === 2 && (
            <div className="animate-in fade-in slide-in-from-right-4 duration-300">
              <div className="mb-8">
                <div className="flex items-center gap-2 mb-1">
                  <span className="text-xs font-semibold text-primary/60 uppercase tracking-widest">
                    Step 02
                  </span>
                </div>
                <h1 className="text-2xl font-bold text-foreground">
                  キー列を指定
                </h1>
                <p className="text-muted-foreground mt-1 text-sm">
                  2つのファイルを紐付けるための「キー」となる列を選択してください。
                  複数選択すると複合キーになります。
                </p>
              </div>

              {/* ファイル情報サマリー */}
              <div className="grid grid-cols-2 gap-4 mb-8">
                <div className="p-4 rounded-lg border border-border bg-card">
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-2 h-2 rounded-full bg-primary" />
                    <span className="text-xs font-semibold text-muted-foreground">ファイル A</span>
                  </div>
                  <p className="text-sm font-medium truncate">{fileA.file?.name}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">
                    {fileA.data?.rows.length} 行 × {fileA.data?.headers.length} 列
                  </p>
                </div>
                <div className="p-4 rounded-lg border border-border bg-card">
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-2 h-2 rounded-full bg-pink-500" />
                    <span className="text-xs font-semibold text-muted-foreground">ファイル B</span>
                  </div>
                  <p className="text-sm font-medium truncate">{fileB.file?.name}</p>
                  <p className="text-xs text-muted-foreground mt-0.5">
                    {fileB.data?.rows.length} 行 × {fileB.data?.headers.length} 列
                  </p>
                </div>
              </div>

              {/* キー列選択 */}
              <div className="p-6 rounded-lg border border-border bg-card mb-4">
                <div className="flex items-center justify-between mb-4">
                  <p className="text-sm font-semibold">共通列（クリックで選択）</p>
                  {keyColumns.length > 0 && (
                    <div className="flex items-center gap-1.5">
                      <span className="text-xs text-muted-foreground">選択中:</span>
                      {keyColumns.map((k) => (
                        <Badge key={k} className="text-xs bg-primary text-primary-foreground">
                          {k}
                        </Badge>
                      ))}
                    </div>
                  )}
                </div>
                {commonHeaders.length > 0 ? (
                  <KeyColumnSelector
                    headers={commonHeaders}
                    selected={keyColumns}
                    onChange={setKeyColumns}
                  />
                ) : (
                  <div className="flex items-center gap-2 text-amber-600">
                    <AlertTriangle className="w-4 h-4" />
                    <span className="text-sm">
                      2つのファイルに共通するヘッダーがありません。
                      ファイルの内容を確認してください。
                    </span>
                  </div>
                )}
              </div>

              {/* 片方にしかない列の警告 */}
              {fileA.data && fileB.data && (
                (() => {
                  const onlyA = fileA.data!.headers.filter(
                    (h) => !fileB.data!.headers.includes(h)
                  );
                  const onlyB = fileB.data!.headers.filter(
                    (h) => !fileA.data!.headers.includes(h)
                  );
                  if (onlyA.length === 0 && onlyB.length === 0) return null;
                  return (
                    <div className="p-4 rounded-lg border border-amber-200 bg-amber-50 mb-6">
                      <div className="flex items-start gap-2">
                        <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                        <div className="text-sm text-amber-800">
                          {onlyA.length > 0 && (
                            <p>
                              <span className="font-semibold">ファイルAのみ:</span>{" "}
                              {onlyA.join(", ")}
                            </p>
                          )}
                          {onlyB.length > 0 && (
                            <p>
                              <span className="font-semibold">ファイルBのみ:</span>{" "}
                              {onlyB.join(", ")}
                            </p>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })()
              )}

              <div className="flex items-center justify-between">
                <Button
                  variant="outline"
                  onClick={() => setStep(1)}
                  className="gap-2"
                >
                  戻る
                </Button>
                <Button
                  onClick={runComparison}
                  disabled={keyColumns.length === 0 || comparing}
                  className="gap-2 px-6"
                  size="lg"
                >
                  {comparing ? (
                    <>
                      <div className="w-4 h-4 border-2 border-primary-foreground border-t-transparent rounded-full animate-spin" />
                      比較中...
                    </>
                  ) : (
                    <>
                      差分を比較する
                      <ArrowRight className="w-4 h-4" />
                    </>
                  )}
                </Button>
              </div>
            </div>
          )}

          {/* ===== STEP 3: DIFF RESULT ===== */}
          {step === 3 && diffResult && (
            <div className="animate-in fade-in slide-in-from-right-4 duration-300">
              <div className="mb-6">
                <div className="flex items-center gap-2 mb-1">
                  <span className="text-xs font-semibold text-primary/60 uppercase tracking-widest">
                    Step 03
                  </span>
                </div>
                <h1 className="text-2xl font-bold text-foreground">差分結果</h1>
                <p className="text-muted-foreground mt-1 text-sm">
                  ピンク背景のセルが変更箇所です。変更前の値は取り消し線で表示されます。
                </p>
              </div>

              {/* 統計バッジ */}
              <div className="grid grid-cols-4 gap-3 mb-6">
                {[
                  {
                    label: "変更",
                    value: diffResult.stats.modified,
                    color: "bg-pink-100 text-pink-800 border-pink-200",
                    dot: "bg-pink-500",
                  },
                  {
                    label: "追加",
                    value: diffResult.stats.added,
                    color: "bg-emerald-100 text-emerald-800 border-emerald-200",
                    dot: "bg-emerald-500",
                  },
                  {
                    label: "削除",
                    value: diffResult.stats.deleted,
                    color: "bg-rose-100 text-rose-800 border-rose-200",
                    dot: "bg-rose-500",
                  },
                  {
                    label: "同一",
                    value: diffResult.stats.same,
                    color: "bg-muted text-muted-foreground border-border",
                    dot: "bg-muted-foreground",
                  },
                ].map((s) => (
                  <div
                    key={s.label}
                    className={`rounded-lg border p-3 ${s.color}`}
                  >
                    <div className="flex items-center gap-1.5 mb-1">
                      <div className={`w-2 h-2 rounded-full ${s.dot}`} />
                      <span className="text-xs font-medium">{s.label}</span>
                    </div>
                    <p className="text-2xl font-bold">{s.value}</p>
                  </div>
                ))}
              </div>

              {/* コントロールバー */}
              <div className="flex items-center justify-between mb-4 gap-4">
                <div className="flex items-center gap-3">
                  <label className="flex items-center gap-2 cursor-pointer select-none">
                    <div
                      onClick={() => setShowOnlyDiffs(!showOnlyDiffs)}
                      className={`w-9 h-5 rounded-full transition-colors duration-200 relative ${
                        showOnlyDiffs ? "bg-primary" : "bg-muted"
                      }`}
                    >
                      <div
                        className={`absolute top-0.5 w-4 h-4 rounded-full bg-white shadow transition-transform duration-200 ${
                          showOnlyDiffs ? "translate-x-4" : "translate-x-0.5"
                        }`}
                      />
                    </div>
                    <span className="text-sm text-foreground/70">差分のみ表示</span>
                  </label>
                  <span className="text-xs text-muted-foreground">
                    {showOnlyDiffs
                      ? `${diffResult.rows.filter((r) => r.status !== "same").length} 行表示`
                      : `${diffResult.rows.length} 行表示`}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={handleExport}
                    className="gap-1.5"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Excelで保存
                  </Button>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={reset}
                    className="gap-1.5"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    最初からやり直す
                  </Button>
                </div>
              </div>

              {/* 凡例 */}
              <div className="flex items-center gap-4 mb-3 text-xs text-muted-foreground">
                <span className="font-medium">凡例:</span>
                <span className="flex items-center gap-1.5">
                  <span className="inline-block w-3 h-3 rounded-sm bg-[#FFD6E7]" />
                  変更セル（ピンク）
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="inline-block w-3 h-3 rounded-sm bg-emerald-100" />
                  追加行
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="inline-block w-3 h-3 rounded-sm bg-rose-100" />
                  削除行
                </span>
              </div>

              {/* テーブル */}
              <div className="rounded-lg border border-border overflow-hidden shadow-sm">
                <DiffTable
                  result={diffResult}
                  keyColumns={keyColumns}
                  showOnlyDiffs={showOnlyDiffs}
                />
              </div>

              {/* キー列情報 */}
              <div className="mt-4 flex items-center gap-2 text-xs text-muted-foreground">
                <Info className="w-3.5 h-3.5" />
                <span>
                  キー列:{" "}
                  {keyColumns.map((k) => (
                    <code key={k} className="mx-0.5 px-1.5 py-0.5 rounded bg-primary/10 text-primary font-mono">
                      {k}
                    </code>
                  ))}
                </span>
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}
