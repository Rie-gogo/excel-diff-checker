/**
 * Home.tsx — Excel差分チェッカー メインページ
 * Design: "Precision Lab" — Swiss Typography × Industrial Precision
 *
 * Step 1: ファイルアップロード（A・B）
 * Step 2: キー列マッピング（A列 ↔ B列）
 * Step 3: 比較列マッピング（A列 ↔ B列）
 * Step 4: 差分結果（ピンクハイライト）
 *
 * 正規化: 半角/全角同一視、スペース（半角・全角）無視
 */

import { useState, useCallback, useRef } from "react";
import {
  parseExcel,
  compareSheets,
  exportDiffToExcel,
  buildDefaultColumnMappings,
  buildDefaultKeyMappings,
  normalize,
} from "@/lib/excelDiff";
import type { SheetData, DiffResult, ColumnMapping, KeyMapping } from "@/lib/excelDiff";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { toast } from "sonner";
import {
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  ArrowRight,
  RotateCcw,
  Download,
  AlertTriangle,
  Info,
  X,
  ChevronDown,
  ArrowLeftRight,
  Plus,
  Trash2,
  Key,
  Columns,
} from "lucide-react";

// ============================
// Types
// ============================
type Step = 1 | 2 | 3 | 4;

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
        {step < 4 && (
          <div
            className={`w-0.5 h-8 mt-1 transition-all duration-500 ${
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
  accentClass,
}: {
  label: string;
  fileState: FileState;
  onFile: (file: File | null) => void;
  onSheetChange: (idx: number) => void;
  accentClass: string;
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

  return (
    <div className="flex-1 min-w-0">
      <div className="flex items-center gap-2 mb-2">
        <div className={`w-3 h-3 rounded-full ${accentClass}`} />
        <span className="text-sm font-semibold text-foreground/80">{label}</span>
      </div>

      {!fileState.file ? (
        <div
          className={`drop-zone rounded-lg p-6 text-center cursor-pointer ${dragging ? "drag-over" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
          onClick={() => inputRef.current?.click()}
        >
          <Upload className="w-8 h-8 mx-auto mb-3 text-primary/50" />
          <p className="text-sm font-medium text-foreground/60">クリックまたはドラッグ＆ドロップ</p>
          <p className="text-xs text-foreground/40 mt-1">.xlsx / .xls / .csv</p>
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); }}
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
                    {fileState.data?.rows.length ?? 0} 行 × {fileState.data?.headers.length ?? 0} 列
                  </p>
                </div>
                <button onClick={() => onFile(null)} className="text-muted-foreground hover:text-foreground transition-colors">
                  <X className="w-4 h-4" />
                </button>
              </div>
              {(fileState.data?.sheetNames.length ?? 0) > 1 && (
                <div className="mt-3 pt-3 border-t border-border">
                  <label className="text-xs font-medium text-muted-foreground block mb-1">シート選択</label>
                  <div className="relative">
                    <select
                      value={fileState.selectedSheet}
                      onChange={(e) => onSheetChange(Number(e.target.value))}
                      className="w-full text-sm border border-input rounded-md px-3 py-1.5 bg-background appearance-none pr-8"
                    >
                      {fileState.data?.sheetNames.map((name, i) => (
                        <option key={i} value={i}>{name}</option>
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

// ============================
// 列マッピング行コンポーネント
// ============================
function MappingRow({
  colA,
  colB,
  headersA,
  headersB,
  usedA,
  usedB,
  onChangeA,
  onChangeB,
  onRemove,
  isKey,
  canRemove,
}: {
  colA: string;
  colB: string;
  headersA: string[];
  headersB: string[];
  usedA: Set<string>;
  usedB: Set<string>;
  onChangeA: (v: string) => void;
  onChangeB: (v: string) => void;
  onRemove: () => void;
  isKey?: boolean;
  canRemove: boolean;
}) {
  const normA = normalize(colA);
  const normB = normalize(colB);
  const matched = normA !== "" && normB !== "" && normA === normB;
  const mismatch = normA !== "" && normB !== "" && normA !== normB;

  return (
    <div className={`flex items-center gap-2 p-2 rounded-lg border transition-colors ${
      isKey ? "border-primary/30 bg-primary/5" : "border-border bg-card"
    }`}>
      {/* A列選択 */}
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5 mb-1">
          <div className="w-2 h-2 rounded-full bg-primary" />
          <span className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wide">ファイルA</span>
        </div>
        <div className="relative">
          <select
            value={colA}
            onChange={(e) => onChangeA(e.target.value)}
            className="w-full text-xs border border-input rounded px-2 py-1.5 bg-background appearance-none pr-6 font-mono"
          >
            <option value="">— 列を選択 —</option>
            {headersA.map((h) => (
              <option key={h} value={h} disabled={usedA.has(h) && h !== colA}>
                {h}
              </option>
            ))}
          </select>
          <ChevronDown className="w-3 h-3 absolute right-1.5 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none" />
        </div>
      </div>

      {/* 中央アイコン */}
      <div className="flex flex-col items-center gap-0.5 shrink-0">
        <ArrowLeftRight className={`w-4 h-4 ${
          matched ? "text-emerald-500" : mismatch ? "text-amber-500" : "text-muted-foreground/40"
        }`} />
        {matched && <span className="text-[9px] text-emerald-500 font-medium">一致</span>}
        {mismatch && <span className="text-[9px] text-amber-500 font-medium">相違</span>}
      </div>

      {/* B列選択 */}
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5 mb-1">
          <div className="w-2 h-2 rounded-full bg-pink-500" />
          <span className="text-[10px] font-semibold text-muted-foreground uppercase tracking-wide">ファイルB</span>
        </div>
        <div className="relative">
          <select
            value={colB}
            onChange={(e) => onChangeB(e.target.value)}
            className="w-full text-xs border border-input rounded px-2 py-1.5 bg-background appearance-none pr-6 font-mono"
          >
            <option value="">— 列を選択 —</option>
            {headersB.map((h) => (
              <option key={h} value={h} disabled={usedB.has(h) && h !== colB}>
                {h}
              </option>
            ))}
          </select>
          <ChevronDown className="w-3 h-3 absolute right-1.5 top-1/2 -translate-y-1/2 text-muted-foreground pointer-events-none" />
        </div>
      </div>

      {/* 削除ボタン */}
      <button
        onClick={onRemove}
        disabled={!canRemove}
        className="shrink-0 p-1.5 rounded text-muted-foreground hover:text-destructive hover:bg-destructive/10 transition-colors disabled:opacity-20 disabled:cursor-not-allowed"
        title="この行を削除"
      >
        <Trash2 className="w-3.5 h-3.5" />
      </button>
    </div>
  );
}

// ============================
// 差分結果テーブル
// ============================
function DiffTable({
  result,
  showOnlyDiffs,
}: {
  result: DiffResult;
  showOnlyDiffs: boolean;
}) {
  const displayRows = showOnlyDiffs
    ? result.rows.filter((r) => r.status !== "same")
    : result.rows;

  const keyLabels = result.keyMappings.map((km) => km.colA);

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
                  keyLabels.includes(h)
                    ? "bg-primary text-primary-foreground"
                    : "bg-primary/80 text-primary-foreground"
                }`}
              >
                {h}
                {keyLabels.includes(h) && (
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
                  {result.headers.map((h) => {
                    const km = result.keyMappings.find((k) => k.colA === h);
                    const cm = result.columnMappings.find((c) => (c.label ?? c.colA) === h);
                    const val = km
                      ? (row.rowB?.[km.colB] ?? "")
                      : cm
                      ? (row.rowB?.[cm.colB] ?? "")
                      : "";
                    return (
                      <td key={h} className="px-3 py-2 whitespace-nowrap border-r border-border/30 last:border-r-0">
                        {val}
                      </td>
                    );
                  })}
                </tr>
              );
            }

            if (row.status === "deleted") {
              return (
                <tr key={ri} className="row-deleted border-b border-border/30">
                  <td className="sticky left-0 z-10 px-3 py-2 whitespace-nowrap bg-rose-100 border-r border-border/30">
                    <Badge className="bg-rose-500 text-white text-[10px] px-1.5 py-0">削除</Badge>
                  </td>
                  {result.headers.map((h) => {
                    const km = result.keyMappings.find((k) => k.colA === h);
                    const cm = result.columnMappings.find((c) => (c.label ?? c.colA) === h);
                    const val = km
                      ? (row.rowA?.[km.colA] ?? "")
                      : cm
                      ? (row.rowA?.[cm.colA] ?? "")
                      : "";
                    return (
                      <td key={h} className="px-3 py-2 whitespace-nowrap border-r border-border/30 last:border-r-0">
                        {val}
                      </td>
                    );
                  })}
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
                  const km = result.keyMappings.find((k) => k.colA === h);
                  const cm = result.columnMappings.find((c) => (c.label ?? c.colA) === h);

                  // キー列はAの値を表示
                  if (km) {
                    const val = row.rowA?.[km.colA] ?? "";
                    return (
                      <td key={h} className="px-3 py-2 whitespace-nowrap border-r border-border/30 last:border-r-0 font-medium">
                        {val}
                      </td>
                    );
                  }

                  const diffInfo = row.diffs.find((d) => d.col === h);
                  const valA = cm ? (row.rowA?.[cm.colA] ?? "") : "";
                  const valB = cm ? (row.rowB?.[cm.colB] ?? "") : "";

                  return (
                    <td
                      key={h}
                      className={`px-3 py-2 whitespace-nowrap border-r border-border/30 last:border-r-0 ${isDiff ? "diff-cell" : ""}`}
                      title={isDiff && diffInfo ? `A: ${diffInfo.valueA} → B: ${diffInfo.valueB}` : undefined}
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

  // キーマッピング
  const [keyMappings, setKeyMappings] = useState<KeyMapping[]>([{ colA: "", colB: "" }]);
  // 比較列マッピング
  const [colMappings, setColMappings] = useState<ColumnMapping[]>([{ colA: "", colB: "", label: "" }]);

  const [diffResult, setDiffResult] = useState<DiffResult | null>(null);
  const [showOnlyDiffs, setShowOnlyDiffs] = useState(false);
  const [comparing, setComparing] = useState(false);

  // ファイル読み込み
  const loadFile = useCallback(
    async (
      file: File | null,
      setter: React.Dispatch<React.SetStateAction<FileState>>,
      sheetIndex = 0
    ) => {
      if (!file) { setter(initFileState()); return; }
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
          setter((s) => ({ ...s, loading: false, error: "ヘッダー行が見つかりませんでした。" }));
          return;
        }
        setter((s) => ({ ...s, loading: false, data, selectedSheet: sheetIndex }));
      } catch {
        setter((s) => ({ ...s, loading: false, error: "ファイルの読み込みに失敗しました。" }));
      }
    },
    []
  );

  const handleFileA = useCallback((file: File | null) => loadFile(file, setFileA, 0), [loadFile]);
  const handleFileB = useCallback((file: File | null) => loadFile(file, setFileB, 0), [loadFile]);
  const handleSheetChangeA = useCallback((idx: number) => { if (fileA.file) loadFile(fileA.file, setFileA, idx); }, [fileA.file, loadFile]);
  const handleSheetChangeB = useCallback((idx: number) => { if (fileB.file) loadFile(fileB.file, setFileB, idx); }, [fileB.file, loadFile]);

  // Step1 → Step2
  const goToStep2 = () => {
    if (!fileA.data || !fileB.data) { toast.error("両方のファイルをアップロードしてください"); return; }
    // デフォルトキーマッピングを生成
    const defaultKeys = buildDefaultKeyMappings(fileA.data.headers, fileB.data.headers);
    setKeyMappings(defaultKeys.length > 0 ? defaultKeys : [{ colA: "", colB: "" }]);
    setStep(2);
  };

  // Step2 → Step3
  const goToStep3 = () => {
    const validKeys = keyMappings.filter((km) => km.colA && km.colB);
    if (validKeys.length === 0) { toast.error("キー列を1つ以上設定してください"); return; }
    // デフォルト比較列マッピングを生成（キー列を除く）
    const defaultCols = buildDefaultColumnMappings(fileA.data!.headers, fileB.data!.headers)
      .filter((cm) => !validKeys.some((km) => km.colA === cm.colA));
    setColMappings(defaultCols.length > 0 ? defaultCols : [{ colA: "", colB: "", label: "" }]);
    setStep(3);
  };

  // Step3 → Step4（比較実行）
  const runComparison = async () => {
    const validKeys = keyMappings.filter((km) => km.colA && km.colB);
    const validCols = colMappings.filter((cm) => cm.colA && cm.colB);
    if (validKeys.length === 0) { toast.error("キー列を1つ以上設定してください"); return; }
    if (validCols.length === 0) { toast.error("比較列を1つ以上設定してください"); return; }

    setComparing(true);
    await new Promise((r) => setTimeout(r, 50));
    try {
      // ラベルを設定
      const mappingsWithLabel = validCols.map((cm) => ({
        ...cm,
        label: cm.label || cm.colA,
      }));
      // キー列も全マッピングに含める
      const allMappings: ColumnMapping[] = [
        ...validKeys.map((km) => ({ colA: km.colA, colB: km.colB, label: km.colA })),
        ...mappingsWithLabel,
      ];
      const result = compareSheets(fileA.data!, fileB.data!, validKeys, allMappings);
      setDiffResult(result);
      setShowOnlyDiffs(result.stats.modified + result.stats.added + result.stats.deleted > 0);
      setStep(4);
      if (result.stats.modified + result.stats.added + result.stats.deleted === 0) {
        toast.success("差分はありません。2つのファイルは完全に一致しています。");
      } else {
        toast.info(`差分を検出: 変更 ${result.stats.modified}件 / 追加 ${result.stats.added}件 / 削除 ${result.stats.deleted}件`);
      }
    } catch {
      toast.error("比較中にエラーが発生しました");
    } finally {
      setComparing(false);
    }
  };

  const reset = () => {
    setStep(1); setFileA(initFileState()); setFileB(initFileState());
    setKeyMappings([{ colA: "", colB: "" }]);
    setColMappings([{ colA: "", colB: "", label: "" }]);
    setDiffResult(null); setShowOnlyDiffs(false);
  };

  const handleExport = () => {
    if (!diffResult) return;
    try {
      exportDiffToExcel(diffResult, fileA.file?.name ?? "ファイルA", fileB.file?.name ?? "ファイルB");
      toast.success("差分結果をExcelファイルとしてダウンロードしました");
    } catch { toast.error("エクスポートに失敗しました"); }
  };

  // 使用済み列セット
  const usedKeyA = new Set(keyMappings.map((km) => km.colA).filter(Boolean));
  const usedKeyB = new Set(keyMappings.map((km) => km.colB).filter(Boolean));
  const usedColA = new Set(colMappings.map((cm) => cm.colA).filter(Boolean));
  const usedColB = new Set(colMappings.map((cm) => cm.colB).filter(Boolean));

  const headersA = fileA.data?.headers ?? [];
  const headersB = fileB.data?.headers ?? [];

  return (
    <div className="min-h-screen flex" style={{ fontFamily: "'DM Sans', 'Noto Sans JP', sans-serif" }}>
      {/* ===== SIDEBAR ===== */}
      <aside className="w-64 shrink-0 flex flex-col" style={{ background: "linear-gradient(160deg, #2D3180 0%, #1a1d4a 100%)" }}>
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

        <nav className="flex-1 px-6 py-8 space-y-1">
          <StepIndicator step={1} current={step} label="ファイルアップロード" sublabel="2つのExcelを選択" />
          <StepIndicator step={2} current={step} label="キー列マッピング" sublabel="行を紐付ける基準列" />
          <StepIndicator step={3} current={step} label="比較列マッピング" sublabel="照合する列の対応付け" />
          <StepIndicator step={4} current={step} label="差分結果" sublabel="相違箇所をハイライト" />
        </nav>

        {/* 正規化の説明 */}
        <div className="px-6 pb-8">
          <div className="rounded-lg bg-white/5 border border-white/10 p-3 space-y-2">
            <p className="text-white/60 text-[11px] font-semibold uppercase tracking-wide">比較ルール</p>
            <div className="space-y-1">
              {[
                "半角・全角は同一視",
                "スペースは無視して比較",
                "ブラウザ内処理（送信なし）",
              ].map((t) => (
                <div key={t} className="flex items-center gap-1.5">
                  <CheckCircle2 className="w-3 h-3 text-emerald-400 shrink-0" />
                  <span className="text-white/40 text-[11px]">{t}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </aside>

      {/* ===== MAIN ===== */}
      <main className="flex-1 min-w-0 overflow-auto bg-background">
        <div className="max-w-4xl mx-auto px-8 py-10">

          {/* ===== STEP 1: UPLOAD ===== */}
          {step === 1 && (
            <div className="animate-in fade-in slide-in-from-right-4 duration-300">
              <div className="mb-8">
                <span className="text-xs font-semibold text-primary/60 uppercase tracking-widest">Step 01</span>
                <h1 className="text-2xl font-bold text-foreground mt-1">ファイルをアップロード</h1>
                <p className="text-muted-foreground mt-1 text-sm">比較したい2つのExcelファイル（.xlsx / .xls / .csv）を選択してください</p>
              </div>

              <div className="flex gap-6 mb-8">
                <FileDropZone label="ファイル A（比較元）" fileState={fileA} onFile={handleFileA} onSheetChange={handleSheetChangeA} accentClass="bg-primary" />
                <FileDropZone label="ファイル B（比較先）" fileState={fileB} onFile={handleFileB} onSheetChange={handleSheetChangeB} accentClass="bg-pink-500" />
              </div>

              {(fileA.data || fileB.data) && (
                <div className="mb-8 p-4 rounded-lg border border-border bg-card">
                  <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider mb-3">ヘッダープレビュー</p>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-xs font-medium text-foreground/60 mb-2">ファイル A</p>
                      <div className="flex flex-wrap gap-1.5">
                        {fileA.data?.headers.map((h) => (
                          <span key={h} className="px-2 py-0.5 rounded bg-primary/10 text-primary text-xs font-mono">{h}</span>
                        )) ?? <span className="text-xs text-muted-foreground">—</span>}
                      </div>
                    </div>
                    <div>
                      <p className="text-xs font-medium text-foreground/60 mb-2">ファイル B</p>
                      <div className="flex flex-wrap gap-1.5">
                        {fileB.data?.headers.map((h) => (
                          <span key={h} className="px-2 py-0.5 rounded bg-pink-500/10 text-pink-700 text-xs font-mono">{h}</span>
                        )) ?? <span className="text-xs text-muted-foreground">—</span>}
                      </div>
                    </div>
                  </div>
                </div>
              )}

              <div className="flex justify-end">
                <Button onClick={goToStep2} disabled={!fileA.data || !fileB.data} className="gap-2 px-6" size="lg">
                  次へ：キー列を設定 <ArrowRight className="w-4 h-4" />
                </Button>
              </div>
            </div>
          )}

          {/* ===== STEP 2: KEY MAPPING ===== */}
          {step === 2 && (
            <div className="animate-in fade-in slide-in-from-right-4 duration-300">
              <div className="mb-8">
                <span className="text-xs font-semibold text-primary/60 uppercase tracking-widest">Step 02</span>
                <h1 className="text-2xl font-bold text-foreground mt-1">キー列マッピング</h1>
                <p className="text-muted-foreground mt-1 text-sm">
                  2つのファイルの行を紐付けるための「キー」となる列を対応付けてください。
                  ファイルAとBで列名が異なる場合もマッピングできます。
                </p>
              </div>

              {/* ファイル情報 */}
              <div className="grid grid-cols-2 gap-4 mb-6">
                <div className="p-3 rounded-lg border border-border bg-card">
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-2 h-2 rounded-full bg-primary" />
                    <span className="text-xs font-semibold text-muted-foreground">ファイル A</span>
                  </div>
                  <p className="text-sm font-medium truncate">{fileA.file?.name}</p>
                  <p className="text-xs text-muted-foreground">{fileA.data?.rows.length} 行 × {fileA.data?.headers.length} 列</p>
                </div>
                <div className="p-3 rounded-lg border border-border bg-card">
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-2 h-2 rounded-full bg-pink-500" />
                    <span className="text-xs font-semibold text-muted-foreground">ファイル B</span>
                  </div>
                  <p className="text-sm font-medium truncate">{fileB.file?.name}</p>
                  <p className="text-xs text-muted-foreground">{fileB.data?.rows.length} 行 × {fileB.data?.headers.length} 列</p>
                </div>
              </div>

              {/* キーマッピング */}
              <div className="p-5 rounded-lg border border-primary/20 bg-primary/5 mb-6">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2">
                    <Key className="w-4 h-4 text-primary" />
                    <span className="text-sm font-semibold text-primary">キー列の対応付け</span>
                  </div>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={() => setKeyMappings([...keyMappings, { colA: "", colB: "" }])}
                    className="gap-1.5 text-xs h-7"
                  >
                    <Plus className="w-3 h-3" /> キー列を追加
                  </Button>
                </div>

                <div className="space-y-2">
                  {keyMappings.map((km, i) => (
                    <MappingRow
                      key={i}
                      colA={km.colA}
                      colB={km.colB}
                      headersA={headersA}
                      headersB={headersB}
                      usedA={new Set(keyMappings.filter((_, j) => j !== i).map((k) => k.colA).filter(Boolean))}
                      usedB={new Set(keyMappings.filter((_, j) => j !== i).map((k) => k.colB).filter(Boolean))}
                      onChangeA={(v) => setKeyMappings(keyMappings.map((k, j) => j === i ? { ...k, colA: v } : k))}
                      onChangeB={(v) => setKeyMappings(keyMappings.map((k, j) => j === i ? { ...k, colB: v } : k))}
                      onRemove={() => setKeyMappings(keyMappings.filter((_, j) => j !== i))}
                      isKey
                      canRemove={keyMappings.length > 1}
                    />
                  ))}
                </div>

                <div className="mt-3 flex items-start gap-2 text-xs text-primary/70">
                  <Info className="w-3.5 h-3.5 shrink-0 mt-0.5" />
                  <span>キー列の値を使って、AとBの行を紐付けます。複合キーにする場合は複数追加してください。</span>
                </div>
              </div>

              <div className="flex items-center justify-between">
                <Button variant="outline" onClick={() => setStep(1)} className="gap-2">戻る</Button>
                <Button
                  onClick={goToStep3}
                  disabled={keyMappings.filter((km) => km.colA && km.colB).length === 0}
                  className="gap-2 px-6" size="lg"
                >
                  次へ：比較列を設定 <ArrowRight className="w-4 h-4" />
                </Button>
              </div>
            </div>
          )}

          {/* ===== STEP 3: COLUMN MAPPING ===== */}
          {step === 3 && (
            <div className="animate-in fade-in slide-in-from-right-4 duration-300">
              <div className="mb-8">
                <span className="text-xs font-semibold text-primary/60 uppercase tracking-widest">Step 03</span>
                <h1 className="text-2xl font-bold text-foreground mt-1">比較列マッピング</h1>
                <p className="text-muted-foreground mt-1 text-sm">
                  どのA列とB列を照合するか対応付けてください。列名が異なっていても指定できます。
                  半角・全角は同一視し、スペースは無視して比較します。
                </p>
              </div>

              {/* 正規化の説明バナー */}
              <div className="flex items-start gap-3 p-4 rounded-lg border border-emerald-200 bg-emerald-50 mb-6">
                <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                <div className="text-sm text-emerald-800">
                  <span className="font-semibold">正規化比較が有効：</span>
                  「Ａ」と「A」、「１」と「1」、「　」と「 」はすべて同一とみなします。スペースは比較から除外されます。
                </div>
              </div>

              {/* 比較列マッピング */}
              <div className="p-5 rounded-lg border border-border bg-card mb-6">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2">
                    <Columns className="w-4 h-4 text-foreground/70" />
                    <span className="text-sm font-semibold">比較列の対応付け</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => {
                        // 未使用の列を自動で追加
                        const unusedA = headersA.find(
                          (h) => !usedColA.has(h) && !usedKeyA.has(h)
                        ) ?? "";
                        const unusedB = headersB.find(
                          (h) => !usedColB.has(h) && !usedKeyB.has(h)
                        ) ?? "";
                        setColMappings([...colMappings, { colA: unusedA, colB: unusedB, label: unusedA }]);
                      }}
                      className="gap-1.5 text-xs h-7"
                    >
                      <Plus className="w-3 h-3" /> 列を追加
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => {
                        // 全列を自動マッピング（キー列除く）
                        const validKeys = keyMappings.filter((km) => km.colA && km.colB);
                        const auto = buildDefaultColumnMappings(headersA, headersB)
                          .filter((cm) => !validKeys.some((km) => km.colA === cm.colA));
                        setColMappings(auto.length > 0 ? auto : [{ colA: "", colB: "", label: "" }]);
                        toast.success("列を自動マッピングしました");
                      }}
                      className="gap-1.5 text-xs h-7"
                    >
                      自動マッピング
                    </Button>
                  </div>
                </div>

                <div className="space-y-2">
                  {colMappings.map((cm, i) => (
                    <div key={i} className="flex items-start gap-2">
                      <div className="flex-1">
                        <MappingRow
                          colA={cm.colA}
                          colB={cm.colB}
                          headersA={headersA.filter((h) => !usedKeyA.has(h))}
                          headersB={headersB.filter((h) => !usedKeyB.has(h))}
                          usedA={new Set(colMappings.filter((_, j) => j !== i).map((c) => c.colA).filter(Boolean))}
                          usedB={new Set(colMappings.filter((_, j) => j !== i).map((c) => c.colB).filter(Boolean))}
                          onChangeA={(v) => setColMappings(colMappings.map((c, j) => j === i ? { ...c, colA: v, label: c.label || v } : c))}
                          onChangeB={(v) => setColMappings(colMappings.map((c, j) => j === i ? { ...c, colB: v } : c))}
                          onRemove={() => setColMappings(colMappings.filter((_, j) => j !== i))}
                          canRemove={colMappings.length > 1}
                        />
                      </div>
                      {/* 表示名入力 */}
                      <div className="w-28 shrink-0 mt-5">
                        <input
                          type="text"
                          value={cm.label ?? cm.colA}
                          onChange={(e) => setColMappings(colMappings.map((c, j) => j === i ? { ...c, label: e.target.value } : c))}
                          placeholder="表示名"
                          className="w-full text-xs border border-input rounded px-2 py-1.5 bg-background font-mono"
                          title="結果テーブルでの列名（省略するとA列名を使用）"
                        />
                        <p className="text-[9px] text-muted-foreground mt-0.5 text-center">表示名</p>
                      </div>
                    </div>
                  ))}
                </div>

                {colMappings.filter((cm) => cm.colA && cm.colB).length === 0 && (
                  <div className="mt-3 flex items-center gap-2 text-amber-600 text-xs">
                    <AlertTriangle className="w-3.5 h-3.5" />
                    <span>比較列を1つ以上設定してください</span>
                  </div>
                )}
              </div>

              <div className="flex items-center justify-between">
                <Button variant="outline" onClick={() => setStep(2)} className="gap-2">戻る</Button>
                <Button
                  onClick={runComparison}
                  disabled={colMappings.filter((cm) => cm.colA && cm.colB).length === 0 || comparing}
                  className="gap-2 px-6" size="lg"
                >
                  {comparing ? (
                    <><div className="w-4 h-4 border-2 border-primary-foreground border-t-transparent rounded-full animate-spin" />比較中...</>
                  ) : (
                    <>差分を比較する <ArrowRight className="w-4 h-4" /></>
                  )}
                </Button>
              </div>
            </div>
          )}

          {/* ===== STEP 4: RESULT ===== */}
          {step === 4 && diffResult && (
            <div className="animate-in fade-in slide-in-from-right-4 duration-300">
              <div className="mb-6">
                <span className="text-xs font-semibold text-primary/60 uppercase tracking-widest">Step 04</span>
                <h1 className="text-2xl font-bold text-foreground mt-1">差分結果</h1>
                <p className="text-muted-foreground mt-1 text-sm">
                  ピンク背景のセルが変更箇所です。変更前の値は取り消し線で表示されます。
                </p>
              </div>

              {/* 統計 */}
              <div className="grid grid-cols-4 gap-3 mb-6">
                {[
                  { label: "変更", value: diffResult.stats.modified, color: "bg-pink-100 text-pink-800 border-pink-200", dot: "bg-pink-500" },
                  { label: "追加", value: diffResult.stats.added, color: "bg-emerald-100 text-emerald-800 border-emerald-200", dot: "bg-emerald-500" },
                  { label: "削除", value: diffResult.stats.deleted, color: "bg-rose-100 text-rose-800 border-rose-200", dot: "bg-rose-500" },
                  { label: "同一", value: diffResult.stats.same, color: "bg-muted text-muted-foreground border-border", dot: "bg-muted-foreground" },
                ].map((s) => (
                  <div key={s.label} className={`rounded-lg border p-3 ${s.color}`}>
                    <div className="flex items-center gap-1.5 mb-1">
                      <div className={`w-2 h-2 rounded-full ${s.dot}`} />
                      <span className="text-xs font-medium">{s.label}</span>
                    </div>
                    <p className="text-2xl font-bold">{s.value}</p>
                  </div>
                ))}
              </div>

              {/* マッピング情報 */}
              <div className="p-3 rounded-lg border border-border bg-card mb-4 text-xs">
                <div className="flex flex-wrap gap-x-6 gap-y-1">
                  <span className="text-muted-foreground">
                    <span className="font-semibold text-foreground">キー列:</span>{" "}
                    {diffResult.keyMappings.map((km) => `${km.colA} ↔ ${km.colB}`).join("、")}
                  </span>
                  <span className="text-muted-foreground">
                    <span className="font-semibold text-foreground">比較列:</span>{" "}
                    {diffResult.columnMappings
                      .filter((cm) => !diffResult.keyMappings.some((km) => km.colA === cm.colA))
                      .map((cm) => `${cm.colA} ↔ ${cm.colB}`)
                      .join("、")}
                  </span>
                </div>
              </div>

              {/* コントロールバー */}
              <div className="flex items-center justify-between mb-4 gap-4">
                <div className="flex items-center gap-3">
                  <label className="flex items-center gap-2 cursor-pointer select-none">
                    <div
                      onClick={() => setShowOnlyDiffs(!showOnlyDiffs)}
                      className={`w-9 h-5 rounded-full transition-colors duration-200 relative ${showOnlyDiffs ? "bg-primary" : "bg-muted"}`}
                    >
                      <div className={`absolute top-0.5 w-4 h-4 rounded-full bg-white shadow transition-transform duration-200 ${showOnlyDiffs ? "translate-x-4" : "translate-x-0.5"}`} />
                    </div>
                    <span className="text-sm text-foreground/70">差分のみ表示</span>
                  </label>
                  <span className="text-xs text-muted-foreground">
                    {showOnlyDiffs
                      ? `${diffResult.rows.filter((r) => r.status !== "same").length} 行`
                      : `${diffResult.rows.length} 行`}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <Button variant="outline" size="sm" onClick={handleExport} className="gap-1.5">
                    <Download className="w-3.5 h-3.5" /> Excelで保存
                  </Button>
                  <Button variant="outline" size="sm" onClick={reset} className="gap-1.5">
                    <RotateCcw className="w-3.5 h-3.5" /> 最初からやり直す
                  </Button>
                </div>
              </div>

              {/* 凡例 */}
              <div className="flex items-center gap-4 mb-3 text-xs text-muted-foreground">
                <span className="font-medium">凡例:</span>
                <span className="flex items-center gap-1.5"><span className="inline-block w-3 h-3 rounded-sm bg-[#FFD6E7]" />変更セル（ピンク）</span>
                <span className="flex items-center gap-1.5"><span className="inline-block w-3 h-3 rounded-sm bg-emerald-100" />追加行</span>
                <span className="flex items-center gap-1.5"><span className="inline-block w-3 h-3 rounded-sm bg-rose-100" />削除行</span>
              </div>

              <div className="rounded-lg border border-border overflow-hidden shadow-sm">
                <DiffTable result={diffResult} showOnlyDiffs={showOnlyDiffs} />
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}
