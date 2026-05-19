# Excel差分チェッカー デザインアイデア

## アプリの目的
2つのExcelファイルをアップロードし、キー列を指定して差分を比較し、相違箇所をピンク背景でハイライト表示する業務ツール

---

<response>
<probability>0.07</probability>
<text>

## Idea A: "Precision Lab" — 精密分析ツール美学

**Design Movement**: Swiss International Typographic Style × Industrial Precision

**Core Principles**:
1. 情報の階層を厳格なグリッドで表現する
2. 余白を「呼吸」として使い、密度の高い表データを際立たせる
3. アクションとフィードバックを明確に分離する
4. 機能美：装飾は排除し、構造そのものが美しさを生む

**Color Philosophy**:
- ベース：オフホワイト (#F8F7F4) — 清潔感と紙のような温かみ
- テキスト：チャコール (#1C1C1E) — 高コントラストで読みやすさ最優先
- アクセント：ディープインディゴ (#3B3F8C) — 信頼性と知性
- 差分ハイライト：ビビッドピンク (#FF4D8F) — 視覚的に即座に識別可能
- ボーダー：ライトグレー (#E0DDD8) — 構造を示すが主張しない

**Layout Paradigm**:
- 左サイドバー（固定）でステップナビゲーション
- メインエリアは3段階のステップ表示（アップロード → キー選択 → 結果）
- 結果テーブルは横スクロール対応の全幅表示

**Signature Elements**:
1. ステップインジケーター：細い横線で繋がれた番号バッジ
2. ファイルドロップゾーン：破線ボーダー + 微細なドットパターン背景
3. テーブルヘッダー：濃いインディゴ背景に白文字

**Interaction Philosophy**:
- ドラッグ&ドロップ時にゾーンが拡大してフィードバック
- キー列選択はチェックボックスリスト（複数選択可）
- 比較実行ボタンは大きく中央配置

**Animation**:
- ステップ遷移：200ms ease-out のスライドイン
- テーブル行の出現：30ms stagger でフェードイン
- ハイライトセル：比較完了時に一瞬フラッシュしてからピンクに定着

**Typography System**:
- 見出し：DM Sans Bold — 幾何学的で現代的
- 本文・テーブル：IBM Plex Mono — データ表示に最適なモノスペース
- ラベル：DM Sans Medium

</text>
</response>

<response>
<probability>0.06</probability>
<text>

## Idea B: "Clean Workspace" — ミニマル業務ツール

**Design Movement**: Bauhaus Functionalism × Contemporary SaaS

**Core Principles**:
1. 一画面一目的：各ステップは独立したカードとして表示
2. 色は意味を持つ：ピンク＝差分、グリーン＝一致、グレー＝未比較
3. タイポグラフィで情報密度を制御する
4. ゼロ学習コスト：初めて使う人でも迷わないUI

**Color Philosophy**:
- ベース：ピュアホワイト (#FFFFFF)
- サーフェス：ライトグレー (#F5F5F5)
- プライマリ：スレートブルー (#4A6FA5) — 落ち着いた信頼感
- 差分：ピンク (#FFB3C6 背景 / #C9184A テキスト)
- ボーダー：#E8E8E8

**Layout Paradigm**:
- シングルカラム中央配置、最大幅 900px
- ステップカード方式：完了ステップは折りたたまれる
- 結果は別セクションとして下部に展開

**Signature Elements**:
1. ステップカード：左ボーダーにカラーラインでステータス表示
2. ファイルアップロード：2カラム並列配置
3. 差分サマリーバッジ：「X行に差分あり」を目立つ数字で表示

**Interaction Philosophy**:
- アコーディオン式ステップ（完了すると自動的に次へ）
- キー列はドロップダウンセレクト
- 結果テーブルはフィルタリング可能（差分のみ表示切替）

**Animation**:
- カード展開：250ms cubic-bezier(0.23, 1, 0.32, 1)
- 差分セルのハイライト：500ms で順次ピンクに染まる
- 数値カウンター：0から差分数までアニメーション

**Typography System**:
- 見出し：Noto Sans JP Bold
- 本文：Noto Sans JP Regular
- テーブル：Noto Sans Mono

</text>
</response>

<response>
<probability>0.08</probability>
<text>

## Idea C: "Data Forge" — ダークモード精密ツール

**Design Movement**: Dark Mode Industrial × Data Visualization

**Core Principles**:
1. 暗い背景でデータを際立たせる
2. 蛍光アクセントで重要情報を強調
3. 密度の高いデータ表示を優先
4. プロフェッショナルな雰囲気

**Color Philosophy**:
- ベース：ダークネイビー (#0F1117)
- サーフェス：ダークグレー (#1A1D27)
- アクセント：エレクトリックブルー (#00D4FF)
- 差分：ネオンピンク (#FF2D78 背景) — 暗い背景で強烈なコントラスト
- テキスト：ライトグレー (#E2E8F0)

**Layout Paradigm**:
- 左サイドバー + メインコンテンツの2カラム
- ステップはサイドバーのタイムラインで管理
- 結果テーブルは全画面幅を活用

**Signature Elements**:
1. グロー効果：アクティブ要素に青いグロー
2. グリッドライン：背景に微細なグリッドパターン
3. ターミナル風の処理ログ表示

**Interaction Philosophy**:
- ホバー時にセルが明るくなる
- 差分行は左ボーダーにピンクのインジケーター
- キーボードショートカット対応

**Animation**:
- スキャン演出：比較実行時に行を上から順にスキャンするアニメーション
- 差分発見時：ピンクのパルスエフェクト

**Typography System**:
- 見出し：Space Grotesk Bold
- 本文：Inter
- テーブル：JetBrains Mono

</text>
</response>

---

## 選択: Idea A — "Precision Lab"

業務ツールとして最も使いやすく、データの視認性が高い。
スイスタイポグラフィの厳格さと、差分ハイライトのビビッドピンクのコントラストが最大の特徴。
