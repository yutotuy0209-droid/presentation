# 卒業研究 PowerPoint 自動生成・再構成プロンプト

このプロンプトは、以下の2つのGitHubリポジトリを連携して、卒業研究用PowerPointを再構成・生成するための作業指示です。

- プレゼン生成基盤: `yutotuy0209-droid/presentation`
- 研究内容の正本: `yutotuy0209-droid/graduation-research-motion-analysis`

---

## 目的

研究内容とPowerPoint生成基盤を分離し、以下を満たす1つの編集可能なPowerPointを作成してください。

1. 研究内容は `graduation-research-motion-analysis` の最新mainを正本とする。
2. スライド生成・逆解析・品質検査は `presentation` を使う。
3. 過去の人間作成PowerPoint、特に6月・7月の資料を教師データとして利用する。
4. スライド比率は今回 **4:3** を優先する。
5. 文字被り、画像との衝突、スライド外へのはみ出し、文字切れを自動検査する。
6. 編集可能なPPTXは原則1つだけを正本として残す。
7. PDF / PNG / HTML / IR / JSON は検証・閲覧用の派生物として扱う。
8. 研究固有の数値・実験結果・未確認事項を捏造しない。

---

## 最初に必ず確認すること

### presentation 側

最新mainを確認したうえで、最低限以下を読む。

- `README.md`
- `ROADMAP.md`
- `src/pptx_to_ir.py`
- `src/ir_to_pptx.js`
- `src/html_to_ir.js`
- `src/lint_ir.js`
- `src/compare_ir.py`
- `src/diff_ir.js`
- `src/image_diff.py`
- `src/render_pptx_powerpoint.ps1`
- `src/roundtrip_loop.py`
- `presets/research-progress.json`
- `presets/scientific.json`
- `skills/`

既存実装を無視して新しい独立系ツールを増やさないこと。
まず現在のIR構造とroundtripフローに統合できるか確認する。

### graduation-research-motion-analysis 側

最新mainを確認したうえで、最低限以下を読む。

- `README.md`
- `PROJECT_STARTUP_MEMO.md`
- `AGENTS.md`
- `src_python/guides/PRESENTATION_TEMPLATE.md`
- 最新打ち合わせPowerPoint
- 6月・7月の過去PowerPoint
- その週の正本資料
- `src_python/actual_code/`
- 関連する実験結果・CSV・画像
- 必要な場合のみ `CHANGELOG.md` の該当箇所

巨大ファイルを最初から全読み込みしない。
必要な部分だけ段階的に取得する。

---

## 情報の優先順位

必ず以下の順で判断する。

1. 最新GitHub main上の実コード・正本資料
2. その週の実験結果・保存データ
3. presentation側の生成・検査ルール
4. 過去PowerPoint
5. 過去チャットや記憶

過去資料と現在実装が食い違う場合は、最新mainを優先する。

「確認済み」「Mock確認済み」「未実機確認」「予定」「要確認」を混同しない。

---

## 研究背景・目的の基本方針

背景・目的は、研究リポジトリの現行内容を確認した上で整理する。

中心メッセージは次の方向性とする。

> 既存の工作機械に外付け可能な比較的簡単な画像計測系を用い、X・Y・Z方向の微小変位を同一のシステムで測定することを目指す。

ただし、

- 「安価」
- 「高精度」
- 「リアルタイム」
- 「従来法より優れる」

などの表現は、GitHub上に根拠がある範囲でのみ使用する。

価格、精度、性能比較を推測で書かない。

背景スライドは可能なら、

`工作機械の熱変位 → 加工精度への影響 → 既存計測の設置/運用上の課題 → 外付け画像計測 → XYZ計測`

の流れで見せる。

---

## 研究の基本構造

最新mainで正式名を再確認したうえで、概念上は以下の4段階を維持する。

- Step 0: スケール校正
- Step 1: 回転中心同定
- Step 2: XY面内変位
- Step 3: Z面外変位

Step 0〜3を毎回1枚ずつ説明する必要はない。

その週の主成果でなければ1枚の全体フローへ圧縮する。
その週に変更・検証したStepだけ個別展開する。

---

## 過去PowerPointの逆解析

特に6月・7月の人間作成PowerPointを優先的に探す。

見つけたら、単なる画像参照ではなく `src/pptx_to_ir.py` 等を使って逆解析する。

抽出対象:

- スライドサイズ
- 4:3 / 16:9
- タイトル位置
- 本文位置
- 余白
- 文字サイズ
- フォント
- 色
- 写真サイズ
- 図の占有率
- フッター位置
- ページ番号
- 要素数
- 1枚あたりの情報量
- テキストと図の比率
- 同じ種類のスライドで繰り返される配置

目的はデザインを丸ごと模写することではなく、

> 人間が実際に使っていた配置規則・余白・情報密度

を抽出すること。

---

## 4:3化

今回の生成では **4:3** を優先する。

目安:

- PowerPoint標準4:3: 10 × 7.5 inch

ただし、過去の採用テンプレートに実寸がある場合はそちらを優先する。

4:3へ変換する際に、16:9の要素座標を単純縮小しない。

再レイアウトする。

特に以下を再計算する。

- タイトル幅
- 本文幅
- 画像幅
- 図表サイズ
- 左右余白
- 縦方向間隔
- フッター
- ページ番号

---

## 生成方針

### 1枚1メッセージ

各スライドには主メッセージを1つだけ置く。

PowerPointを研究ノート代わりにしない。

詳細説明はGitHub・Word・発表者ノート側に残す。

### デザイン

原則:

- 白背景
- 大きなタイトル
- 十分な余白
- 図や写真を大きく
- 本文は少なめ
- 色は意味の分類にのみ使用
- 同じ意味は同じ色
- 不要なカードUIを増やさない
- グラデーション、装飾アイコン、ダッシュボード風カード乱立を避ける
- AI生成特有の均一な角丸カード配置を避ける

---

## PowerPoint品質ゲート

### 生成前

IR段階で最低限以下を検査する。

- text / text の重なり
- text / image の重なり
- text / table の重なり
- スライド外へのはみ出し
- 小さすぎる文字
- 過密なテキスト
- 画像が小さすぎないか
- タイトルと本文の余白

既存の `src/lint_ir.js` を拡張して対応できる場合は、必ずそこへ統合する。

別の重複lintシステムをむやみに作らない。

### 生成後

PPTX生成後は必ず以下を行う。

1. PowerPoint本体でレンダリング
2. 全スライドをPNG化
3. 目視または画像差分で確認
4. 文字切れ
5. 重なり
6. 画像欠落
7. アスペクト比崩れ
8. 小さすぎる文字
9. 数式崩れ
10. フッター/ページ番号崩れ

ERRORが残っているPowerPointを完成扱いにしない。

---

## Roundtrip

可能なら次のループを使う。

```text
過去PPTX
  ↓
pptx_to_ir.py
  ↓
IR
  ↓
lint_ir.js
  ↓
ir_to_pptx.js
  ↓
PowerPointレンダリング
  ↓
pptx_to_ir.py
  ↓
compare_ir.py / diff_ir.js
  ↓
image_diff.py
  ↓
修正
```

目標は、

- 変換前後でレイアウト意味が維持される
- 文字が被らない
- 人間作成PPTの余白感を維持する
- 編集可能性を失わない

こと。

---

## 正本PowerPointの運用

同じ用途のPPTXを大量に残さない。

悪い例:

- final.pptx
- final2.pptx
- final_修正版.pptx
- final_本当の最終版.pptx

良い運用:

- 編集可能なPPTX: 1つ
- 過去版: Git履歴
- PDF: 配布用
- PNG: レイアウト検査用
- HTML: スマホ閲覧用
- IR / JSON: 解析用

---

## 今回の作業で実施してほしいこと

1. `presentation` の現行構造を確認する。
2. `graduation-research-motion-analysis` の最新研究内容を確認する。
3. 6月・7月のPowerPointを探す。
4. 見つけたPPTXをIRへ逆解析する。
5. 人間作成PPTの配置・余白・文字サイズを抽出する。
6. 4:3用の研究進捗Presetを作る、または既存Presetを拡張する。
7. IR lintへ文字被り・はみ出し検査を統合する。
8. 背景・目的を現在研究内容に合わせて再構成する。
9. 最新研究内容を使ってPowerPointを1つ生成する。
10. PowerPoint本体でレンダリングする。
11. 全ページ画像を確認する。
12. ERRORがあれば自動または再レイアウトで修正する。
13. 再生成・再検査する。
14. 最終PPTXを1つだけ正本として残す。
15. 変更したコード、Preset、検査結果、生成物を報告する。

---

## 重要な禁止事項

- 過去チャットだけを根拠に研究内容を作らない。
- GitHubを確認せずに関数名・実験値を書くこと。
- 研究結果を推測で補うこと。
- 未確認値を確定値として書くこと。
- 16:9を単純縮小して4:3にすること。
- 文字が被ったまま「完成」とすること。
- スライドをスクリーンショット化してテンプレート代わりにすること。
- 研究用PPTXの大量な版違いを作ること。
- presentation側へ研究固有の実験データをコピーすること。

---

## 完了報告

最後に以下を簡潔に報告する。

- presentation 開始HEAD
- graduation-research-motion-analysis 開始HEAD
- 参照した過去PowerPoint
- 採用したスライド比率
- 生成したPPTX
- 生成枚数
- overlapエラー件数
- out-of-bounds件数
- warning件数
- PowerPointレンダリング結果
- 変更したファイル
- 研究固有データをpresentationへコピーしていないこと
- 未確認事項
- commit hash / PR

---

## 作業開始時の短い指示

以下だけを別AIへ渡しても開始できる。

> `yutotuy0209-droid/presentation` をPowerPoint生成基盤、`yutotuy0209-droid/graduation-research-motion-analysis` を研究内容の正本として、卒業研究用PowerPointを再構成してください。最初に両方の最新mainを確認し、presentation側のPPTX↔IR roundtrip、lint、render、image diff、presetを確認してください。研究側ではREADME、PROJECT_STARTUP_MEMO、AGENTS、PRESENTATION_TEMPLATE、6月・7月の過去PPT、最新打ち合わせ資料、actual_codeを必要範囲だけ確認してください。今回のスライドは4:3を優先し、6月・7月の人間作成PPTから余白・フォント・配置を逆解析して再利用してください。文字被り、画像との衝突、はみ出し、文字切れを自動検査し、ERRORが残るPPTXを完成扱いにしないでください。編集可能なPPTXは1つだけ正本として残し、PDF/PNG/HTML/IR/JSONは派生物としてください。研究固有の内容をpresentationへコピーせず、研究内容は常にgraduation-research-motion-analysisから取得してください。
