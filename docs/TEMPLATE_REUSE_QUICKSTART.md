# 4:3 発表資料テンプレートの再利用手順

> 目的: 新しい生成器を増やさず、既存のGrid Deckを使って定例発表の4:3資料を作る。研究固有の文章・測定値・写真は、このリポジトリに持ち込まない。

## 「テンプレート」を3つに分ける

| 種類 | 入力 | 変更できるもの | 注意点 |
| --- | --- | --- | --- |
| 内容・構成の雛形 | `presets/research-progress.json` | 目的／前回との差分／根拠／課題／次の作業という話の順序 | 話の順序を示すPreset。これを`design`に指定してもGrid Deckのスタイルにはならない |
| デザイン・配置のPreset | `glanceable-scientific` + Grid Deck JSON | 4:3、書体、色、余白、ブロック数、カード位置、画像枠、役割別配色 | 対応キーに限って上書き可能。任意のPowerPoint機能の再現を保証しない |
| PowerPointネイティブテンプレート | 手元の `.potx` | マスター・テーマなど | Windows版PowerPointの最終処理で適用する。Linuxの生成・CIでは見た目を検証できない |

既存の人間編集済みPPTXは別物。元ファイルを正本として保持し、`pptx_reverse.py` / `pptx_bundle.py` で内容・設計傾向を抽出する。逆変換しただけで完全なGrid DSLテンプレートになるとは限らない。

## 今すぐ試す: 既存エンジンを利用した4:3の例

`examples/research_update_4x3_customizable.json` は3ページの**架空の文例**。実験結果ではない。実際のデータに置き換えてから使用する。

```powershell
npm ci
python src/grid_deck.py build examples/research_update_4x3_customizable.json `
  --out-spec output/research_update_4x3.pixel.json `
  --out-ir output/research_update_4x3.ir.json `
  --pptx output/research_update_4x3.pptx `
  --strict
```

変更箇所は小さく保つ。

1. **文章**: 各スライドの`title`、`blocks.*.title`、`blocks.*.body`。
2. **配置**: `grid`を1列・2列・2×2等に変更し、同じIDの`blocks`を編集する。
3. **見た目**: `design.preset`を起点に`design.theme`（書体・色など）と`design.grid`（カード間隔など）を必要な項目だけ上書きする。
4. **寸法**: `canvas.preset`は研究発表時`4:3`を維持する。既存の画像の縦横比や実測図を勝手に伸縮・再解釈しない。

`research-progress.json`の構成方針は文章を作る際の指針として使い、レイアウトの実装は`src/grid_deck.py`に一本化する。

## ネイティブ `.potx` を使う場合（Windows + PowerPointが必要）

上のJSONに次をトップレベルで追加し、ローカルPC上の実際のファイルパスに置き換える。

```json
"template": {"path": "C:/path/to/existing-template.potx", "applyOnWindows": true}
```

Grid Deckの生成後、次を実行する。元の`output/research_update_4x3.pptx`は上書きしない。

```powershell
python -m pip install -r requirements-windows-powerpoint.txt
python src/powerpoint_finalize.py plan `
  output/research_update_4x3.ir.json `
  output/research_update_4x3.generation-report.json `
  --out output/research_update_4x3.finalize-plan.json
python src/powerpoint_finalize.py apply `
  output/research_update_4x3.pptx `
  output/research_update_4x3.finalize-plan.json `
  --out output/research_update_4x3.final.pptx
```

`ApplyTemplate`の呼び出しは実装されているが、ローカルの`.potx`で実際にデザインが維持されるかは別途レンダリングと目視確認が必要。

## 次回発表で「使える」と判定する条件

- 元の人間編集PPTXを変更・上書きしていない。
- 新規生成資料のページ比率が4:3で、文字切れ・図の縦横比崩れ・要素の衝突がない。
- 数値・写真・実験結果は研究リポジトリ側の出典に照らして確認済み。未確認値は未確認と明示する。
- Grid DSL → IR → lint → PPTX が成功し、Windowsでの表示・アニメーション・ネイティブテンプレートは実機で別途確認する。
- 人間の修正内容は次回の候補として記録し、**一度の修正で共通Presetを自動変更しない**。

## リポジトリ間の境界

- `graduation-research-motion-analysis`: 発表原稿、実験データ、写真、最終PPTXの正本。
- `presentation`: 汎用プリセット、テンプレートの適用方法、生成・検証コード、研究内容を含まない例。
- `cameradisp`: 計測・画像処理。PowerPoint生成器を重複実装しない。

既存の自動生成UIは未マージPR（#3）にある。内容の重複・CI・コンフリクトを確認するまで、これを標準の生成経路として扱わない。
