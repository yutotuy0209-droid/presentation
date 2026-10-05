# presentation

個人利用を中心に、再利用可能なプレゼン資料生成コード・構成パターン・変換時の注意点を保存するリポジトリです。

卒業研究で得たノウハウのうち、研究固有のデータ・文章・実験結果を含まない部分だけを切り出しています。

## Current scope

- Pythonで任意の非暗号化PPTXを素材・XML・編集用JSONへ展開し、未対応要素を保持して復元する：[PPTX_SOURCE_BUNDLE.md](docs/PPTX_SOURCE_BUNDLE.md)

- HTML / CSS スライドのレンダリング
- Puppeteer による高解像度 PNG 出力
- PNG から PowerPoint (PPTX) への変換
- 科学・研究発表向けの構成Preset
- 研究進捗報告向けPreset
- 数式変換を壊しにくくするSkill
- 発表全体の構成を選ぶSkill
- レンダリング用スモークテスト

## Structure

```text
presentation/
├─ src/
│  └─ html_to_pptx.js
├─ presets/
│  ├─ scientific.json
│  ├─ research-progress.json
│  └─ story-driven.json
├─ skills/
│  ├─ scientific-presentation/
│  │  └─ SKILL.md
│  ├─ equation-safe-presentation/
│  │  └─ SKILL.md
│  └─ presentation-structure/
│     └─ SKILL.md
└─ tests/
   ├─ test_puppeteer.js
   └─ fixtures/
      └─ test_slide.html
```

## Setup

```bash
npm install
```

Chrome / Chromium が通常の場所にない場合は `CHROME_PATH` を設定してください。

## Usage

```bash
node src/html_to_pptx.js output/result.pptx slide1.html slide2.html
```

## Design principle

既存の Marp、PptxGenJS、PowerPoint などを置き換えるのではなく、それらを使うときの「自分の定石」を保存する層として使います。

- どの構成を選ぶか
- 何を図にするか
- 数式をどう保持するか
- どこを画像化し、どこを編集可能なまま残すか
- 変換で崩れたときにどう直すか

を、使い捨ての会話ではなく再利用可能なコード・Preset・Skillとして残します。

## Origin

汎用処理は `graduation-research-motion-analysis` 内で作成したプレゼン生成コードを基にしています。研究固有の内容はこのリポジトリにはコピーしません。


## Pixel Layout Engine

AIがPowerPointの座標・文字サイズ・余白を毎回考えなくてもよいように、`src/pixel_layout_engine.py` を用意しています。

### 最小入力

```json
{
  "canvas": {"preset": "4:3"},
  "slides": [
    {
      "title": "研究背景と目的",
      "bullets": [
        "熱変位が加工精度へ影響する",
        "外付け画像計測でX・Y・Z変位を測定する"
      ]
    }
  ]
}
```

### 生成

```bash
python src/pixel_layout_engine.py build input.json \
  --out-ir output/deck.json \
  --pptx output/deck.pptx \
  --strict
```

Windows + PowerPoint環境では最終レンダリングまで1コマンドで実行できます。

```powershell
python src/pixel_layout_engine.py build input.json `
  --out-ir output/deck.json `
  --pptx output/deck.pptx `
  --render-dir output/rendered `
  --strict
```

エンジンが自動で行うこと:

- 4:3 / 16:9のcanvas設定
- 1px単位の座標処理
- px → PowerPoint native EMU変換
- 本文の高さ推定
- 文字サイズの自動縮小
- 長文の自動改ページ
- foreground要素の衝突回避
- スライド外へのはみ出し防止
- 画像の縦横比維持
- 表の高さ推定
- ページ番号とフッター
- common IR lint
- PPTX生成
- 任意でPowerPoint本体によるPNGレンダリング

詳細は `docs/PIXEL_LAYOUT_ENGINE.md` を参照してください。


## 2D Grid Deck

PowerPointの配置を毎回AIに座標指定させず、**意味ブロック + 2次元Grid + Design Preset** から生成します。

```json
{
  "design": "glanceable-scientific",
  "slides": [{
    "title": "入力 → 処理 → 出力",
    "grid": [["input", "process", "output"]],
    "blocks": {
      "input": {"type": "card", "role": "input", "title": "Camera"},
      "process": {"type": "card", "role": "process", "title": "POC"},
      "output": {"type": "metric", "role": "output", "value": "10.1 µm", "label": "XY displacement"}
    },
    "connectors": [
      {"from": "input", "to": "process"},
      {"from": "process", "to": "output"}
    ]
  }]
}
```

生成:

```bash
npm run build:grid
```

対応:

- 1×1 / 左右 / 上下 / 2×2 / span
- card / metric / image / SVG / GIF / video
- optional connector
- design role: input / process / output / human / warning
- animated GIF
- video埋め込み
- native PowerPoint animation用の予約情報

設計詳細: [GRID_DECK_SYSTEM.md](docs/GRID_DECK_SYSTEM.md)

デザインPreset:

- `glanceable-scientific`
- `glanceable-soft-graphite`
- `glanceable-warm`

原則は **「一目で構造、数秒で関係、読めば具体」** です。

## Python + PowerPoint finalization

Grid/PPTX生成後、Windows版PowerPoint本体でしか扱いにくい機能は `src/powerpoint_finalize.py` で後処理できます。

現在の対象:

- Fade / Appear / Wipe / Fly entrance animation
- click / with_previous / after_previous trigger
- slide transition
- hyperlink
- GIF / SVG / video / audioを含む生成PPTXとの共存

生成時に `ir_to_pptx.js` が `shapeMap` をreportへ出すため、Pythonは形状を推測せず、生成したshape indexを使います。

Windows:

```powershell
python -m pip install -r requirements-windows-powerpoint.txt

python src/powerpoint_finalize.py plan `
  output/grid_deck.ir.json `
  output/grid_deck.generation-report.json `
  --out output/grid_deck.finalize-plan.json

python src/powerpoint_finalize.py apply `
  output/grid_deck.pptx `
  output/grid_deck.finalize-plan.json `
  --out output/grid_deck.final.pptx
```

### 動画・音声の考え方

人間が指定するのは基本的にファイルパスだけです。

```json
{
  "type": "video",
  "path": "assets/experiment.mp4"
}
```

または:

```json
{
  "type": "audio",
  "path": "assets/narration.mp3"
}
```

生成時にそのファイルを読み込み、PowerPoint mediaとして処理します。
「PPTX内に単なる文字列のパスだけを置く」設計にはしません。

## Design skill

Grid Deckを作るときは [Glanceable Grid Presentation Skill](skills/glanceable-grid-presentation/SKILL.md) を使用します。

設計原則:

> 一目で構造、数秒で関係、読めば具体。

生成前に `src/design_lint.py` が、長すぎるタイトル、Card本文過多、Block過多、Connector過多などを点検します。

## Background / template

Grid DSL can specify a per-slide background color or image:

```json
{
  "background": {
    "color": "F7F9FC",
    "image": "assets/background.png"
  }
}
```

Existing PowerPoint template files can be handed to the Windows finalizer:

```json
{
  "template": {
    "path": "templates/lab-template.potx",
    "applyOnWindows": true
  }
}
```

The base deck stays deterministic; Windows PowerPoint applies the native template in the finalization stage.

## PPTX reverse conversion

Existing PowerPoint files can be converted back to text/JSON:

```bash
python src/pptx_reverse.py existing.pptx --out-dir output/reversed
```

Outputs:

- `slide_ir.json` — text, geometry, tables and media references for supported objects
- `slides.md` — human-readable slide text and geometry
- `source_bundle/` — loss-aware package preserving master/theme/media/unknown OOXML when optional lxml support is available
- `reverse-report.json` — limitations and output paths

Two reverse modes are intentionally separate:

1. **Editable IR** for analysis/re-generation.
2. **Loss-aware source bundle** when preserving the original PowerPoint package matters more than semantic simplification.

An arbitrary old PPTX cannot be perfectly reconstructed as the high-level Grid DSL because PowerPoint does not store concepts such as `role: input` or `grid: [[A,B,C]]`. Those meanings can only be inferred, so the reverse tool does not silently invent them.

## Round-trip first

Reverse conversion is a core architecture requirement, not only an import utility.

Coverage is judged by whether a real human-created PPTX feature can be detected, represented or safely preserved, regenerated/retained, and verified after round-trip.

Human-edited decks are also design evidence: spacing, typography, geometry, palette, image/text ratio and recurring layout families should be extracted and compared rather than encoded as one-off prompt rules.

See [ROUNDTRIP_FIRST_ARCHITECTURE.md](docs/ROUNDTRIP_FIRST_ARCHITECTURE.md).
