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

## Auto Deck

Markdown / JSON / TXT から、compact spec → IR → lint → editable PPTX → 任意のPowerPoint実レンダリングまでを1コマンドで実行できます。

```powershell
python src/auto_deck.py examples/auto_deck_brief.md --output output/report.pptx
```

Windows + PowerPoint 環境で最終PNG確認まで行う場合:

```powershell
python src/auto_deck.py examples/auto_deck_brief.md --output output/report.pptx --render
```

GUI入口:

```powershell
python src/auto_deck_ui.py
```

詳細は `skills/auto-deck-pipeline/SKILL.md` を参照してください。

