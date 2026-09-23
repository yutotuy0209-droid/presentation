# presentation

個人利用を中心に、再利用可能なプレゼン資料生成コード・構成パターン・変換時の注意点を保存するリポジトリです。

卒業研究で得たノウハウのうち、研究固有のデータ・文章・実験結果を含まない部分だけを切り出しています。

## Current scope

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
