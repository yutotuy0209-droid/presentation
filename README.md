# presentation

再利用可能なプレゼン資料生成コードを置くためのリポジトリです。

卒業研究リポジトリで作成した資料生成処理のうち、研究固有のデータ・文章・実験結果を含まない部分だけを切り出しています。

## Current scope

- HTML / CSS スライドのレンダリング
- Puppeteer による高解像度 PNG 出力
- PNG から PowerPoint (PPTX) への変換
- レンダリング用スモークテスト

## Structure

```text
presentation/
├─ src/
│  └─ html_to_pptx.js
├─ tests/
│  ├─ test_puppeteer.js
│  └─ fixtures/
│     └─ test_slide.html
├─ package.json
└─ .gitignore
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

## Origin

汎用処理は `graduation-research-motion-analysis` 内で作成したプレゼン生成コードを基にしています。
研究固有の内容はこのリポジトリにはコピーしません。
