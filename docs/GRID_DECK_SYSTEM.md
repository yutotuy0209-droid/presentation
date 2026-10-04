# 2D Grid Deck System

PowerPointを毎回AIに座標指定させず、**意味ブロック + 2次元Grid + Design Preset** から決定論的に生成する。

目的:

> 一目で構造、数秒で関係、読めば具体。

---

## 1. 基本DSL

### 1ブロック

```json
{
  "grid": [["A"]],
  "blocks": {
    "A": {"type": "card", "title": "結論", "body": "詳細"}
  }
}
```

### 左右

```json
{
  "grid": [["A", "B"]]
}
```

### 上下

```json
{
  "grid": [
    ["A"],
    ["B"]
  ]
}
```

### 大1 + 小2

同じIDを矩形状に繰り返すとspanになる。

```json
{
  "grid": [
    ["photo", "result"],
    ["photo", "note"]
  ]
}
```

---

## 2. Connectorは配置と分離

```json
{
  "connectors": [
    {"from": "input", "to": "process"},
    {"from": "process", "to": "output"}
  ]
}
```

配置はGrid、意味関係はConnectorで表す。

これにより同じ3ブロックを、

- 比較
- 手順
- 因果
- 入出力

へ使い分けられる。

---

## 3. Block types

| type | 用途 |
|---|---|
| `card` | 見出し + 短い説明 |
| `metric` | 数値を主役にする |
| `text` | 文章中心 |
| `image` | PNG/JPEG等 |
| `svg` | 関係図・模式図 |
| `gif` | 短い動作説明 |
| `video` | MP4等 |

GIF/SVGはPowerPoint上では画像、videoはembedded mediaとして出力する。

---

## 4. Design presets

共通の構造ルールを保ったまま見た目だけ切り替える。

```json
"design": "glanceable-scientific"
```

候補:

- `glanceable-scientific`
- `glanceable-soft-graphite`
- `glanceable-warm`

意味色:

- input
- process
- output
- human
- warning
- neutral

AIはRGB値を毎回決めず、roleだけ選ぶ。

---

## 5. Glanceable hierarchy

### 3秒

- タイトル
- 大きな固まり
- 左→右 / 上→下の流れ
- 何が重要か

### 10秒

- 各Blockの役割
- 入力 / 処理 / 出力
- Human / System
- 矢印の意味

### 30秒

- 数値
- 条件
- ファイル名
- 詳細説明

詳細を消すのではなく、視覚優先度を下げる。

---

## 6. Animation / GIF / Video

### L0: Static

- PNG
- SVG
- editable PowerPoint shapes

### L1: Embedded motion

- GIF
- MP4

短い処理・カメラ動作・計測変化はGIFを優先する。

### L2: Native PowerPoint animation

DSLでは先に予約できる。

```json
{
  "animation": {
    "entrance": "fade",
    "order": 2,
    "trigger": "click"
  }
}
```

現時点ではこの情報をIRへ保存するが、PPTX writerは適用しない。
generation reportへ `animationReservations` として出す。

将来Windows PowerPoint COM backendからAppear/Fade等を適用できる。

---

## 7. Build

```bash
python src/grid_deck.py build examples/grid_deck_glanceable.json \
  --out-spec output/grid.pixel.json \
  --out-ir output/grid.ir.json \
  --pptx output/grid.pptx \
  --strict
```

内部:

```text
Grid DSL
   ↓
grid_deck.py
   ↓
Pixel Layout Spec
   ↓
pixel_layout_engine.py
   ↓
Slide IR
   ↓
lint
   ↓
ir_to_pptx.js
   ↓
editable PPTX
```

---

## 8. Design rules

- 主要Blockは原則3〜7個。
- Blockタイトルだけでも全体構造を説明できる。
- Card本文は原則1〜4行。
- 色は装飾ではなく意味分類に使う。
- Connectorは必要な関係だけ。
- 画像・GIF・SVGを主役にできる。
- 14pt未満の本文を前提にしない。
- サムネイル表示でも構造が分かること。
- AIは微調整ではなく、Content / Grid / Roleを決める。
- 座標・余白・基本フォントはシステムが決める。

---

## 9. 将来

- Grid template gallery
- 人間選択回数によるtemplate優先順位
- PowerPoint COM native animation backend
- GIF/MP4 poster frame
- Markdown → Grid DSL adapter
- codeconvert relationship JSON → Grid Deck
