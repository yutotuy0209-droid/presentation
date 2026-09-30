# Pixel Layout Engine

## 目的

`src/pixel_layout_engine.py` は、PowerPoint生成時にAIへ座標・余白・フォントサイズ・文字被り修正を何度も考えさせないための決定論的レイアウトエンジンです。

AIは原則として「何を伝えるか」だけを短いJSONで指定します。

Python側が、

- 1px単位の座標
- 余白
- レイアウト
- 文字量
- 文字サイズ
- 改ページ
- 衝突回避
- 画像サイズ
- 表サイズ
- PowerPoint座標への変換

を担当します。

これにより、毎回LLMへ巨大なPowerPoint生成コードを出力させる必要をなくし、トークン消費を抑えます。

---

## 設計

```text
短い内容JSON
   ↓
pixel_layout_engine.py
   ↓
1px layout solver
   ↓
common Slide IR v2 (EMU)
   ↓
lint_ir.js
   ↓
ir_to_pptx.js
   ↓
editable PPTX
   ↓
PowerPoint renderer（任意）
   ↓
PNG確認
```

HTMLをPowerPointの正本にはしません。

最終的なPowerPoint geometryは、既存リポジトリ方針どおり common IR / EMU を経由します。

---

## なぜpxで考えるのか

PowerPoint内部はEMUで座標を保持しますが、人間やAIが毎回巨大なEMU値を書くのは非効率です。

本ツールでは入力をpxに統一します。

96 dpiの場合:

- 1 px = 9,525 EMU
- 960 × 720 px = 10 × 7.5 inch = 4:3
- 1280 × 720 px = 13.333... × 7.5 inch = 16:9

したがって、スライド作成側は960×720等の通常の画像座標感覚で扱えます。

---

## 最小入力

```json
{
  "slides": [
    {
      "title": "研究背景と目的",
      "bullets": [
        "工作機械では熱変位が加工精度へ影響する",
        "外付け画像計測でX・Y・Z方向の微小変位を測定する"
      ]
    }
  ]
}
```

canvas未指定時は4:3です。

つまり、単純なスライドであればAIが指定する必要があるのはタイトルと内容だけです。

---

## Shorthand

AIのトークン削減のため、以下を直接使用できます。

### body

```json
{
  "title": "結果",
  "body": "実験結果の説明"
}
```

### bullets

```json
{
  "title": "今週の進捗",
  "bullets": [
    "Step 0を実施",
    "Step 1を実施",
    "Step 2を実施"
  ]
}
```

### image

```json
{
  "title": "実験画像",
  "body": "実機画像と解析結果を比較する",
  "image": "../assets/result.png"
}
```

本文と画像が同時に存在する場合、layout未指定なら画像付きレイアウトを自動選択します。

### table

```json
{
  "title": "結果一覧",
  "table": [
    ["項目", "値"],
    ["X", "10.2 µm"],
    ["Y", "3.1 µm"]
  ]
}
```

---

## blocks

細かく指定したい場合のみ `blocks` を使います。

```json
{
  "title": "研究フロー",
  "layout": "two-column",
  "blocks": [
    {
      "type": "bullets",
      "slot": "left",
      "items": ["Step 0", "Step 1"]
    },
    {
      "type": "bullets",
      "slot": "right",
      "items": ["Step 2", "Step 3"]
    }
  ]
}
```

利用可能な主要type:

- `text`
- `bullets`
- `callout`
- `metric`
- `image`
- `table`
- `spacer`
- `line`
- `box`

---

## layout

### title-body

標準縦積み。

### two-column

左右2列。

### image-left

画像を左、文章を右。

### image-right

文章を左、画像を右。

### auto / 未指定

内容から自動選択。

---

## 座標を残すハイブリッド制御

完全自動配置ではなく、**重要な構図は座標で固定し、細部だけ自動化する**ことを推奨します。

指定レベルは3段階です。

### 1. fixed — 完全固定

```json
{
  "type": "text",
  "text": "ここは絶対に動かさない",
  "position": {
    "x": 80,
    "y": 160,
    "w": 360,
    "h": 100,
    "mode": "fixed"
  }
}
```

- x / y / w / h を1px単位で厳守
- 他要素と重なればERROR
- ボックス外に文字が収まらなければERROR
- Pythonが勝手に位置を変更しない

研究図、写真、重要な結論、左右比較など、構図上の基準点に使います。

### 2. prefer — 座標を優先しつつ微調整

```json
{
  "type": "text",
  "text": "この辺りに置きたい",
  "position": {
    "x": 80,
    "y": 300,
    "w": 360,
    "h": 90,
    "mode": "prefer",
    "axis": "y",
    "maxNudgePx": 24
  }
}
```

- まず指定座標へ配置
- 衝突時だけ1px単位で移動
- `maxNudgePx` より遠くへは逃がさない
- `axis: "x" | "y" | "both"` で動かす方向を制限

「位置はほぼ決めたいが、1〜数十px程度の調整はPythonへ任せたい」要素に使います。

### 3. flow — 自動配置

座標を省略した通常の本文等です。

- semantic region内で自動配置
- 衝突時は空き領域を探索
- 長文は縮小・改ページ

### Named regions

1つずつ座標を書く代わりに、スライド側で領域を固定できます。

```json
{
  "title": "実験結果",
  "regions": {
    "photo":  {"x": 60,  "y": 150, "w": 420, "h": 450},
    "result": {"x": 520, "y": 150, "w": 370, "h": 450}
  },
  "blocks": [
    {"type": "image", "slot": "photo", "path": "result.png"},
    {"type": "text", "slot": "result", "text": "結果の説明"}
  ]
}
```

これにより、AIは「写真は左領域、説明は右領域」とだけ指定でき、各要素の細かなy座標はPythonに任せられます。

### タイトル位置

```json
{
  "title": "研究背景",
  "titleBox": {"x": 70, "y": 32, "w": 820, "h": 55}
}
```

過去PowerPointのタイトル位置を再現するときに使用します。

### 推奨バランス

1枚のスライドについて、

- タイトル位置
- 主画像
- 主図
- 結論
- 左右カラム境界

など**構図を決める数個の要素だけ座標指定**し、
本文や注釈の細かな位置調整はPythonに任せます。

この方式なら、デザイン意図を保持しつつ、AIが全shapeのx/yを毎回出力する必要もありません。

---

## 1px衝突回避

foregroundとして以下を扱います。

- text
- picture
- table

各要素は整数pxの矩形として扱います。

配置候補が既存要素と重なった場合:

1. 希望Xを維持
2. Yを1pxずつ下へ探索
3. それでも置けない場合はsemantic zone全体を1px単位で探索
4. 空き領域がなければLayoutError

つまり「被っているがとりあえずPPTXを出力する」という挙動にはしません。

---

## 文字の自動処理

テキストについては、全角/ASCIIのおおよその文字幅を使って必要行数を推定します。

処理順:

1. 希望フォントサイズで入るか確認
2. 0.5ptずつ縮小
3. `minBodySizePt`まで縮小
4. それでも入らなければ文章を分割
5. 継続スライドを自動生成
6. タイトルに「（続き）」を付与

そのため、AIが長文の分割位置を毎回決める必要はありません。

---

## 画像

PNG / JPEGはPython標準ライブラリだけで画像寸法を読み取ります。

画像は指定領域に対してcontainされ、縦横比を維持します。

画像とテキストは同じforegroundとして扱うため、通常の生成では重なりません。

---

## PowerPoint品質ゲート

`--strict` を付けると、内部layout validationに加えて `src/lint_ir.js` も通します。

主なERROR:

- overflow
- foreground overlap
- 極端に小さい本文文字

PPTX出力前にERRORがあれば停止します。

---

## 1コマンド生成

```bash
python src/pixel_layout_engine.py build spec.json \
  --out-ir output/deck.json \
  --pptx output/deck.pptx \
  --strict
```

生成物:

- `deck.json`
- `deck.layout-report.json`
- `deck.lint.json`
- `deck.pptx`
- `deck.generation-report.json`

---

## PowerPoint本体でレンダリング

Windows + Microsoft PowerPoint環境:

```powershell
python src/pixel_layout_engine.py build spec.json `
  --out-ir output/deck.json `
  --pptx output/deck.pptx `
  --render-dir output/rendered `
  --strict
```

既存の `src/render_pptx_powerpoint.ps1` を使用します。

PowerPoint自身をレンダラーにするため、HTML/ブラウザ表示との差を最終確認できます。

---

## AIへ渡す情報量

従来:

```text
「タイトルをx=0.8, y=0.5...
本文をx=...
文字サイズは...
画像は...
重なったから直して...
」
```

Pixel Layout Engine:

```json
{
  "slides": [
    {
      "title": "背景",
      "bullets": ["課題A", "課題B"],
      "image": "machine.png"
    }
  ]
}
```

レイアウト知識をコード側へ固定することで、毎回同じ説明をLLMへ渡す必要を減らします。

---

## 研究リポジトリとの分離

このツールには研究固有の数値や文章を埋め込みません。

- `presentation`: 生成・配置・検査
- `graduation-research-motion-analysis`: 研究内容・実験結果

という分離を維持します。

卒論スライドを作る場合も、研究内容は研究リポジトリから取得し、短いspecへ変換して本ツールへ渡します。

---

## 今後の拡張候補

- 表の行単位自動改ページ
- native chart
- SVGのeditable変換
- group / connector
- speaker notes
- 過去PPTXからthemeを自動生成
- PowerPointレンダリング画像を使った視覚的自動評価
- 人間修正履歴からmargin/font/presetを学習
- 画像内の注釈位置最適化

これらもAIへ毎回判断させるのではなく、再利用可能なアルゴリズムとして実装する。
