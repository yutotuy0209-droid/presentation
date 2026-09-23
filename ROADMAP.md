# ROADMAP

候補を消さずに残す設計メモ。HTMLを正本にせず、PowerPointのnative geometryに近い共通IRを中心にする。

```text
                    ┌──────────── PPTX
                    │               ↑↓
source/generator ↔ Slide IR (EMU) ← human-edited PPTX
                    │
                    ├─ SVG
                    ├─ PNG
                    └─ HTML (preview/helper only)
```

## Equality

1. Semantic equal: 内容・意味が同じ
2. Structural equal: shape種類・個数・順序が同じ
3. Geometry/visual equal: x/y/width/height等が許容差内

未対応要素は黙って近似せず unsupported / approximate / rasterized としてreportする。

## Status

- [x] PPTX → common IR v2 (EMU)
- [~] common IR → PPTX: editable text shapeから開始
- [x] HTML → helper IR: 主経路ではない
- [x] IR before/after diff foundation
- [x] 修正履歴JSONL foundation
- [x] overlap / overflow / tiny-text / text-density lint foundation
- [x] semantic / structural / geometry comparison foundation
- [x] round-trip fidelity harness
- [ ] picture/media lossless round trip
- [ ] groups/connectors
- [ ] tables/charts
- [ ] theme/master inheritance
- [ ] native equations + equation meaning checks
- [ ] SVG → editable/faithful PPTX handling
- [ ] PPTX render screenshots + visual regression
- [ ] repeated human corrections → preset suggestions
- [ ] thesis-defense / business / poster presets

## Round-trip

```bash
python src/pptx_to_ir.py original.pptx --out original.json
node src/ir_to_pptx.js original.json rebuilt.pptx generation-report.json
python src/pptx_to_ir.py rebuilt.pptx --out rebuilt.json
python src/compare_ir.py original.json rebuilt.json --out comparison.json

# 一括
python src/roundtrip_test.py original.pptx --work output/roundtrip
```

Current limitation: readerはcommon shapes/text/picture・graphic frame metadataとnative EMU geometryを抽出する。writerはまずeditable text shapeを再構築する段階。pictures, groups, tables/charts, connectors, theme inheritance, native equationsは専用handlerが必要。

## Experience loop

AI生成 → PPTX → 人間修正 → PPTX再解析 → IR差分 → experience。

繰り返される人間修正だけをpreset/skill候補にする。

## Non-goals

- HTML/CSSをPowerPointの正本にする
- 未対応PowerPoint機能を完全可逆と偽る
- 人間の最終確認をなくす
