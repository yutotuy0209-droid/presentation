# ROADMAP

候補を消さずに残すための設計メモ。実装済み・未実装を混同しないことを優先する。

## Core architecture

```text
PPTX ─┐
HTML ─┼─> common slide IR (JSON) ─> PPTX / HTML / SVG
MD   ─┘               │
                      ├─> lint / QA
human-edited PPTX ────┴─> diff ─> experience
```

共通IRを中心にし、PPTXとHTMLを完全同一にするのではなく、意味・位置・サイズ・スタイルを可能な範囲で往復可能にする。

## Candidate backlog

- [ ] P01 PPTX → common JSON
- [ ] P02 common JSON → PPTX
- [ ] P03 HTML → common JSON
- [ ] P04 common JSON → HTML
- [ ] P05 PPTX ↔ HTML round trip via IR
- [ ] P06 AI生成版と人間修正版の差分抽出
- [ ] P07 修正履歴をJSONL/CSVへ保存
- [ ] P08 overlap検出
- [ ] P09 overflow検出
- [ ] P10 alignment検出
- [ ] P11 小さすぎる文字の検出
- [ ] P12 文字密度・情報過多の検出
- [ ] P13 LaTeX/数式 → SVG
- [ ] P14 数式変換前後の検査
- [ ] P15 SVG → PPTX
- [ ] P16 PPTX → SVG/HTML
- [ ] P17 Markdown → slides
- [ ] P18 Marp連携
- [ ] P19 グラフの自動配置
- [ ] P20 表の自動配置
- [ ] P21 before/after visual comparison
- [ ] P22 slide linter
- [x] P23 scientific preset
- [x] P24 research-progress preset
- [ ] P25 thesis-defense preset
- [ ] P26 business preset
- [ ] P27 poster preset
- [ ] P28 人間修正履歴からpreset改善候補を生成
- [ ] P29 PPTX全ページの画像化
- [ ] P30 visual regression

## Experience loop

人間によるPowerPoint上の修正を失敗扱いにしない。生成前後の差分を保存し、繰り返される修正だけを将来のpreset/skill候補にする。

特に数式は完全自動化を前提にせず、意味保存 → 視覚品質 → 編集性の順で扱う。

## Non-goals for now

- PowerPointそのものを置き換えること
- すべての変換を完全可逆にすること
- 人間の最終確認をなくすこと
