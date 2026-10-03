# Design Preference Loop

PowerPointの「デザインセンス」を長いpromptへ保存するのではなく、**比較 → 人間選択 → 使用 → 短い要約**として蓄積する。

```text
content
  ↓
4 design candidates
  ↓
human click / short comfort feedback / elapsed time
  ↓
choice history
  ↓
usage + recency summary
  ↓
next presentation: top 3 + exploration 1
```

## Why

文章ルールを増やすと、常時コンテキストが肥大化し、重要度の低い好みまで毎回モデルへ渡される。

その代わり、

- 生の選択履歴はCSV/JSONL
- 使用回数・選択回数は集計
- OpenCode等が読むのは短いsummaryのみ
- presentation/design taskのときだけ読む

とする。

## Signals

候補順には次を使える。

- 過去の選択回数
- 実際にbuildした回数
- 直近選択
- 発表種別との一致
- 「見やすい / 普通 / 眩しい・疲れる」の短い人間評価

目の負担について医学的な効果を断定しない。ここで保存するのは本人の主観的な見やすさ評価。

## Exploration

常に前回案だけを再生すると好みが固定されるため、4案なら原則 **3案を履歴優先、1案を探索** とする。

単発の選択をSkillやPresetへ昇格しない。複数回・複数資料で繰り返された修正だけを候補にする。

## Usage count

「生成した回数」と「人間が選んだ回数」と「実際にbuildした回数」を混ぜない。

少なくとも、

`selected_count`, `build_count`, `last_selected_at`

を区別する。

## Context budget

常時ルールには「design preference summaryをpresentation taskのときだけ読む」という1行だけを置く。

履歴全文は通常コンテキストへ入れない。
