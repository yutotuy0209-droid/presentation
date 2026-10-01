# PythonでPPTXを逆変換・復元する

`src/pptx_bundle.py`は、PPTXを元パッケージ・XML・素材・編集用JSONへ展開し、再びPPTXへ戻すツールです。依存は`lxml`だけです。元ファイルと出力は上書きしません。

```powershell
python -m pip install lxml
python src/pptx_bundle.py unpack input.pptx --out output/source
python src/pptx_bundle.py pack output/source --out output/restored.pptx
```

## 取り出せるもの

`original.pptx`は元ファイルです。`package/`にはすべてのパッケージ部品を抽出します。画像は`package/ppt/media/`、グラフは`package/ppt/charts/`、埋め込みExcelなどは`package/ppt/embeddings/`にあります。マスター、テーマ、notes、グループ、アニメーション、SmartArt、数式、未知の拡張部品も元データとして保持します。存在するものだけが抽出されます。

`deck.json`の`slides`は実際の発表順の図形概要です。ファイル名の番号順とは限りません。`texts`にはスライド、表、notes、SmartArtなどのDrawingML文字列を部品名と番号で記録します。抽出した概要は完全なレイアウト形式ではなく、元のパッケージが復元の基準です。

`slides.md`も同時に生成します。ページ番号・本文・使用素材・notesをテキストで確認できます。目的や実験結果の意味は推測して書き足しません。Markdownは閲覧用で，PPTXへの編集反映はJSONの`texts[].text`で行います。

## 文章を変更する

`deck.json`の`texts`内にある`text`だけを変更してpackします。`original_text`、`part`、`index`、ハッシュは変えないでください。文章の実行例です。

```python
import json
from pathlib import Path

path = Path('output/source/deck.json')
deck = json.loads(path.read_text(encoding='utf-8'))
for item in deck['texts']:
    if item['part'] == 'ppt/slides/slide1.xml' and item['text'] == '旧タイトル':
        item['text'] = '新タイトル'
path.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding='utf-8')
```

編集対象のrunの書式、段落、他の図形を残して文字列を置換します。新しい段落やrunを作る機能ではありません。改行・長文の追加後はPowerPointで見た目を確認してください。数式は元のXMLとして保持しますが、数式を普通の文章へ変換しません。

## 図形や素材を変更する

`package/`の同じ部品名のファイルを変更すると、その部品だけを再出力へ反映します。同じ画像形式・対象部品を維持して画像を差し替えられます。座標や色はXML上級編集です。関係ID、部品名、図形IDを変更する場合、参照の整合性は編集者が管理してください。部品追加・削除やスライドの再構成をJSONだけで行う機能はありません。

未編集の往復はZIPも含むファイル全体のバイト一致です。編集時は未変更部品のバイトを保持し、変更部品のXMLを再直列化します。PPTXのSHAが同じとは限りません。XMLのコメント・名前空間・未知要素を保持しますが、変更部品のXML表現自体のバイト一致は保証しません。

## 対応範囲と限界

通常の暗号化されていないOOXML形式のPPTXを対象とします。古いバイナリ`.ppt`、パスワード付きファイル、破損ファイルは対象外です。外部リンクの内容は取得せず参照だけを残します。埋め込みコード・動画を実行しません。署名付きパッケージは未編集復元のみ許可します。解凍先を逸脱するパス、重複部品名、1GiBを超える展開は拒否します。

「あらゆるPPTXを、意味まで理解して同じPython図形へ変換できる」という保証はしません。対応できない要素を削除せず、元の部品として保持することで逆変換の損失を抑えます。既存`pptx_to_ir.py`の図形座標・文章抽出は解析向け、このbundleは復元・素材再利用向けです。

```powershell
python -m unittest discover -s tests -p test_pptx_bundle.py
```
