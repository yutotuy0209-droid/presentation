# Real-PC round-trip loop

PowerPointとの「＝」を近づけるため、実PC上で同じPPTXを反復変換・比較するハーネス。

## Why

HTMLやブラウザを正解にしない。Windows上にMicrosoft PowerPointがある場合はPowerPoint自身でPNGを書き出し、構造比較と視覚比較を分離する。

## Run

構造ループ:

```powershell
.\scripts\run_roundtrip_loop.ps1 -Pptx .\sample.pptx -Iterations 5
```

PowerPoint実レンダリングも含む:

```powershell
.\scripts\run_roundtrip_loop.ps1 -Pptx .\sample.pptx -Iterations 5 -Visual
```

Visual比較には任意でPillowを使う:

```powershell
python -m pip install Pillow
```

## Output

`output/loop/iteration-XX/` に input IR, rebuilt PPTX, rebuilt IR, comparison, generation report, PNG renderを保存し、全体を `history.json` に追記する。

## Stop behavior

- 完全一致オプションを使えば差分ゼロで停止できる。
- 連続2回で差分件数が変わらなければ、現実装の限界で安定したとみなし停止する。
- 未対応Shapeはgeneration reportへ残す。

このループは自動でコードを書き換えない。まず再現可能な差分データを集め、改善コードと教師データを分離する。
