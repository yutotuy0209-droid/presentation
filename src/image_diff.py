#!/usr/bin/env python3
"""Compare two folders of PowerPoint-rendered PNGs.

Uses Pillow when installed. If Pillow is unavailable, reports that visual
comparison is unavailable without failing structural round-trip tests.
"""
from __future__ import annotations
import json,sys
from pathlib import Path

def main():
    a,b=map(Path,sys.argv[1:3])
    try:
        from PIL import Image,ImageChops,ImageStat
    except Exception:
        print(json.dumps({"status":"unavailable","reason":"Install Pillow: python -m pip install Pillow"})); return
    aa=sorted(a.glob("*.PNG"))+sorted(a.glob("*.png")); bb=sorted(b.glob("*.PNG"))+sorted(b.glob("*.png"))
    rows=[]
    for x,y in zip(aa,bb):
        ia=Image.open(x).convert("RGB"); ib=Image.open(y).convert("RGB")
        if ia.size!=ib.size: rows.append({"a":x.name,"b":y.name,"equal":False,"reason":"size","sizes":[ia.size,ib.size]}); continue
        diff=ImageChops.difference(ia,ib); stat=ImageStat.Stat(diff)
        mean=sum(stat.mean)/len(stat.mean); bbox=diff.getbbox()
        rows.append({"a":x.name,"b":y.name,"equal":bbox is None,"mean_abs_difference":mean,"changed_bbox":bbox})
    print(json.dumps({"status":"ok","slides":rows,"count_a":len(aa),"count_b":len(bb)},ensure_ascii=False))
if __name__=="__main__": main()
