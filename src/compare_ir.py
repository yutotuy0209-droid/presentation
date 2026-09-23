#!/usr/bin/env python3
"""Compare two PPTX-derived IR files at semantic, structural and visual-geometry levels."""
from __future__ import annotations
import argparse,json
from pathlib import Path

def compare(a,b,tol=12700): # ~1 pt
    out={"semantic":{"equal":True,"differences":[]},"structural":{"equal":True,"differences":[]},"geometry":{"equal":True,"differences":[]}}
    if len(a.get("slides",[]))!=len(b.get("slides",[])):
        out["structural"]["equal"]=False; out["structural"]["differences"].append({"slides":[len(a.get("slides",[])),len(b.get("slides",[]))]})
    for i,(sa,sb) in enumerate(zip(a.get("slides",[]),b.get("slides",[])),1):
        ea=sa.get("elements",[]); eb=sb.get("elements",[])
        if len(ea)!=len(eb):
            out["structural"]["equal"]=False; out["structural"]["differences"].append({"slide":i,"elements":[len(ea),len(eb)]})
        for x,y in zip(ea,eb):
            if x.get("type")!=y.get("type"):
                out["structural"]["equal"]=False; out["structural"]["differences"].append({"slide":i,"a":x.get("type"),"b":y.get("type")})
            if x.get("text","")!=y.get("text",""):
                out["semantic"]["equal"]=False; out["semantic"]["differences"].append({"slide":i,"element":x.get("id"),"before":x.get("text",""),"after":y.get("text","")})
            for k in ("x","y","width","height"):
                if abs(int(x.get(k,0))-int(y.get(k,0)))>tol:
                    out["geometry"]["equal"]=False; out["geometry"]["differences"].append({"slide":i,"element":x.get("id"),"field":k,"before":x.get(k),"after":y.get(k),"delta":int(y.get(k,0))-int(x.get(k,0))})
    return out

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("before",type=Path); ap.add_argument("after",type=Path); ap.add_argument("--out",type=Path); ap.add_argument("--tolerance-emu",type=int,default=12700)
    x=ap.parse_args(); result=compare(json.loads(x.before.read_text(encoding="utf-8")),json.loads(x.after.read_text(encoding="utf-8")),x.tolerance_emu)
    s=json.dumps(result,ensure_ascii=False,indent=2)
    if x.out: x.out.write_text(s,encoding="utf-8")
    else: print(s)
if __name__=="__main__": main()
