#!/usr/bin/env python3
"""Automated real-PC PPTX round-trip loop.

Repeatedly:
1) PPTX -> IR
2) IR -> rebuilt PPTX
3) rebuilt PPTX -> IR
4) compare IR
5) optionally render both PPTX files through local PowerPoint
6) measure image difference
7) archive each iteration and stop on threshold/stability.

Designed for Windows + Microsoft PowerPoint, but structural checks work anywhere.
"""
from __future__ import annotations
import argparse, json, shutil, subprocess, sys, time
from pathlib import Path

HERE=Path(__file__).resolve().parent

def run(cmd):
    p=subprocess.run(cmd,text=True,capture_output=True)
    if p.returncode:
        raise RuntimeError(f"command failed: {cmd}\nSTDOUT:\n{p.stdout}\nSTDERR:\n{p.stderr}")
    return p.stdout.strip()

def load(p): return json.loads(Path(p).read_text(encoding="utf-8"))

def counts(c):
    return {
      "semantic":len(c["semantic"]["differences"]),
      "structural":len(c["structural"]["differences"]),
      "geometry":len(c["geometry"]["differences"]),
    }

def render_pptx(pptx:Path,out:Path):
    ps=HERE/"render_pptx_powerpoint.ps1"
    if not ps.exists(): return {"status":"unavailable","reason":"renderer script missing"}
    out.mkdir(parents=True,exist_ok=True)
    try:
        run(["powershell","-NoProfile","-ExecutionPolicy","Bypass","-File",str(ps),"-InputPptx",str(pptx.resolve()),"-OutputDir",str(out.resolve())])
        return {"status":"ok"}
    except Exception as e: return {"status":"error","reason":str(e)}

def image_diff(a:Path,b:Path):
    script=HERE/"image_diff.py"
    try:
        raw=run([sys.executable,str(script),str(a),str(b)])
        return json.loads(raw)
    except Exception as e: return {"status":"unavailable","reason":str(e)}

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("pptx",type=Path)
    ap.add_argument("--work",type=Path,default=Path("output/loop"))
    ap.add_argument("--iterations",type=int,default=5)
    ap.add_argument("--geometry-tolerance-emu",type=int,default=12700)
    ap.add_argument("--visual",action="store_true",help="Use installed Microsoft PowerPoint for PNG rendering.")
    ap.add_argument("--stop-when-structural-equal",action="store_true")
    args=ap.parse_args()
    args.work.mkdir(parents=True,exist_ok=True)
    source=args.pptx.resolve()
    history=[]
    previous_counts=None
    for n in range(1,args.iterations+1):
        d=args.work/f"iteration-{n:02d}"; d.mkdir(parents=True,exist_ok=True)
        original_ir=d/"input.json"; rebuilt=d/"rebuilt.pptx"; rebuilt_ir=d/"rebuilt.json"; comparison=d/"comparison.json"; generation=d/"generation.json"
        run([sys.executable,str(HERE/"pptx_to_ir.py"),str(source),"--out",str(original_ir)])
        run(["node",str(HERE/"ir_to_pptx.js"),str(original_ir),str(rebuilt),str(generation)])
        run([sys.executable,str(HERE/"pptx_to_ir.py"),str(rebuilt),"--out",str(rebuilt_ir)])
        run([sys.executable,str(HERE/"compare_ir.py"),str(original_ir),str(rebuilt_ir),"--out",str(comparison),"--tolerance-emu",str(args.geometry_tolerance_emu)])
        c=counts(load(comparison)); item={"iteration":n,"input":str(source),"rebuilt":str(rebuilt),"differences":c,"generation":load(generation)}
        if args.visual:
            ra=render_pptx(source,d/"render-input"); rb=render_pptx(rebuilt,d/"render-rebuilt")
            item["render"]={"input":ra,"rebuilt":rb}
            if ra.get("status")=="ok" and rb.get("status")=="ok":
                item["visual"]=image_diff(d/"render-input",d/"render-rebuilt")
        history.append(item)
        (args.work/"history.json").write_text(json.dumps(history,ensure_ascii=False,indent=2),encoding="utf-8")
        print(json.dumps(item,ensure_ascii=False))
        if args.stop_when_structural_equal and c["semantic"]==0 and c["structural"]==0 and c["geometry"]==0: break
        if previous_counts==c and n>=2:
            # Loop has stabilized; continuing would only reproduce the same current limitation.
            break
        previous_counts=c
        source=rebuilt.resolve()
    print(f"history: {args.work/'history.json'}")
if __name__=="__main__": main()
