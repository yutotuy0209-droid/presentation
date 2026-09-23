#!/usr/bin/env python3
"""PPTX -> loss-aware Slide IR.

Uses only Python's standard library. The IR stores native PowerPoint geometry in
EMU so a PPTX round trip does not need HTML/CSS as an intermediate format.
"""
from __future__ import annotations
import argparse, json, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS={
 "p":"http://schemas.openxmlformats.org/presentationml/2006/main",
 "a":"http://schemas.openxmlformats.org/drawingml/2006/main",
 "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS="http://schemas.openxmlformats.org/package/2006/relationships"

def q(ns,tag): return "{"+NS[ns]+"}"+tag
def ints(node, xpath):
    x=node.find(xpath,NS)
    return int(x.get("x","0")),int(x.get("y","0")) if x is not None else (0,0)

def geometry(node):
    xfrm=node.find(".//a:xfrm",NS) or node.find(".//p:xfrm",NS)
    if xfrm is None: return {"x":0,"y":0,"width":0,"height":0,"rotation":0}
    off=xfrm.find("a:off",NS); ext=xfrm.find("a:ext",NS)
    if off is None: off=xfrm.find("p:off",NS)
    if ext is None: ext=xfrm.find("p:ext",NS)
    return {
      "x":int(off.get("x","0")) if off is not None else 0,
      "y":int(off.get("y","0")) if off is not None else 0,
      "width":int(ext.get("cx","0")) if ext is not None else 0,
      "height":int(ext.get("cy","0")) if ext is not None else 0,
      "rotation":int(xfrm.get("rot","0"))/60000.0,
    }

def text_info(node):
    paras=[]
    runs=[]
    for p in node.findall(".//a:p",NS):
        parts=[]
        for r in p.findall("a:r",NS):
            t=r.find("a:t",NS); value=t.text if t is not None and t.text else ""
            rp=r.find("a:rPr",NS)
            info={"text":value}
            if rp is not None:
                if rp.get("sz"): info["fontSizePt"]=int(rp.get("sz"))/100
                if rp.get("b"): info["bold"]=rp.get("b")=="1"
                if rp.get("i"): info["italic"]=rp.get("i")=="1"
                latin=rp.find("a:latin",NS)
                if latin is not None and latin.get("typeface"): info["fontFamily"]=latin.get("typeface")
            runs.append(info); parts.append(value)
        paras.append("".join(parts))
    return "\n".join(paras),runs

def rels(z, slide_no):
    name=f"ppt/slides/_rels/slide{slide_no}.xml.rels"
    if name not in z.namelist(): return {}
    root=ET.fromstring(z.read(name))
    return {r.get("Id"):r.get("Target") for r in root.findall("{"+REL_NS+"}Relationship")}

def parse_pptx(path:Path):
    with zipfile.ZipFile(path) as z:
        pres=ET.fromstring(z.read("ppt/presentation.xml"))
        size=pres.find("p:sldSz",NS)
        deck={"version":2,"unit":"emu","meta":{"source":"pptx","filename":path.name},
              "slideSize":{"width":int(size.get("cx")),"height":int(size.get("cy"))},"slides":[]}
        slide_names=sorted(
          [n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
          key=lambda n:int(Path(n).stem.replace("slide","")))
        for slide_name in slide_names:
            no=int(Path(slide_name).stem.replace("slide",""))
            root=ET.fromstring(z.read(slide_name)); relationship=rels(z,no)
            spTree=root.find(".//p:spTree",NS); elements=[]
            if spTree is not None:
                for zi,node in enumerate(list(spTree)):
                    local=node.tag.split("}")[-1]
                    if local not in {"sp","pic","graphicFrame","cxnSp","grpSp"}: continue
                    nv=node.find(".//p:cNvPr",NS)
                    eid=nv.get("id") if nv is not None else str(zi)
                    name=nv.get("name") if nv is not None else local
                    g=geometry(node); text,runs=text_info(node)
                    el={"id":eid,"name":name,"type":local,"z":zi,**g,"text":text,"runs":runs}
                    if local=="pic":
                        blip=node.find(".//a:blip",NS)
                        rid=blip.get(q("r","embed")) if blip is not None else None
                        if rid: el["mediaTarget"]=relationship.get(rid)
                    elements.append(el)
            deck["slides"].append({"id":f"slide-{no}","number":no,"elements":elements})
        return deck

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("pptx",type=Path); ap.add_argument("--out",type=Path,required=True)
    a=ap.parse_args(); data=parse_pptx(a.pptx)
    a.out.parent.mkdir(parents=True,exist_ok=True); a.out.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")
if __name__=="__main__": main()
