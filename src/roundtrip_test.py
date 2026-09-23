#!/usr/bin/env python3
"""Prepare a PPTX round-trip fidelity check.

Step 1 parses original PPTX to IR. The generated PPTX can then be parsed again
and compared with compare_ir.py. Kept as a small orchestrator to make fidelity
testing repeatable.
"""
from __future__ import annotations
import argparse,subprocess,sys
from pathlib import Path

def main():
 ap=argparse.ArgumentParser(); ap.add_argument("pptx",type=Path); ap.add_argument("--work",type=Path,default=Path("output/roundtrip"))
 a=ap.parse_args(); a.work.mkdir(parents=True,exist_ok=True)
 original=a.work/"original.json"; rebuilt=a.work/"rebuilt.pptx"; rebuilt_ir=a.work/"rebuilt.json"; report=a.work/"generation-report.json"
 here=Path(__file__).resolve().parent
 subprocess.run([sys.executable,str(here/"pptx_to_ir.py"),str(a.pptx),"--out",str(original)],check=True)
 subprocess.run(["node",str(here/"ir_to_pptx.js"),str(original),str(rebuilt),str(report)],check=True)
 subprocess.run([sys.executable,str(here/"pptx_to_ir.py"),str(rebuilt),"--out",str(rebuilt_ir)],check=True)
 subprocess.run([sys.executable,str(here/"compare_ir.py"),str(original),str(rebuilt_ir),"--out",str(a.work/"comparison.json")],check=True)
 print(a.work)
if __name__=="__main__": main()
