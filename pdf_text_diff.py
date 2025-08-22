#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Clean PDF text diff using pdfplumber + difflib only.
- Извлекает текст через pdfplumber (pdfminer.six)
- Нормализует переносы и пробелы
- Делит на предложения
- Сравнивает списки предложений с difflib.SequenceMatcher
- Результат: JSON + HTML с <del>/<ins>
"""

import argparse, re, json, pathlib
from dataclasses import dataclass
from typing import List, Dict, Tuple
from difflib import SequenceMatcher

# -------- нормализация --------
WS = re.compile(r"\s+")
HYPHEN = re.compile(r"(\w)-\n(\w)")
SENT_SPLIT = re.compile(r'(?<=[\.\!\?])\s+(?=[A-ZА-Я0-9])', re.UNICODE)

def normalize(s: str) -> str:
    if not s: return ""
    s = s.replace("\r", "")
    s = HYPHEN.sub(r"\1\2", s)    # убираем переносы по дефису
    s = s.replace("\n", " ")
    s = WS.sub(" ", s).strip()
    return s

def to_sentences(txt: str) -> List[str]:
    txt = normalize(txt)
    if not txt: return []
    if SENT_SPLIT.search(txt):
        return [p.strip() for p in SENT_SPLIT.split(txt) if p.strip()]
    return [txt]

@dataclass
class Sentence:
    text: str
    page: int

# -------- извлечение текста --------
def extract_pdfplumber(path: str) -> List[Sentence]:
    import pdfplumber
    out: List[Sentence] = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            txt = page.extract_text() or ""
            for s in to_sentences(txt):
                out.append(Sentence(s, i+1))
    return out

# -------- дифф --------
def token_diff(a: str, b: str) -> List[Tuple[str,str]]:
    A, B = a.split(), b.split()
    sm = SequenceMatcher(a=A, b=B, autojunk=False)
    out: List[Tuple[str,str]] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            out.append(("eq"," ".join(A[i1:i2])))
        elif tag == "delete":
            out.append(("del"," ".join(A[i1:i2])))
        elif tag == "insert":
            out.append(("ins"," ".join(B[j1:j2])))
        elif tag == "replace":
            if i1<i2: out.append(("del"," ".join(A[i1:i2])))
            if j1<j2: out.append(("ins"," ".join(B[j1:j2])))
    return out

def align_sentence_lists(old: List[Sentence], new: List[Sentence]) -> List[Dict]:
    sm = SequenceMatcher(a=[s.text for s in old], b=[s.text for s in new], autojunk=False)
    diffs: List[Dict] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal": continue
        if tag in ("replace","delete"):
            for i in range(i1, i2):
                diffs.append({"type":"delete", "old_text":old[i].text, "old_page":old[i].page})
        if tag in ("replace","insert"):
            for j in range(j1, j2):
                diffs.append({"type":"insert", "new_text":new[j].text, "new_page":new[j].page})
    # схлопываем delete+insert в replace
    merged: List[Dict] = []
    k = 0
    while k < len(diffs):
        d = diffs[k]
        if d["type"]=="delete" and k+1<len(diffs) and diffs[k+1]["type"]=="insert":
            rep = {
                "type":"replace",
                "old_text": d["old_text"], "old_page": d["old_page"],
                "new_text": diffs[k+1]["new_text"], "new_page": diffs[k+1]["new_page"],
                "parts": token_diff(d["old_text"], diffs[k+1]["new_text"])
            }
            merged.append(rep); k+=2
        else:
            if d["type"]=="delete":
                merged.append({"type":"delete","old_text":d["old_text"],"old_page":d["old_page"],
                               "parts": token_diff(d["old_text"], "")})
            else:
                merged.append({"type":"insert","new_text":d["new_text"],"new_page":d["new_page"],
                               "parts": token_diff("", d["new_text"])})
            k+=1
    return merged

# -------- HTML отчёт --------
def _mark(parts: List[Tuple[str,str]]):
    l=[]; r=[]
    for op,t in parts:
        if not t: continue
        if op=="eq": l.append(t); r.append(t)
        elif op=="del": l.append(f"<del>{t}</del>")
        elif op=="ins": r.append(f"<ins>{t}</ins>")
    return " ".join(l), " ".join(r)

def write_html(diffs: List[Dict], path: str):
    rows=[]
    for d in diffs:
        if d["type"]=="replace":
            left,right=_mark(d["parts"]); lp=d["old_page"]; rp=d["new_page"]
        elif d["type"]=="delete":
            left,right=_mark(d["parts"]); lp=d["old_page"]; rp=""
        else:
            left,right=_mark(d["parts"]); lp=""; rp=d["new_page"]
        rows.append(f"""
        <tr>
          <td><small>p.{lp}</small><div>{left}</div></td>
          <td><small>p.{rp}</small><div>{right}</div></td>
        </tr>""")
    html = """<!doctype html><html lang="en"><head><meta charset="utf-8">
<title>PDF diff (pdfplumber + difflib)</title>
<style>
body{font-family:system-ui,Segoe UI,Roboto,Arial}
table{width:100%;border-collapse:collapse}
td{vertical-align:top;width:50%;border:1px solid #ddd;padding:10px}
del{background:#ffecec;text-decoration:line-through}
ins{background:#eaffea;text-decoration:none}
small{color:#666}
</style></head><body>
<h2>Text diff</h2>
<table>""" + "".join(rows) + "</table></body></html>"
    pathlib.Path(path).write_text(html, encoding="utf-8")

# -------- CLI --------
def main():
    ap = argparse.ArgumentParser(description="Clean PDF text diff (pdfplumber + difflib)")
    ap.add_argument("old_pdf"); ap.add_argument("new_pdf")
    ap.add_argument("--out", default="out")
    args = ap.parse_args()

    outdir = pathlib.Path(args.out); outdir.mkdir(parents=True, exist_ok=True)

    old = extract_pdfplumber(args.old_pdf)
    new = extract_pdfplumber(args.new_pdf)

    diffs = align_sentence_lists(old, new)

    # (опционально) добавить имя файлов в каждый элемент (раскомментируй, если надо)
    # old_name = pathlib.Path(args.old_pdf).name
    # new_name = pathlib.Path(args.new_pdf).name
    # for d in diffs:
    #     d["old_file"] = old_name
    #     d["new_file"] = new_name

    meta = {
        "old_file": pathlib.Path(args.old_pdf).name,
        "new_file": pathlib.Path(args.new_pdf).name,
    }

    (outdir/"diff.json").write_text(
        json.dumps({"meta": meta, "diffs": diffs}, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    write_html(diffs, str(outdir/"diff.html"))

    print(json.dumps({
        "html_report": str(outdir/"diff.html"),
        "json": str(outdir/"diff.json"),
        "stats": {"old_sentences": len(old), "new_sentences": len(new), "diff_items": len(diffs)},
        "meta": meta
    }, ensure_ascii=False))

if __name__ == "__main__":
    main()
