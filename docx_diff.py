#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Кроссплатформенный DOCX diff (side-by-side):
- Извлекает текст из .docx через Mammoth (в Markdown-подобный плоский текст)
- Нормализует переносы/пробелы
- Делит на "предложения" (простая эвристика)
- Выравнивает списки предложений с difflib.SequenceMatcher
- Рисует HTML-отчёт: слева "старое" (с <del>), справа "новое" (с <ins>)
"""

import argparse, pathlib, re, json
from dataclasses import dataclass
from typing import List, Dict, Tuple
from difflib import SequenceMatcher
import mammoth

# ---------- Нормализация ----------
WS = re.compile(r"\s+")
SENT_SPLIT = re.compile(r'(?<=[\.\!\?])\s+(?=[A-ZА-Я0-9])', re.UNICODE)

def normalize(s: str) -> str:
    if not s:
        return ""
    s = s.replace("\xa0"," ").replace("\r","")
    s = WS.sub(" ", s).strip()
    return s

def to_sentences(txt: str) -> List[str]:
    txt = normalize(txt)
    if not txt:
        return []
    if SENT_SPLIT.search(txt):
        return [p.strip() for p in SENT_SPLIT.split(txt) if p.strip()]
    return [txt]

@dataclass
class Sentence:
    text: str
    # (страниц в DOCX нет в чистом виде; оставляем только текст)

# ---------- Извлечение текста из DOCX ----------
def docx_to_text(path: str) -> str:
    """Берём Markdown из Mammoth и приводим к плоскому тексту."""
    with open(path, "rb") as f:
        md = mammoth.convert_to_markdown(f).value
    # уберём markdown-разметку по минимуму (**, __, #, *, >) — чтобы difflib не путался
    md = re.sub(r"[*_`>#~-]+", " ", md)
    return normalize(md)

def extract_sentences(path: str) -> List[Sentence]:
    txt = docx_to_text(path)
    return [Sentence(s) for s in to_sentences(txt)]

# ---------- Diff utils ----------
def token_diff(a: str, b: str) -> List[Tuple[str,str]]:
    """Покомпонентный diff по словам для красивой подсветки."""
    A, B = a.split(), b.split()
    sm2 = SequenceMatcher(a=A, b=B, autojunk=False)
    parts: List[Tuple[str,str]] = []
    for tag,i1,i2,j1,j2 in sm2.get_opcodes():
        if tag == "equal":
            parts.append(("eq"," ".join(A[i1:i2])))
        elif tag == "delete":
            parts.append(("del"," ".join(A[i1:i2])))
        elif tag == "insert":
            parts.append(("ins"," ".join(B[j1:j2])))
        elif tag == "replace":
            if i1<i2: parts.append(("del"," ".join(A[i1:i2])))
            if j1<j2: parts.append(("ins"," ".join(B[j1:j2])))
    return parts

def align_sentence_lists(old: List[Sentence], new: List[Sentence]) -> List[Dict]:
    sm = SequenceMatcher(a=[s.text for s in old], b=[s.text for s in new], autojunk=False)
    diffs: List[Dict] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        if tag in ("replace","delete"):
            for i in range(i1, i2):
                diffs.append({"type":"delete","old_text":old[i].text})
        if tag in ("replace","insert"):
            for j in range(j1, j2):
                diffs.append({"type":"insert","new_text":new[j].text})
    # Схлопнем попарно delete+insert в replace
    merged: List[Dict] = []
    k = 0
    while k < len(diffs):
        d = diffs[k]
        if d["type"]=="delete" and k+1<len(diffs) and diffs[k+1]["type"]=="insert":
            rep = {
                "type":"replace",
                "old_text": d["old_text"],
                "new_text": diffs[k+1]["new_text"],
                "parts": token_diff(d["old_text"], diffs[k+1]["new_text"])
            }
            merged.append(rep); k += 2
        else:
            if d["type"]=="delete":
                merged.append({"type":"delete","old_text":d["old_text"],"parts":token_diff(d["old_text"], "")})
            else:
                merged.append({"type":"insert","new_text":d["new_text"],"parts":token_diff("", d["new_text"])})
            k += 1
    return merged

# ---------- HTML отчёт ----------
def _mark(parts: List[Tuple[str,str]]) -> Tuple[str,str]:
    left, right = [], []
    for op, t in parts:
        if not t:
            continue
        if op == "eq":
            left.append(t); right.append(t)
        elif op == "del":
            left.append(f"<del>{t}</del>")
        elif op == "ins":
            right.append(f"<ins>{t}</ins>")
    return " ".join(left), " ".join(right)

def write_html(diffs: List[Dict], path: str):
    rows = []
    for d in diffs:
        if d["type"]=="replace":
            l, r = _mark(d["parts"])
        elif d["type"]=="delete":
            l, r = _mark(d["parts"])
        else:
            l, r = _mark(d["parts"])
        rows.append(f"""
        <tr>
          <td><div>{l}</div></td>
          <td><div>{r}</div></td>
        </tr>""")

    html = """<!doctype html><html lang="ru"><head><meta charset="utf-8">
<title>DOCX diff (side-by-side)</title>
<style>
body{font-family:system-ui,Segoe UI,Roboto,Arial}
table{width:100%;border-collapse:collapse}
td{vertical-align:top;width:50%;border:1px solid #ddd;padding:10px}
del{background:#ffecec;text-decoration:line-through}
ins{background:#eaffea;text-decoration:none}
</style></head><body>
<h2>Text diff</h2>
<table>""" + "".join(rows) + "</table></body></html>"
    pathlib.Path(path).write_text(html, encoding="utf-8")

# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="DOCX side-by-side text diff (Mammoth + difflib)")
    ap.add_argument("old_docx"); ap.add_argument("new_docx")
    ap.add_argument("--out", default="out_docx_diff")
    args = ap.parse_args()

    out = pathlib.Path(args.out); out.mkdir(parents=True, exist_ok=True)

    old_sents = extract_sentences(args.old_docx)
    new_sents = extract_sentences(args.new_docx)

    diffs = align_sentence_lists(old_sents, new_sents)

    # JSON на всякий случай
    (out/"diff.json").write_text(json.dumps(diffs, ensure_ascii=False, indent=2), encoding="utf-8")
    write_html(diffs, str(out/"diff_docx.html"))

    print(json.dumps({
        "html_report": str(out/"diff_docx.html"),
        "json": str(out/"diff.json"),
        "stats": {"old_sentences": len(old_sents), "new_sentences": len(new_sents), "diff_items": len(diffs)}
    }, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
