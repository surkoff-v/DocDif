# pdf_text_diff.py
#
# pip install pymupdf rapidfuzz regex beautifulsoup4 
# при необходимости pytesseract, pdf2image
#


import sys, json, re, pathlib
from dataclasses import dataclass
import fitz  # PyMuPDF
from difflib import SequenceMatcher

SENT_SPLIT = re.compile(r'(?<=[\.\!\?])\s+(?=[A-ZА-Я0-9])')
HYPHEN_JOIN = re.compile(r'(\w)-\n(\w)')  # изъятие переносов
WS = re.compile(r'\s+')

@dataclass
class Sentence:
    text: str
    page: int
    section_path: str  # например "2.1>Требования>SLA" (можно заполнять позже)

def normalize(txt: str) -> str:
    txt = txt.replace('\r', '')
    txt = HYPHEN_JOIN.sub(r'\1\2', txt)
    txt = txt.replace('\n', ' ')
    txt = WS.sub(' ', txt).strip()
    return txt

def extract_sentences(pdf_path: str) -> list[Sentence]:
    doc = fitz.open(pdf_path)
    result = []
    for pno in range(len(doc)):
        page = doc[pno]
        # Берём текст странично; при желании можно пройтись по блокам.
        txt = page.get_text()
        txt = normalize(txt)
        if not txt:
            continue
        for s in SENT_SPLIT.split(txt):
            s = s.strip()
            if s:
                result.append(Sentence(s, pno + 1, ""))  # страницы 1-базные
    return result

def token_diff(a: str, b: str):
    a_tok = a.split()
    b_tok = b.split()
    sm = SequenceMatcher(a=a_tok, b=b_tok, autojunk=False)
    chunks = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        chunks.append((tag, ' '.join(a_tok[i1:i2]), ' '.join(b_tok[j1:j2])))
    return chunks

def sentence_align(old: list[Sentence], new: list[Sentence]):
    # Базовое выравнивание по предложениям (позиционный SequenceMatcher).
    sm = SequenceMatcher(a=[s.text for s in old], b=[s.text for s in new], autojunk=False)
    diffs = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            continue
        if tag in ('replace', 'delete'):
            for i in range(i1, i2):
                diffs.append({
                    'type': 'delete',
                    'old_text': old[i].text,
                    'old_page': old[i].page,
                    'new_text': '',
                    'new_page': None,
                    'tokens': token_diff(old[i].text, new[j1].text) if tag == 'replace' and j1 < j2 else []
                })
        if tag in ('replace', 'insert'):
            for j in range(j1, j2):
                diffs.append({
                    'type': 'insert',
                    'old_text': '',
                    'old_page': None,
                    'new_text': new[j].text,
                    'new_page': new[j].page,
                    'tokens': token_diff(old[i1].text, new[j].text) if tag == 'replace' and i1 < i2 else []
                })
    return diffs

def html_report(diffs: list[dict], out_html: str):
    def mark_tokens(tokens):
        parts_old, parts_new = [], []
        for tag, a, b in tokens:
            if tag == 'equal':
                parts_old.append(a); parts_new.append(b)
            elif tag in ('delete', 'replace'):
                if a: parts_old.append(f'<del>{a}</del>')
            if tag in ('insert', 'replace'):
                if b: parts_new.append(f'<ins>{b}</ins>')
        return ' '.join(parts_old), ' '.join(parts_new)

    rows = []
    for d in diffs:
        if d['type'] == 'delete' and d['tokens']:
            left, right = mark_tokens(d['tokens'])
        elif d['type'] == 'insert' and d['tokens']:
            left, right = mark_tokens(d['tokens'])
        else:
            left, right = (d['old_text'] or ''), (d['new_text'] or '')
        rows.append(f"""
        <tr>
          <td><small>p.{d.get('old_page') or ''}</small><div>{left}</div></td>
          <td><small>p.{d.get('new_page') or ''}</small><div>{right}</div></td>
        </tr>""")
    html = f"""
    <html><head><meta charset="utf-8">
    <style>
    body{{font-family:system-ui,Segoe UI,Roboto,Arial;}}
    table{{width:100%;border-collapse:collapse}}
    td{{vertical-align:top;width:50%;border:1px solid #ddd;padding:8px}}
    del{{background:#ffecec;text-decoration:line-through}}
    ins{{background:#eaffea;text-decoration:none}}
    small{{color:#666}}
    </style></head><body>
    <h2>Text diff</h2>
    <table>{''.join(rows)}</table>
    </body></html>"""
    pathlib.Path(out_html).write_text(html, encoding='utf-8')

def annotate_pdf(pdf_in: str, phrases: list[tuple[int, str]], pdf_out: str):
    # phrases: list of (page_number, phrase) для подсветки
    doc = fitz.open(pdf_in)
    for pno, phrase in phrases:
        try:
            page = doc[pno - 1]
        except: 
            continue
        for rect in page.search_for(phrase, hit_max=16):
            page.add_highlight_annot(rect)
    doc.save(pdf_out, incremental=False)

def main(old_pdf: str, new_pdf: str, out_dir: str):
    out = pathlib.Path(out_dir); out.mkdir(parents=True, exist_ok=True)
    old_sents = extract_sentences(old_pdf)
    new_sents = extract_sentences(new_pdf)
    diffs = sentence_align(old_sents, new_sents)

    # Сохраняем JSON
    json_path = out / "diff.json"
    json_path.write_text(json.dumps(diffs, ensure_ascii=False, indent=2), encoding='utf-8')

    # HTML отчёт
    html_report(diffs, str(out / "diff.html"))

    # Подсветки: удалённые → в старом, добавленные → в новом
    del_phr = [(d['old_page'], d['old_text']) for d in diffs if d['type']=='delete' and d['old_text']]
    ins_phr = [(d['new_page'], d['new_text']) for d in diffs if d['type']=='insert' and d['new_text']]
    annotate_pdf(old_pdf, del_phr[:200], str(out / "old_annotated.pdf"))
    annotate_pdf(new_pdf, ins_phr[:200], str(out / "new_annotated.pdf"))

    print(json.dumps({
        "html_report": str(out / "diff.html"),
        "json": str(json_path),
        "pdf_old_annot": str(out / "old_annotated.pdf"),
        "pdf_new_annot": str(out / "new_annotated.pdf")
    }, ensure_ascii=False))

if __name__ == "__main__":
    old_pdf, new_pdf, out_dir = sys.argv[1], sys.argv[2], sys.argv[3]
    main(old_pdf, new_pdf, out_dir)
