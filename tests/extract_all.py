"""Извлечение plain-text из всех .docx в папке «обучение»."""
import os
import sys
import zipfile
import re
from xml.etree import ElementTree as ET

ROOT = os.path.join(os.path.dirname(__file__), '..', 'обучение')
OUT  = os.path.join(os.path.dirname(__file__), 'training-texts')
os.makedirs(OUT, exist_ok=True)

NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def extract_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            with z.open("word/document.xml") as f:
                tree = ET.parse(f)
    except Exception as e:
        return f"<EXTRACTION ERROR: {e}>"
    paragraphs = []
    for p in tree.iter(NS + "p"):
        texts = [t.text or "" for t in p.iter(NS + "t")]
        paragraphs.append("".join(texts))
    return "\n".join(paragraphs)

def safe_name(p):
    rel = os.path.relpath(p, ROOT)
    s = re.sub(r'[\\/:*?"<>|]', '_', rel)
    return s[:200]  # ограничение длины имени

total = 0
for dp, _, files in os.walk(ROOT):
    for f in files:
        if f.lower().endswith('.docx') and not f.startswith('~$'):
            src = os.path.join(dp, f)
            out_name = safe_name(src).replace('.docx', '.txt')
            dst = os.path.join(OUT, out_name)
            try:
                text = extract_docx(src)
                with open(dst, 'w', encoding='utf-8') as w:
                    w.write(text)
                total += 1
            except Exception as e:
                print(f"skip {src}: {e}", file=sys.stderr)

print(f"extracted {total} docs to {OUT}")
