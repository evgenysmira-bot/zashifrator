import zipfile, sys
from xml.etree import ElementTree as ET

path = sys.argv[1]
ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
with zipfile.ZipFile(path) as z:
    with z.open("word/document.xml") as f:
        tree = ET.parse(f)
paragraphs = []
for p in tree.iter(ns + "p"):
    texts = [t.text or "" for t in p.iter(ns + "t")]
    paragraphs.append("".join(texts))
text = "\n".join(paragraphs)
sys.stdout.reconfigure(encoding="utf-8")
print(text)
