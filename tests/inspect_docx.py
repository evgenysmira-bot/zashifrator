"""Показывает сырую структуру docx около проблемного места."""
import zipfile, sys
from xml.etree import ElementTree as ET

path = sys.argv[1]
needle = sys.argv[2] if len(sys.argv) > 2 else "Реквизиты Заказчика"
NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

with zipfile.ZipFile(path) as z:
    with z.open("word/document.xml") as f:
        tree = ET.parse(f)

root = tree.getroot()
paragraphs = list(root.iter(NS + "p"))
for i, p in enumerate(paragraphs):
    text = "".join(t.text or "" for t in p.iter(NS + "t"))
    if needle in text:
        print(f"=== paragraph #{i}: {text[:200]}")
        # Также покажем несколько соседних
        for j in range(i, min(i + 8, len(paragraphs))):
            q = paragraphs[j]
            # Структура: runs, sdt, hyperlink и т.д.
            qtext = "".join(t.text or "" for t in q.iter(NS + "t"))
            # Считаем элементы верхнего уровня
            children = [c.tag.replace(NS, '') for c in q]
            sdt_count = sum(1 for c in q.iter(NS + "sdt"))
            print(f"  p#{j}  sdts={sdt_count}  children={children[:10]}")
            print(f"       text=[{qtext[:180]}]")
        print()
        break
