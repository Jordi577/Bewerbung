from pathlib import Path
from zipfile import ZipFile
from pdf2image import convert_from_path
from pypdf import PdfReader
import hashlib
import re

p = Path(r'C:\Users\Jordi Buscaglia\Documents\Bewerbung')
ims = convert_from_path(str(p/'Lebenslauf_Jordi_Buscaglia_Design_2.pdf'), dpi=150,
    poppler_path=r'C:\Users\Jordi Buscaglia\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\poppler\Library\bin')
for i, im in enumerate(ims):
    im.save(p/'.gestaltung'/'render2'/f'page-{i+1}.png')
text = ' '.join(' '.join(x.extract_text().split()) for x in PdfReader(p/'Lebenslauf_Jordi_Buscaglia_Design_2.pdf').pages)
tex = (p/'cv.tex').read_text(encoding='utf-8')
items = re.findall(r'\\item (.+)', tex)
items = [x.replace(r'\glqq ', '„').replace(r'\grqq', '“') for x in items]
assert len(items) == 9
compact=''.join(text.split())
assert all(''.join(x.split()) in compact for x in items), [x for x in items if ''.join(x.split()) not in compact]
assert text.count('Forschung zu Retrieval-Augmented Generation') == 3
assert text.count('Machbarkeitsstudien zur Codegenerierung') == 3
assert len(ims) == 1
for expected in ['Jun. 2026 - heute', 'Sep. 2025 - Jun. 2026', 'Jan. 2025 - Sep. 2025',
                 'Okt. 2024 - Jan. 2025', 'Jan. 2024 - Sep. 2024', '2025 - heute',
                 '2020 - 2025', '2011 - 2019', 'Git, GitHub, Codex', 'Python, C',
                 'Verhandlungssicher', 'Grundkenntnisse', 'Muttersprache']:
    assert expected in text, expected
with ZipFile(p/'Lebenslauf_Jordi_Buscaglia_Design_2.docx') as z:
    media = [n for n in z.namelist() if n.startswith('word/media/')]
    assert any(hashlib.sha256(z.read(n)).digest() == hashlib.sha256((p/'foto.png').read_bytes()).digest() for n in media)
print('Alle 9 Tätigkeitsbeschreibungen exakt vorhanden. Zeiträume und Kenntnisse geprüft. Foto original eingebettet. PDF: 1 Seite.')
for name in ['Lebenslauf_Jordi_Buscaglia_Design_2.docx','Lebenslauf_Jordi_Buscaglia_Design_2.pdf']:
    print(name, (p/name).stat().st_size)

