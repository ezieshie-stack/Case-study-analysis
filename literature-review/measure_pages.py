import re, html
from reportlab.pdfbase.pdfmetrics import stringWidth

PAGE_W, PAGE_H = 595.276, 841.89
M = 2.5 * 28.3465
TEXT_W, TEXT_H = PAGE_W - 2*M, PAGE_H - 2*M
LEAD = 18.0
LPP = TEXT_H / LEAD
TW = 20.0  # twips per point

import sys, zipfile
_p = sys.argv[1]
if _p.endswith('.docx'):
    x = zipfile.ZipFile(_p).read('word/document.xml').decode('utf-8')
else:
    x = open(_p, encoding='utf-8').read()
paras = re.findall(r'<w:p\b.*?</w:p>', x, re.S)

def attr(p, tag, name):
    m = re.search(r'<w:%s\b[^>]*w:%s="(-?\d+)"' % (tag, name), p)
    return int(m.group(1))/TW if m else 0.0

total = 0.0
for p in paras:
    txt = html.unescape(' '.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p, re.S)))
    font = 'Times-Bold' if '<w:b/>' in p else 'Times-Roman'
    before, after = attr(p,'spacing','before'), attr(p,'spacing','after')
    left    = attr(p,'ind','left')
    firstLn = attr(p,'ind','firstLine')
    hanging = attr(p,'ind','hanging')
    if not txt.strip():
        total += 1 + (before+after)/LEAD
        continue
    # width available on line 1 vs subsequent lines
    w1 = TEXT_W - left - firstLn + hanging
    wn = TEXT_W - left
    words, line, n = txt.split(), '', 1
    for w in words:
        trial = (line + ' ' + w).strip()
        if stringWidth(trial, font, 12) <= (w1 if n == 1 else wn):
            line = trial
        else:
            n += 1; line = w
    total += n + (before+after)/LEAD

print(f'lines/page : {LPP:.1f}')
print(f'total lines: {total:.1f}')
print(f'PAGES      : {total/LPP:.2f}')
