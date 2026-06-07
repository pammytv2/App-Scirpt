
import re

with open('PdfTemplate.html', 'r', encoding='utf-8') as f:
    html = f.read()

with open('img/tt1_base64.txt', 'r', encoding='utf-8') as f:
    tt1 = f.read().strip()

with open('img/tt2_base64.txt', 'r', encoding='utf-8') as f:
    tt2 = f.read().strip()

matches = re.findall(r'<img src="data:image/png;base64,([^"]+)"', html)

for i, m in enumerate(matches):
    if m == tt1:
        print(f"Image {i} is tt1")
    elif m == tt2:
        print(f"Image {i} is tt2")
    else:
        print(f"Image {i} is something else (starts with {m[:20]})")
