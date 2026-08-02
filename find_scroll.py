import sys
with open(r'e:\네이버자동주문\naver_order_worker.py', encoding='utf-8', errors='replace') as f:
    lines = f.readlines()
for i, l in enumerate(lines):
    if 'def _scroll' in l or 'def scroll' in l:
        print(f'{i+1}: {l.strip()}')
