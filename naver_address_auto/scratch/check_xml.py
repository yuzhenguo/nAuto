import re

with open("배송지 관리.xml", "r", encoding="utf-8") as f:
    text = f.read()

tags = []
for m in re.finditer(r'<(/)?([a-zA-Z0-9\._]+)[^>]*?(/)?>', text):
    is_close = m.group(1) == '/'
    tag_name = m.group(2)
    is_self = m.group(3) == '/'
    if is_self:
        continue
    if not is_close:
        tags.append((tag_name, m.start()))
    else:
        if tags and tags[-1][0] == tag_name:
            tags.pop()
        else:
            line_no = text[:m.start()].count('\n') + 1
            exp = tags[-1][0] if tags else "NONE"
            opened_line = text[:tags[-1][1]].count('\n') + 1 if tags else 0
            print(f"Mismatch at line {line_no}: closing </{tag_name}>, expected </{exp}> (opened at line {opened_line})")
            break
else:
    print("All tags matched!")
