import xml.etree.ElementTree as ET
import re

with open("배송지 관리.xml", "r", encoding="utf-8") as f:
    source = f.read()

def test_parse(source):
    root = ET.fromstring(source)

    def parse_bounds(b_str):
        m = re.match(r"\[(\d+),(\d+)\]\[(\d+),(\d+)\]", b_str or "")
        if m:
            return int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
        return None

    default_y_ranges = []
    for node in root.iter():
        txt = (node.attrib.get("text") or "") + " " + (node.attrib.get("content-desc") or "")
        if "기본배송지" in txt:
            b = parse_bounds(node.attrib.get("bounds"))
            if b:
                default_y_ranges.append((b[1] - 250, b[3] + 350))

    non_default_btns = []
    cards_found = False

    for node in root.iter():
        if node.attrib.get("is-collection-item") == "true":
            del_btns = [e for e in node.iter() if (e.attrib.get("text") == "삭제" or e.attrib.get("content-desc") == "삭제")]
            if not del_btns:
                continue
            cards_found = True
            
            texts = [(e.attrib.get("text") or "") + " " + (e.attrib.get("content-desc") or "") for e in node.iter()]
            card_text = " ".join(texts)
            card_bounds = parse_bounds(node.attrib.get("bounds"))
            if card_bounds:
                c_x1, c_y1, c_x2, c_y2 = card_bounds
            else:
                c_y1, c_y2 = -1, -1

            if "기본배송지" in card_text:
                print(f"  [DEFAULT CARD] (bounds={node.attrib.get('bounds')}) -> excluded")
                if c_y1 != -1:
                    default_y_ranges.append((c_y1 - 10, c_y2 + 10))
                continue

            for btn in del_btns:
                b = parse_bounds(btn.attrib.get("bounds"))
                if not b:
                    continue
                bx1, by1, bx2, by2 = b
                cx, cy = (bx1 + bx2) // 2, (by1 + by2) // 2
                if 250 <= cy <= 1900:
                    non_default_btns.append((cy, cx))

    if cards_found:
        non_default_btns.sort()
        return [(cx, cy) for cy, cx in non_default_btns]

    all_delete_btns = []
    for node in root.iter():
        if node.attrib.get("text") == "삭제" or node.attrib.get("content-desc") == "삭제":
            b = parse_bounds(node.attrib.get("bounds"))
            if not b:
                continue
            bx1, by1, bx2, by2 = b
            cy = (by1 + by2) // 2
            cx = (bx1 + bx2) // 2
            if not (250 <= cy <= 1900):
                continue
            is_in_default = any(y1 <= cy <= y2 for y1, y2 in default_y_ranges)
            if is_in_default:
                print(f"  [DEFAULT Y] y={cy} inside default range -> excluded")
                continue
            all_delete_btns.append((cy, cx))

    all_delete_btns.sort()
    return [(cx, cy) for cy, cx in all_delete_btns]

result = test_parse(source)
print("Result targets:", result)
