import xml.etree.ElementTree as ET
import re

tree = ET.parse('scratch/user_screen.xml')
root = tree.getroot()
safe_seller = '필오넷'

def parse_bounds(b_str):
    m = re.match(r'\[(\d+),(\d+)\]\[(\d+),(\d+)\]', b_str)
    if m:
        x1, y1, x2, y2 = map(int, m.groups())
        return x1, y1, x2, y2, (x1+x2)//2, (y1+y2)//2
    return 0, 0, 0, 0, 0, 0

print("=== Checking elements ===")
for elem in root.iter():
    rid = elem.attrib.get('resource-id', '')
    cd = elem.attrib.get('content-desc', '')
    txt = elem.attrib.get('text', '')
    b = elem.attrib.get('bounds', '')
    if 'basic_product_card_information' in rid or '필오넷' in cd or '필오넷' in txt:
        x1, y1, x2, y2, cx, cy = parse_bounds(b)
        print(f"Tag={elem.tag.split('.')[-1]}, rid={rid[:30]}, cd={cd[:20]}, text={txt[:20]}, bounds={b}, center=({cx}, {cy})")
