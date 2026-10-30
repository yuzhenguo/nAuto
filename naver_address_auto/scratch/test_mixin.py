import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from unittest.mock import MagicMock
from naver_pay_flow import PayAddressMixin
import xml.etree.ElementTree as ET

class DummyWorker(PayAddressMixin):
    def __init__(self, xml_source):
        self.driver = MagicMock()
        self.driver.page_source = xml_source
        self.device_id = "test_device"

    def _log(self, msg):
        try:
            print(msg)
        except Exception:
            pass

with open("배송지 관리.xml", "r", encoding="utf-8") as f:
    xml_content = f.read()

# 1. Test full XML
worker = DummyWorker(xml_content)
targets = worker._pay_parse_delete_targets_from_xml()
print("Parsed targets count:", len(targets) if targets is not None else None)
assert targets == [(798, 969), (798, 1338)], f"Targets mismatch: {targets}"
first_target = worker._pay_find_non_default_delete_btn()
assert first_target == (798, 969), f"First target mismatch: {first_target}"
print("Test 1 Passed: Only non-default addresses are targeted!")

# 2. Test when only the default address exists in the list
root = ET.fromstring(xml_content)
list_view = None
for node in root.iter():
    if node.attrib.get("class") == "android.widget.ListView" and node.attrib.get("bounds") == "[60,372][1020,1599]":
        list_view = node
        break

if list_view is not None:
    # Keep only child 0 (Card 0, default), remove Card 1 and Card 2
    items = [c for c in list_view if c.attrib.get("is-collection-item") == "true"]
    for it in items[1:]:
        list_view.remove(it)
    only_default_xml = ET.tostring(root, encoding="utf-8").decode("utf-8")
    
    worker_default_only = DummyWorker(only_default_xml)
    targets_default_only = worker_default_only._pay_parse_delete_targets_from_xml()
    print("When only default address exists -> targets:", targets_default_only)
    assert targets_default_only == [], f"Expected empty list [], got {targets_default_only}"
    first_default_only = worker_default_only._pay_find_non_default_delete_btn()
    assert first_default_only is None, f"Expected None, got {first_default_only}"
    print("Test 2 Passed: When only default address remains, deletion correctly returns None and finishes!")
