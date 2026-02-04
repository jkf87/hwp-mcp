# -*- coding: utf-8 -*-
import sys
import os
import json

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__))))

from src.tools.hwp_template_engine import HwpTemplateEngine

try:
    engine = HwpTemplateEngine(visible=False)
    # Use absolute path for safety
    path = os.path.abspath("회의록 양식.hwp")
    if engine.open(path):
        fields = engine.get_field_list()
        print(json.dumps(fields, ensure_ascii=False, indent=2))
    else:
        print(f"Failed to open {path}")
    engine.quit()
except Exception as e:
    print(f"Error: {e}")
