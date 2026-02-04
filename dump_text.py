# -*- coding: utf-8 -*-
import sys
import os
import logging

# Add project root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__))))

from src.tools.hwp_controller import HwpController

# Configure logging
logging.basicConfig(level=logging.INFO)

try:
    hwp = HwpController()
    path = os.path.abspath("회의록 양식.hwp")
    
    if hwp.open_document(path):
        text = hwp.get_text()
        with open("dump.txt", "w", encoding="utf-8") as f:
            f.write(text)
        print("Text dumped to dump.txt")
    else:
        print(f"Failed to open {path}")
        
    hwp.close_document(save=False)
    hwp.disconnect()
    
except Exception as e:
    print(f"Error: {e}")
