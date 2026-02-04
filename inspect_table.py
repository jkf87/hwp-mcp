# -*- coding: utf-8 -*-
import sys
import os
import json
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
        print("Document opened.")
        text = hwp.get_text()
        print("--- Text Content ---")
        print(text)
        print("--------------------")
    else:
        print(f"Failed to open {path}")
        
    hwp.close_document(save=False)
    hwp.disconnect()
    
except Exception as e:
    print(f"Error: {e}")