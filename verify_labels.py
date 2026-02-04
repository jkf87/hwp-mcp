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
        print("Document opened.")
        
        # Tuple of (Korean Label, English Name)
        labels_to_check = [
            ("회의명", "Meeting Title"), 
            ("일시", "Date/Time"), 
            ("장소", "Place"), 
            ("참석자", "Attendees"), 
            ("내용", "Content"), 
            ("주요 내용", "Main Content"), 
            ("주요내용", "MainContent_NoSpace"),
            ("안건", "Agenda"), 
            ("결과", "Result")
        ]
        
        for ko_label, en_name in labels_to_check:
            if hwp.find_text(ko_label):
                print(f"FOUND: {en_name} ({ko_label})")
            else:
                print(f"NOT FOUND: {en_name}")
            
    else:
        print(f"Failed to open {path}")
        
    hwp.close_document(save=False)
    hwp.disconnect()
    
except Exception as e:
    print(f"Error: {e}")