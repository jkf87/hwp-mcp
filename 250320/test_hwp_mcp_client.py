#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import time
from mcp.client.stdio import stdio_client

def log(message):
    """Print a message to stderr."""
    print(message, file=sys.stderr)

def main():
    """Test the HWP MCP server using the MCP stdio client."""
    log("Starting HWP MCP client test...")
    
    try:
        # Connect to the MCP server (assuming it's running)
        client = stdio_client.connect("hwp")
        log("Connected to HWP MCP server")
        
        # Test creating a new document
        log("Testing: Create new document")
        result = client.hwp_create()
        log(f"Result: {result}")
        time.sleep(1)
        
        # Test inserting text
        log("Testing: Insert text")
        result = client.hwp_insert_text(text="안녕하세요! HWP MCP 테스트입니다.")
        log(f"Result: {result}")
        time.sleep(1)
        
        # Test setting font
        log("Testing: Set font")
        result = client.hwp_set_font(name="맑은 고딕", size=12, bold=True)
        log(f"Result: {result}")
        time.sleep(1)
        
        # Test inserting paragraph
        log("Testing: Insert paragraph")
        result = client.hwp_insert_paragraph()
        log(f"Result: {result}")
        time.sleep(1)
        
        # Test inserting table
        log("Testing: Insert table")
        result = client.hwp_insert_table(rows=3, cols=3)
        log(f"Result: {result}")
        time.sleep(1)
        
        # Test getting text content
        log("Testing: Get text content")
        result = client.hwp_get_text()
        log(f"Result: {result[:100]}...")  # Show just the first 100 chars
        time.sleep(1)
        
        # Test saving document
        log("Testing: Save document")
        test_file_path = os.path.join(os.getcwd(), "hwp_mcp_test_result.hwp")
        result = client.hwp_save(path=test_file_path)
        log(f"Result: {result}")
        
        log("All tests completed successfully!")
        log(f"Test document saved to: {test_file_path}")
        
        # Keep the document open - manually close later
        log("Test completed. The HWP document will remain open.")
        
    except Exception as e:
        log(f"Error during test: {str(e)}")
        import traceback
        log(traceback.format_exc())
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 