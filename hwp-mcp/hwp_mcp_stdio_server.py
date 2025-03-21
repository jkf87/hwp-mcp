#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HWP MCP STDIO Server
This is the main entry point for the HWP MCP server that communicates via standard I/O.
"""

import sys
import json
from src.tools.hwp_controller import HwpController
from src.utils.command_parser import CommandParser

def main():
    """Main function to run the HWP MCP STDIO server."""
    hwp_controller = HwpController()
    command_parser = CommandParser()
    
    try:
        while True:
            line = sys.stdin.readline().strip()
            if not line:
                continue
                
            try:
                command = command_parser.parse(line)
                result = hwp_controller.execute(command)
                sys.stdout.write(json.dumps(result, ensure_ascii=False) + '\n')
                sys.stdout.flush()
            except Exception as e:
                error_response = {"status": "error", "message": str(e)}
                sys.stdout.write(json.dumps(error_response, ensure_ascii=False) + '\n')
                sys.stdout.flush()
    except KeyboardInterrupt:
        print("Server terminated by user", file=sys.stderr)
        sys.exit(0)

if __name__ == "__main__":
    main() 