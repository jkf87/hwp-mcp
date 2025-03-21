#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Command Parser
Parses commands received from standard input.
"""

import json

class CommandParser:
    """Parser for HWP MCP commands."""
    
    def __init__(self):
        """Initialize the command parser."""
        pass
    
    def parse(self, command_string):
        """
        Parse a command string into a structured command object.
        
        Args:
            command_string (str): The command string to parse
            
        Returns:
            dict: A structured command object
            
        Raises:
            ValueError: If the command string is invalid
        """
        try:
            # Parse JSON command
            command = json.loads(command_string)
            
            # Validate command structure
            if not isinstance(command, dict):
                raise ValueError("Command must be a JSON object")
                
            if "type" not in command:
                raise ValueError("Command must have a 'type' field")
                
            # Validate command parameters
            if "params" in command and not isinstance(command["params"], dict):
                raise ValueError("Command params must be a JSON object")
                
            return command
            
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON format: {str(e)}")
        except Exception as e:
            raise ValueError(f"Error parsing command: {str(e)}") 