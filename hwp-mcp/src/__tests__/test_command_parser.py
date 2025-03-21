#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Tests for Command Parser
"""

import pytest
from src.utils.command_parser import CommandParser

def test_command_parser_valid_json():
    """Test parsing valid JSON commands."""
    parser = CommandParser()
    
    # Test basic command
    command = parser.parse('{"type": "open_document", "params": {"path": "test.hwp"}}')
    assert command["type"] == "open_document"
    assert command["params"]["path"] == "test.hwp"
    
    # Test command without params
    command = parser.parse('{"type": "get_text"}')
    assert command["type"] == "get_text"
    assert "params" not in command

def test_command_parser_invalid_json():
    """Test parsing invalid JSON commands."""
    parser = CommandParser()
    
    # Test invalid JSON
    with pytest.raises(ValueError) as e:
        parser.parse('{type: "open_document"}')
    assert "Invalid JSON format" in str(e.value)
    
    # Test non-dict JSON
    with pytest.raises(ValueError) as e:
        parser.parse('["open_document"]')
    assert "Command must be a JSON object" in str(e.value)
    
    # Test missing type
    with pytest.raises(ValueError) as e:
        parser.parse('{"params": {"path": "test.hwp"}}')
    assert "Command must have a 'type' field" in str(e.value)
    
    # Test invalid params
    with pytest.raises(ValueError) as e:
        parser.parse('{"type": "open_document", "params": "test.hwp"}')
    assert "Command params must be a JSON object" in str(e.value) 