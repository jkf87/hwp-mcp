#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Tests for HWP Controller
"""

import pytest
import os
import tempfile
from unittest.mock import patch, MagicMock
from src.tools.hwp_controller import HwpController

class TestHwpController:
    """Test suite for HWP Controller."""

    @patch('win32com.client.Dispatch')
    def test_initialize_hwp(self, mock_dispatch):
        """Test HWP initialization."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp
        
        # Initialize controller
        controller = HwpController()
        
        # Verify initialization
        mock_dispatch.assert_called_once_with("HWPFrame.HwpObject")
        assert controller.hwp is not None

    @patch('win32com.client.Dispatch')
    def test_open_document(self, mock_dispatch):
        """Test opening a document."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp
        
        # Initialize controller
        controller = HwpController()
        
        # Test open document
        result = controller.execute({"type": "open_document", "params": {"path": "test.hwp"}})
        
        # Verify results
        mock_hwp.Open.assert_called_once_with("test.hwp")
        assert result["status"] == "success"
        assert "Document opened" in result["message"]

    @patch('win32com.client.Dispatch')
    def test_save_document(self, mock_dispatch):
        """Test saving a document."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp
        
        # Initialize controller
        controller = HwpController()
        
        # Create a temp file path
        with tempfile.NamedTemporaryFile(suffix='.hwp', delete=False) as temp_file:
            temp_path = temp_file.name
        
        try:
            # Test save document
            result = controller.execute({"type": "save_document", "params": {"path": temp_path}})
            
            # Verify results
            mock_hwp.SaveAs.assert_called_once_with(temp_path)
            assert result["status"] == "success"
            assert "Document saved" in result["message"]
        finally:
            # Clean up
            if os.path.exists(temp_path):
                os.remove(temp_path)

    @patch('win32com.client.Dispatch')
    def test_get_text(self, mock_dispatch):
        """Test getting text from a document."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_hwp.GetTextFile.return_value = "Test document content"
        mock_dispatch.return_value = mock_hwp
        
        # Initialize controller
        controller = HwpController()
        
        # Test get text
        result = controller.execute({"type": "get_text"})
        
        # Verify results
        mock_hwp.GetTextFile.assert_called_once_with("TEXT", "")
        assert result["status"] == "success"
        assert result["data"] == "Test document content"

    @patch('win32com.client.Dispatch')
    def test_insert_text(self, mock_dispatch):
        """Test inserting text into a document."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp
        
        # Initialize controller
        controller = HwpController()
        
        # Test insert text
        result = controller.execute({"type": "insert_text", "params": {"text": "Hello, World!"}})
        
        # Verify results
        mock_hwp.InsertText.assert_called_once_with("Hello, World!")
        assert result["status"] == "success"
        assert "Text inserted successfully" in result["message"] 