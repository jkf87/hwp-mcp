#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HWP Controller
Provides functionality to interact with Hangul (HWP) documents.
"""

class HwpController:
    """Controller class for interacting with Hangul (HWP) documents."""
    
    def __init__(self):
        """Initialize the HWP controller."""
        self.hwp = None
        self._initialize_hwp()
    
    def _initialize_hwp(self):
        """Initialize the Hangul (HWP) application."""
        try:
            # Import required libraries for HWP interaction
            import win32com.client
            # Create HWP application instance
            self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            # Optional: Set properties or configurations
        except Exception as e:
            raise Exception(f"Failed to initialize HWP: {str(e)}")
    
    def execute(self, command):
        """
        Execute a command on the HWP document.
        
        Args:
            command (dict): The command to execute with its parameters
            
        Returns:
            dict: Result of the command execution
        """
        if not self.hwp:
            raise Exception("HWP application not initialized")
            
        cmd_type = command.get("type")
        params = command.get("params", {})
        
        # Handle different command types
        if cmd_type == "open_document":
            return self._open_document(params.get("path"))
        elif cmd_type == "save_document":
            return self._save_document(params.get("path"))
        elif cmd_type == "get_text":
            return self._get_text()
        elif cmd_type == "insert_text":
            return self._insert_text(params.get("text"))
        else:
            return {"status": "error", "message": f"Unknown command type: {cmd_type}"}
    
    def _open_document(self, path):
        """Open an HWP document."""
        try:
            self.hwp.Open(path)
            return {"status": "success", "message": f"Document opened: {path}"}
        except Exception as e:
            return {"status": "error", "message": f"Failed to open document: {str(e)}"}
    
    def _save_document(self, path):
        """Save the current HWP document."""
        try:
            self.hwp.SaveAs(path)
            return {"status": "success", "message": f"Document saved: {path}"}
        except Exception as e:
            return {"status": "error", "message": f"Failed to save document: {str(e)}"}
    
    def _get_text(self):
        """Get text from the current HWP document."""
        try:
            text = self.hwp.GetTextFile("TEXT", "")
            return {"status": "success", "data": text}
        except Exception as e:
            return {"status": "error", "message": f"Failed to get text: {str(e)}"}
    
    def _insert_text(self, text):
        """Insert text into the current HWP document."""
        try:
            self.hwp.InsertText(text)
            return {"status": "success", "message": "Text inserted successfully"}
        except Exception as e:
            return {"status": "error", "message": f"Failed to insert text: {str(e)}"} 