#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import traceback
import logging
import ssl
from threading import Thread

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    filename="hwp_mcp_stdio_server.log",
    filemode="a",
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)

# 추가 스트림 핸들러 설정 (별도로 추가)
stderr_handler = logging.StreamHandler(sys.stderr)
stderr_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
logger = logging.getLogger("hwp-mcp-stdio-server")
logger.addHandler(stderr_handler)

# Optional: Disable SSL certificate validation for development
ssl._create_default_https_context = ssl._create_unverified_context

# Set up paths
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

try:
    # Import FastMCP library
    from mcp.server.fastmcp import FastMCP
    logger.info("FastMCP successfully imported")
except ImportError as e:
    logger.error(f"Failed to import FastMCP: {str(e)}")
    print(f"Error: Failed to import FastMCP. Please install with 'pip install mcp'", file=sys.stderr)
    sys.exit(1)

# Try to import HwpController
try:
    from src.tools.hwp_controller import HwpController
    logger.info("HwpController imported successfully")
except ImportError as e:
    logger.error(f"Failed to import HwpController: {str(e)}")
    # Try alternate paths
    try:
        sys.path.append(os.path.join(current_dir, "src"))
        sys.path.append(os.path.join(current_dir, "src", "tools"))
        from hwp_controller import HwpController
        logger.info("HwpController imported from alternate path")
    except ImportError as e2:
        logger.error(f"Could not find HwpController in any path: {str(e2)}")
        print(f"Error: Could not find HwpController module", file=sys.stderr)
        sys.exit(1)

# Initialize FastMCP server
mcp = FastMCP(
    "hwp-mcp",
    version="0.1.0",
    description="HWP MCP Server for controlling Hangul Word Processor",
    dependencies=["pywin32>=305"],
    env_vars={}
)

# Global HWP controller instance
hwp_controller = None

def get_hwp_controller():
    """Get or create HwpController instance."""
    global hwp_controller
    if hwp_controller is None:
        logger.info("Creating HwpController instance...")
        try:
            hwp_controller = HwpController()
            if not hwp_controller.connect(visible=True):
                logger.error("Failed to connect to HWP program")
                return None
            logger.info("Successfully connected to HWP program")
        except Exception as e:
            logger.error(f"Error creating HwpController: {str(e)}", exc_info=True)
            return None
    return hwp_controller

@mcp.tool()
def hwp_create() -> str:
    """Create a new HWP document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.create_new_document():
            logger.info("Successfully created new document")
            return "New document created successfully"
        else:
            return "Error: Failed to create new document"
    except Exception as e:
        logger.error(f"Error creating document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_open(path: str) -> str:
    """Open an existing HWP document."""
    try:
        if not path:
            return "Error: File path is required"
        
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.open_document(path):
            logger.info(f"Successfully opened document: {path}")
            return f"Document opened: {path}"
        else:
            return "Error: Failed to open document"
    except Exception as e:
        logger.error(f"Error opening document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_save(path: str = None) -> str:
    """Save the current HWP document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if path:
            if hwp.save_document(path):
                logger.info(f"Successfully saved document to: {path}")
                return f"Document saved to: {path}"
            else:
                return "Error: Failed to save document"
        else:
            temp_path = os.path.join(os.getcwd(), "temp_document.hwp")
            if hwp.save_document(temp_path):
                logger.info(f"Successfully saved document to temporary location: {temp_path}")
                return f"Document saved to: {temp_path}"
            else:
                return "Error: Failed to save document"
    except Exception as e:
        logger.error(f"Error saving document: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_text(text: str) -> str:
    """Insert text at the current cursor position."""
    try:
        if not text:
            return "Error: Text is required"
        
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.insert_text(text):
            logger.info("Successfully inserted text")
            return "Text inserted successfully"
        else:
            return "Error: Failed to insert text"
    except Exception as e:
        logger.error(f"Error inserting text: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_set_font(
    name: str = None, 
    size: int = None, 
    bold: bool = False, 
    italic: bool = False, 
    underline: bool = False
) -> str:
    """Set font properties for selected text."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.set_font(name, size, bold, italic):
            logger.info("Successfully set font")
            return "Font set successfully"
        else:
            return "Error: Failed to set font"
    except Exception as e:
        logger.error(f"Error setting font: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_table(rows: int, cols: int) -> str:
    """Insert a table at the current cursor position."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.insert_table(rows, cols):
            logger.info(f"Successfully inserted {rows}x{cols} table")
            return f"Table inserted with {rows} rows and {cols} columns"
        else:
            return "Error: Failed to insert table"
    except Exception as e:
        logger.error(f"Error inserting table: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_insert_paragraph() -> str:
    """Insert a new paragraph."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.insert_paragraph():
            logger.info("Successfully inserted paragraph")
            return "Paragraph inserted successfully"
        else:
            return "Error: Failed to insert paragraph"
    except Exception as e:
        logger.error(f"Error inserting paragraph: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_get_text() -> str:
    """Get the text content of the current document."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        text = hwp.get_text()
        if text is not None:
            logger.info("Successfully retrieved document text")
            return text
        else:
            return "Error: Failed to get document text"
    except Exception as e:
        logger.error(f"Error getting text: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_close(save: bool = True) -> str:
    """Close the HWP document and connection."""
    try:
        global hwp_controller
        if hwp_controller and hwp_controller.is_hwp_running:
            if hwp_controller.disconnect():
                logger.info("Successfully closed HWP connection")
                hwp_controller = None
                return "HWP connection closed successfully"
            else:
                return "Error: Failed to close HWP connection"
        else:
            return "HWP is already closed"
    except Exception as e:
        logger.error(f"Error closing HWP: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_execute_script(script: str) -> str:
    """
    JavaScript 코드를 실행합니다.
    
    Args:
        script: 실행할 JavaScript 코드
        
    Returns:
        str: 실행 결과
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # 스크립트 실행
        result = hwp.execute_script(script)
        logger.info("Successfully executed JavaScript")
        return result if result else "Script executed successfully"
    except Exception as e:
        logger.error(f"Error executing JavaScript: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_create_table(rows: int, cols: int) -> str:
    """
    현재 커서 위치에 표를 생성합니다.
    
    Args:
        rows: 표의 행 수
        cols: 표의 열 수
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        if hwp.insert_table(rows, cols):
            logger.info(f"Successfully created {rows}x{cols} table")
            return f"Table created with {rows} rows and {cols} columns"
        else:
            return "Error: Failed to create table"
    except Exception as e:
        logger.error(f"Error creating table: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_move_to_cell(table_idx: int, row: int, col: int) -> str:
    """
    표의 특정 셀로 커서를 이동합니다.
    
    Args:
        table_idx: 표 인덱스 (1부터 시작)
        row: 이동할 행 번호 (1부터 시작)
        col: 이동할 열 번호 (1부터 시작)
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # JavaScript를 사용하여 구현 (더 정확한 방법)
        js_code = f"""
        function moveToCell() {{
            var hwp = this;
            
            // 문서 시작으로 이동
            hwp.SetPos(0, 0, 0);
            
            // 표 찾기
            var tableCount = 0;
            while (true) {{
                var found = hwp.HAction.GetDefault("TableCellBlock", hwp.HParameterSet.HShapeObject.HSet);
                if (!found) break;
                
                tableCount++;
                if (tableCount == {table_idx}) {{
                    // 원하는 표를 찾음
                    
                    // 표의 첫 번째 셀로 이동
                    hwp.HAction.Run("TableCellBlock");
                    hwp.HAction.Run("Cancel");
                    
                    // 지정된 행으로 이동
                    for (var r = 1; r < {row}; r++) {{
                        hwp.HAction.Run("TableLowerCell");
                    }}
                    
                    // 지정된 열로 이동
                    for (var c = 1; c < {col}; c++) {{
                        hwp.HAction.Run("TableRightCell");
                    }}
                    
                    return "Successfully moved to cell";
                }}
                
                // 다음 표로 이동
                hwp.HAction.Run("TableRightCell");
            }}
            
            return "Table not found";
        }}
        
        moveToCell();
        """
        
        result = hwp.execute_script(js_code)
        logger.info(f"Move to cell result: {result}")
        return result if result else "Error: Failed to move to cell"
    except Exception as e:
        logger.error(f"Error moving to cell: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_fill_cell(text: str) -> str:
    """
    현재 셀에 텍스트를 입력합니다.
    
    Args:
        text: 입력할 텍스트
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # 현재 셀 내용 지우기
        hwp.run_script("this.SelectAll();")
        hwp.run_script("this.Delete();")
        
        # 텍스트 입력
        if hwp.insert_text(text):
            logger.info("Successfully filled cell")
            return "Cell filled successfully"
        else:
            return "Error: Failed to fill cell"
    except Exception as e:
        logger.error(f"Error filling cell: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_create_field(field_name: str) -> str:
    """
    현재 커서 위치에 필드를 생성합니다.
    
    Args:
        field_name: 생성할 필드 이름
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        js_code = f"""
        function createField() {{
            var hwp = this;
            
            hwp.HAction.GetDefault("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            hwp.HParameterSet.HFieldCreate.FieldName = "{field_name}";
            hwp.HParameterSet.HFieldCreate.FieldType = 0;  // 누름틀 필드
            return hwp.HAction.Execute("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
        }}
        
        createField();
        """
        
        result = hwp.execute_script(js_code)
        logger.info(f"Create field result: {result}")
        return f"Field '{field_name}' created successfully" if result else f"Error: Failed to create field '{field_name}'"
    except Exception as e:
        logger.error(f"Error creating field: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_fill_field(field_name: str, value: str) -> str:
    """
    필드에 값을 채웁니다.
    
    Args:
        field_name: 필드 이름
        value: 채울 값
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        js_code = f"""
        function fillField() {{
            var hwp = this;
            hwp.PutFieldText("{field_name}", "{value}");
            return true;
        }}
        
        fillField();
        """
        
        result = hwp.execute_script(js_code)
        logger.info(f"Fill field result: {result}")
        return f"Field '{field_name}' filled successfully" if result else f"Error: Failed to fill field '{field_name}'"
    except Exception as e:
        logger.error(f"Error filling field: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

if __name__ == "__main__":
    logger.info("Starting HWP MCP stdio server")
    try:
        # Run the FastMCP server with stdio transport
        mcp.run(transport="stdio")
    except Exception as e:
        logger.error(f"Error running server: {str(e)}", exc_info=True)
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1) 