#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import traceback
import logging
import ssl
from threading import Thread
import time

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
def hwp_insert_text(text: str, preserve_linebreaks: bool = True) -> str:
    """Insert text at the current cursor position."""
    try:
        if not text:
            return "Error: Text is required"
        
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # 줄바꿈 문자 처리
        if preserve_linebreaks and ('\n' in text or '\\n' in text):
            # 이스케이프된 줄바꿈 문자(\n)와 실제 줄바꿈 문자 모두 처리
            processed_text = text.replace('\\n', '\n')
            lines = processed_text.split('\n')
            
            success = True
            for i, line in enumerate(lines):
                if not hwp.insert_text(line):
                    success = False
                    break
                # 마지막 줄이 아니면 줄바꿈 삽입
                if i < len(lines) - 1:
                    hwp.insert_paragraph()
            
            if success:
                logger.info("Successfully inserted text with line breaks")
                return "Text with line breaks inserted successfully"
            else:
                return "Error: Failed to insert text with line breaks"
        elif hwp.insert_text(text):
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
    underline: bool = False,
    select_previous_text: bool = False
) -> str:
    """Set font properties for selected text."""
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # 현재 선택된 텍스트에 대해 글자 모양 설정
        if hwp.set_font_style(
            font_name=name,
            font_size=size,
            bold=bold,
            italic=italic,
            underline=underline,
            select_previous_text=select_previous_text
        ):
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
def hwp_ping_pong(message: str = "핑") -> str:
    """
    핑퐁 테스트용 함수입니다. 핑을 보내면 퐁을 응답하고, 퐁을 보내면 핑을 응답합니다.
    
    Args:
        message: 테스트 메시지 (기본값: "핑")
        
    Returns:
        str: 응답 메시지
    """
    try:
        logger.info(f"핑퐁 테스트 함수 호출됨: 메시지 - {message}")
        
        # 메시지에 따라 응답 생성
        if message == "핑":
            response = "퐁"
        elif message == "퐁":
            response = "핑"
        else:
            response = f"모르는 메시지입니다: {message} (핑 또는 퐁을 보내주세요)"
        
        # 응답 생성 시간 기록
        current_time = time.strftime("%Y-%m-%d %H:%M:%S")
        
        # 응답 데이터 구성
        result = {
            "response": response,
            "original_message": message,
            "timestamp": current_time
        }
        
        return json.dumps(result, ensure_ascii=False)
    except Exception as e:
        logger.error(f"핑퐁 테스트 함수 오류: {str(e)}", exc_info=True)
        return f"테스트 오류 발생: {str(e)}"

@mcp.tool()
def hwp_create_table_pyhwpx(rows: int, cols: int, data: str = None, has_header: bool = False) -> str:
    """
    pyhwpx를 사용하여 현재 커서 위치에 표를 생성합니다.
    
    Args:
        rows: 표의 행 수
        cols: 표의 열 수
        data: 표에 채울 데이터 (JSON 형식의 2차원 배열 문자열, 예: '[["항목1", "항목2"], ["값1", "값2"]]')
        has_header: 첫 번째 행을 헤더로 처리할지 여부
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # pyhwpx 방식으로 표 생성
        js_code = f"""
        function createTableWithPyhwpx() {{
            var hwp = this;
            
            // 표 생성
            hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
            hwp.HParameterSet.HTableCreation.Rows = {rows};
            hwp.HParameterSet.HTableCreation.Cols = {cols};
            hwp.HParameterSet.HTableCreation.WidthType = 0;  // 단에 맞춤
            hwp.HParameterSet.HTableCreation.HeightType = 1;  // 절대값
            hwp.HParameterSet.HTableCreation.WidthValue = 0;
            hwp.HParameterSet.HTableCreation.HeightValue = 1000;
            
            // 열 너비 설정
            var colWidth = 8000 / {cols};  // 균등하게 분배
            hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", {cols});
            for (var i = 0; i < {cols}; i++) {{
                hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, colWidth);
            }}
            
            hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
            
            return "표 생성 완료";
        }}
        
        createTableWithPyhwpx();
        """
        
        # 표 생성 실행
        result = hwp.execute_script(js_code)
        logger.info(f"표 생성 결과: {result}")
        
        # 데이터가 제공된 경우 표 채우기
        if data:
            try:
                # JSON 문자열을 파이썬 객체로 변환
                data_array = json.loads(data)
                
                # 표의 크기 확인
                if len(data_array) > rows:
                    data_array = data_array[:rows]  # 행 수 초과 시 자르기
                
                # 각 셀에 데이터 입력
                for row_idx, row_data in enumerate(data_array):
                    if len(row_data) > cols:
                        row_data = row_data[:cols]  # 열 수 초과 시 자르기
                    
                    for col_idx, cell_data in enumerate(row_data):
                        # 셀 위치 계산 (List 인덱스는 0부터 시작, 표는 생성 직후 위치 찾기)
                        cell_list_idx = 3 + row_idx * cols + col_idx
                        
                        # 셀 데이터 입력을 위한 JavaScript 코드
                        cell_js = f"""
                        function fillCell() {{
                            var hwp = this;
                            
                            // 셀로 이동
                            hwp.SetPos({cell_list_idx}, 0, 0);
                            
                            // 셀 내용 선택 및 삭제
                            hwp.SelectAll();
                            hwp.Delete();
                            
                            // 텍스트 입력
                            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
                            hwp.HParameterSet.HInsertText.Text = "{str(cell_data)}";
                            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
                            
                            return "셀 데이터 입력 완료";
                        }}
                        
                        fillCell();
                        """
                        
                        # 셀 데이터 입력 실행
                        hwp.execute_script(cell_js)
            
                # 헤더 스타일 설정 (첫 번째 행이 헤더인 경우)
                if has_header and data_array and len(data_array) > 0:
                    for col_idx in range(min(cols, len(data_array[0]))):
                        header_js = f"""
                        function styleHeader() {{
                            var hwp = this;
                            
                            // 헤더 셀로 이동
                            hwp.SetPos({3 + col_idx}, 0, 0);
                            
                            // 셀 내용 선택
                            hwp.SelectAll();
                            
                            // 글자 굵게 설정
                            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
                            hwp.HParameterSet.HCharShape.Bold = 1;
                            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet);
                            
                            // 가운데 정렬
                            hwp.HAction.GetDefault("ParaShape", hwp.HParameterSet.HParaShape.HSet);
                            hwp.HParameterSet.HParaShape.Align = 1;  // 가운데 정렬
                            hwp.HAction.Execute("ParaShape", hwp.HParameterSet.HParaShape.HSet);
                            
                            return "헤더 스타일 적용 완료";
                        }}
                        
                        styleHeader();
                        """
                        
                        # 헤더 스타일 적용 실행
                        hwp.execute_script(header_js)
                        
                return f"표 생성 및 데이터 입력 완료 ({rows}x{cols} 크기)"
            except Exception as data_error:
                logger.error(f"표 데이터 입력 중 오류: {str(data_error)}", exc_info=True)
                return f"표는 생성되었으나 데이터 입력 중 오류 발생: {str(data_error)}"
        
        return f"표 생성 완료 ({rows}x{cols} 크기)"
    except Exception as e:
        logger.error(f"pyhwpx 표 생성 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_table_set_cell_text(row: int, col: int, text: str) -> str:
    """
    pyhwpx를 사용하여 표의 특정 셀에 텍스트를 입력합니다.
    
    Args:
        row: 셀의 행 번호 (1부터 시작)
        col: 셀의 열 번호 (1부터 시작)
        text: 입력할 텍스트
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # 셀 이동 및 텍스트 입력을 위한 JavaScript 코드
        js_code = f"""
        function setCellText() {{
            var hwp = this;
            
            // 문서의 현재 표를 찾아서 이동
            // 주의: 현재 커서 위치의 표를 기준으로 작동
            
            // 현재 표에서 특정 셀로 이동
            var moveSuccess = false;
            
            try {{
                // 표의 첫 번째 셀로 이동
                hwp.TableCellRefresh();
                hwp.HAction.Run("TableCellBlock");
                hwp.HAction.Run("Cancel");
                
                // 지정된 행으로 이동 (1행부터 시작)
                for (var r = 1; r < {row}; r++) {{
                    hwp.HAction.Run("TableLowerCell");
                }}
                
                // 지정된 열로 이동 (1열부터 시작)
                for (var c = 1; c < {col}; c++) {{
                    hwp.HAction.Run("TableRightCell");
                }}
                
                moveSuccess = true;
            }} catch (e) {{
                return "셀 이동 실패: " + e.message;
            }}
            
            if (!moveSuccess) {{
                return "지정된 셀({row}, {col})로 이동할 수 없습니다.";
            }}
            
            // 셀 내용 선택 및 삭제
            hwp.SelectAll();
            hwp.Delete();
            
            // 텍스트 입력
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "{text}";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            return "셀 텍스트 설정 완료";
        }}
        
        setCellText();
        """
        
        # JavaScript 실행
        result = hwp.execute_script(js_code)
        logger.info(f"셀 텍스트 설정 결과: {result}")
        
        return f"셀({row}, {col})에 텍스트 입력 완료"
    except Exception as e:
        logger.error(f"셀 텍스트 설정 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_table_merge_cells(start_row: int, start_col: int, end_row: int, end_col: int) -> str:
    """
    pyhwpx를 사용하여 표의 특정 범위의 셀을 병합합니다.
    
    Args:
        start_row: 시작 행 번호 (1부터 시작)
        start_col: 시작 열 번호 (1부터 시작)
        end_row: 종료 행 번호 (1부터 시작)
        end_col: 종료 열 번호 (1부터 시작)
        
    Returns:
        str: 결과 메시지
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return "Error: Failed to connect to HWP program"
        
        # 셀 병합을 위한 JavaScript 코드
        js_code = f"""
        function mergeCells() {{
            var hwp = this;
            
            try {{
                // 시작 셀로 이동
                hwp.TableCellRefresh();
                hwp.HAction.Run("TableCellBlock");
                hwp.HAction.Run("Cancel");
                
                // 지정된 시작 행으로 이동
                for (var r = 1; r < {start_row}; r++) {{
                    hwp.HAction.Run("TableLowerCell");
                }}
                
                // 지정된 시작 열로 이동
                for (var c = 1; c < {start_col}; c++) {{
                    hwp.HAction.Run("TableRightCell");
                }}
                
                // 셀 선택 모드로 전환
                hwp.HAction.Run("TableCellBlock");
                
                // 종료 셀까지 선택 영역 확장
                var rowDiff = {end_row} - {start_row};
                var colDiff = {end_col} - {start_col};
                
                // 행 방향 확장
                for (var r = 0; r < rowDiff; r++) {{
                    hwp.HAction.Run("TableLowerCell");
                    hwp.HAction.Run("TableCellBlockExtend");
                }}
                
                // 열 방향 확장
                for (var c = 0; c < colDiff; c++) {{
                    hwp.HAction.Run("TableRightCell");
                    hwp.HAction.Run("TableCellBlockExtend");
                }}
                
                // 셀 병합 실행
                hwp.HAction.Run("TableMergeCell");
                
                return "셀 병합 완료";
            }} catch (e) {{
                return "셀 병합 실패: " + e.message;
            }}
        }}
        
        mergeCells();
        """
        
        # JavaScript 실행
        result = hwp.execute_script(js_code)
        logger.info(f"셀 병합 결과: {result}")
        
        return f"셀 병합 완료 ({start_row},{start_col}) - ({end_row},{end_col})"
    except Exception as e:
        logger.error(f"셀 병합 중 오류: {str(e)}", exc_info=True)
        return f"Error: {str(e)}"

@mcp.tool()
def hwp_create_complete_document(document_spec: dict) -> dict:
    """
    전체 문서를 한 번의 호출로 작성합니다. 문서 구조, 내용 및 서식을 JSON으로 정의하여 전달합니다.
    
    Args:
        document_spec (dict): 문서 사양을 담은 딕셔너리. 다음과 같은 구조를 가집니다:
            {
                "title": "문서 제목",             # 선택 사항
                "filename": "저장할_파일명.hwp",  # 선택 사항, 저장할 경우
                "elements": [                    # 필수: 문서를 구성하는 요소 배열
                    {
                        "type": "heading",       # 요소 유형 (heading, text, paragraph, table, etc.)
                        "content": "제목",        # 요소 내용
                        "properties": {          # 요소 속성 (선택 사항)
                            "font_size": 16,
                            "bold": true,
                            ...
                        }
                    },
                    ...
                ],
                "special_type": {               # 특수 문서 유형 (선택 사항)
                    "type": "report",           # 보고서 등 특수 문서 유형
                    "params": { ... }           # 특수 문서에 필요한 매개변수
                },
                "save": true                    # 저장 여부 (선택 사항)
            }
    
    Returns:
        dict: 문서 생성 결과
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"status": "error", "message": "Failed to connect to HWP program"}
        
        # 새 문서 생성
        if not hwp.create_new_document():
            return {"status": "error", "message": "Failed to create new document"}
        
        # 문서 사양 유효성 검사
        if not document_spec:
            return {"status": "error", "message": "Document specification is required"}
        
        if "special_type" in document_spec:
            # 특수 문서 유형 처리 (보고서 등)
            special_type = document_spec["special_type"]
            special_type_name = special_type.get("type", "")
            special_params = special_type.get("params", {})
            
            # 보고서 처리
            if special_type_name == "report":
                return _create_report(hwp, special_params, document_spec)
            
            # 편지 처리
            elif special_type_name == "letter":
                return _create_letter(hwp, special_params, document_spec)
            
            else:
                return {"status": "error", "message": f"Unknown special document type: {special_type_name}"}
        
        # 일반 문서 처리
        elif "elements" in document_spec:
            elements = document_spec.get("elements", [])
            
            # 문서 요소 처리
            for element in elements:
                element_type = element.get("type", "")
                content = element.get("content", "")
                properties = element.get("properties", {})
                
                # 요소 유형에 따른 처리
                if element_type == "heading":
                    # 제목 스타일 설정
                    font_size = properties.get("font_size", 16)
                    bold = properties.get("bold", True)
                    hwp.set_font(None, font_size, bold, False)
                    hwp.insert_text(content)
                    hwp.insert_paragraph()
                
                elif element_type == "text":
                    # 텍스트 스타일 설정
                    font_size = properties.get("font_size", 10)
                    bold = properties.get("bold", False)
                    italic = properties.get("italic", False)
                    hwp.set_font(None, font_size, bold, italic)
                    hwp.insert_text(content)
                
                elif element_type == "paragraph":
                    hwp.insert_paragraph()
                
                elif element_type == "table":
                    rows = properties.get("rows", 0)
                    cols = properties.get("cols", 0)
                    data = properties.get("data", [])
                    
                    if rows > 0 and cols > 0:
                        hwp.insert_table(rows, cols)
                        
                        # 테이블 데이터 채우기 (구현 필요)
                        # 현재는 표만 생성하고 데이터는 처리하지 않음
                
                else:
                    logger.warning(f"Unknown element type: {element_type}")
        
        else:
            return {"status": "error", "message": "Document must contain 'elements' or 'special_type'"}
        
        # 문서 저장
        if document_spec.get("save", False):
            filename = document_spec.get("filename", "generated_document.hwp")
            if hwp.save_document(filename):
                return {
                    "status": "success", 
                    "message": "Document created and saved successfully",
                    "saved_path": filename
                }
            else:
                return {
                    "status": "partial_success", 
                    "message": "Document created but failed to save"
                }
        
        return {"status": "success", "message": "Document created successfully"}
    
    except Exception as e:
        logger.error(f"Error creating document: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

def _create_report(hwp, params, document_spec):
    """보고서 문서를 생성합니다."""
    try:
        title = params.get("title", "보고서 제목")
        author = params.get("author", "작성자")
        date = params.get("date", time.strftime("%Y년 %m월 %d일"))
        sections = params.get("sections", [{"title": "섹션 제목", "content": "섹션 내용"}])
        
        # 제목 페이지
        hwp.set_font(None, 22, True, False)
        hwp.insert_text(title)
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        hwp.set_font(None, 14, False, False)
        hwp.insert_text(f"작성자: {author}")
        hwp.insert_paragraph()
        hwp.insert_text(f"작성일: {date}")
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        # 각 섹션
        for section in sections:
            section_title = section.get("title", "")
            section_content = section.get("content", "")
            
            hwp.set_font(None, 16, True, False)
            hwp.insert_text(section_title)
            hwp.insert_paragraph()
            
            hwp.set_font(None, 12, False, False)
            hwp.insert_text(section_content)
            hwp.insert_paragraph()
            hwp.insert_paragraph()
        
        # 문서 저장
        result = {"status": "success", "message": "Report created successfully"}
        if document_spec.get("save", False):
            filename = document_spec.get("filename", "report.hwp")
            if hwp.save_document(filename):
                result["saved_path"] = filename
            else:
                result["message"] = "Report created but failed to save"
                result["status"] = "partial_success"
        
        return result
    
    except Exception as e:
        logger.error(f"Error creating report: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

def _create_letter(hwp, params, document_spec):
    """편지 문서를 생성합니다."""
    try:
        title = params.get("title", "제목 없음")
        recipient = params.get("recipient", "받는 사람")
        content = params.get("content", "내용을 입력하세요.")
        sender = params.get("sender", "보내는 사람")
        date = params.get("date", time.strftime("%Y년 %m월 %d일"))
        
        # 제목 (굵게, 크게)
        hwp.set_font(None, 16, True, False)
        hwp.insert_text(title)
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        # 받는 사람
        hwp.set_font(None, 12, False, False)
        hwp.insert_text(f"받는 사람: {recipient}")
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        # 내용
        hwp.set_font(None, 12, False, False)
        hwp.insert_text(content)
        hwp.insert_paragraph()
        hwp.insert_paragraph()
        
        # 날짜 (오른쪽 정렬)
        # 오른쪽 정렬은 현재 구현되어 있지 않으므로 공백으로 대체
        hwp.set_font(None, 12, False, False)
        hwp.insert_text("".ljust(40) + date)
        hwp.insert_paragraph()
        
        # 보내는 사람 (오른쪽 정렬, 굵게)
        hwp.set_font(None, 12, True, False)
        hwp.insert_text("".ljust(40) + sender)
        
        # 문서 저장
        result = {"status": "success", "message": "Letter created successfully"}
        if document_spec.get("save", False):
            filename = document_spec.get("filename", "letter.hwp")
            if hwp.save_document(filename):
                result["saved_path"] = filename
            else:
                result["message"] = "Letter created but failed to save"
                result["status"] = "partial_success"
        
        return result
    
    except Exception as e:
        logger.error(f"Error creating letter: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

@mcp.tool()
def hwp_create_document_from_text(content: str, title: str = None, format_content: bool = True, save_filename: str = None, preserve_linebreaks: bool = True) -> dict:
    """
    단일 문자열로 된 텍스트 내용으로 문서를 생성합니다.
    
    Args:
        content (str): 문서 내용 (형식을 자동으로 감지하고 처리)
        title (str, optional): 문서 제목. 없으면 첫 줄을 제목으로 사용.
        format_content (bool): 내용 자동 포맷팅 여부 (줄바꿈, 문단 구분 등)
        save_filename (str, optional): 저장할 파일 이름. 제공되지 않으면 저장하지 않음.
        preserve_linebreaks (bool): 줄바꿈 유지 여부. True이면 원본 텍스트의 모든 줄바꿈 유지.
        
    Returns:
        dict: 문서 생성 결과
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"status": "error", "message": "Failed to connect to HWP program"}
        
        # 새 문서 생성
        if not hwp.create_new_document():
            return {"status": "error", "message": "Failed to create new document"}
        
        # 내용이 없는 경우
        if not content:
            return {"status": "error", "message": "Document content is required"}
        
        # 내용을 줄로 분리
        lines = content.split('\n')
        
        # 빈 줄을 기준으로 블록 구분
        blocks = []
        current_block = []
        
        for line in lines:
            if line.strip():  # 빈 줄이 아닌 경우
                current_block.append(line)
            else:  # 빈 줄인 경우 블록 구분
                if current_block:
                    blocks.append(current_block)
                    current_block = []
        
        # 마지막 블록 추가
        if current_block:
            blocks.append(current_block)
        
        # 제목 처리
        if not title and blocks:
            # 첫 번째 블록의 첫 번째 줄을 제목으로 사용
            title = blocks[0][0]
            if len(blocks[0]) > 1:
                blocks[0] = blocks[0][1:]  # 첫 번째 줄 제거
            else:
                blocks = blocks[1:]  # 첫 번째 블록 제거
        
        # 제목 추가
        if title:
            # 먼저 폰트 설정 후 텍스트 입력 (수정된 방식)
            hwp.set_font(None, 16, True, False)
            hwp.insert_text(title)
            hwp.insert_paragraph()
            hwp.insert_paragraph()
        
        # 내용 자동 포맷팅
        if format_content:
            # 블록 단위로 처리
            for block in blocks:
                # 블록 내 첫 번째 줄로 블록 유형 판단
                first_line = block[0].strip() if block else ""
                
                # 제목 형식 감지 (예: #으로 시작하면 제목)
                if first_line.startswith('#'):
                    level = 0
                    for char in first_line:
                        if char == '#':
                            level += 1
                        else:
                            break
                    
                    heading_text = first_line[level:].strip()
                    font_size = max(11, 16 - (level - 1))  # 제목 레벨에 따라 글자 크기 조정
                    
                    # 먼저 폰트 설정 후 텍스트 입력 (수정된 방식)
                    hwp.set_font(None, font_size, True, False)
                    hwp.insert_text(heading_text)
                    hwp.insert_paragraph()
                    
                    # 제목 이후의 줄들 처리 (있을 경우)
                    if len(block) > 1:
                        hwp.set_font(None, 11, False, False)
                        for line in block[1:]:
                            hwp.insert_text(line)
                            hwp.insert_paragraph()
                
                # 글머리 기호 감지 (예: - 또는 * 으로 시작하면 글머리 기호)
                elif first_line.startswith(('-', '*', '•')):
                    hwp.set_font(None, 11, False, False)
                    for line in block:
                        line_stripped = line.strip()
                        if line_stripped.startswith(('-', '*', '•')):
                            content_text = line_stripped[1:].strip()
                            hwp.insert_text(f"• {content_text}")
                        else:
                            hwp.insert_text(line_stripped)
                        hwp.insert_paragraph()
                
                # 시 또는 줄바꿈이 중요한 텍스트 (각 줄을 개별적으로 처리)
                elif preserve_linebreaks:
                    hwp.set_font(None, 11, False, False)
                    for line in block:
                        hwp.insert_text(line)
                        hwp.insert_paragraph()
                
                # 일반 텍스트 (블록 전체를 하나의 단락으로 처리)
                else:
                    hwp.set_font(None, 11, False, False)
                    block_text = '\n'.join(block)
                    hwp.insert_text(block_text)
                    hwp.insert_paragraph()
                
                # 블록 사이에 추가 줄바꿈
                hwp.insert_paragraph()
        
        # 자동 포맷팅 없이 그대로 삽입 (줄바꿈 보존)
        else:
            hwp.set_font(None, 11, False, False)
            for line in lines:
                if line.strip():  # 내용이 있는 줄
                    hwp.insert_text(line)
                hwp.insert_paragraph()  # 빈 줄이든 내용이 있는 줄이든 항상 줄바꿈
        
        # 문서 저장
        result = {"status": "success", "message": "Document created from text successfully"}
        if save_filename:
            if hwp.save_document(save_filename):
                result["saved_path"] = save_filename
            else:
                result["message"] = "Document created but failed to save"
                result["status"] = "partial_success"
        
        return result
    
    except Exception as e:
        logger.error(f"Error creating document from text: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

@mcp.tool()
def hwp_batch_operations(operations: list) -> dict:
    """
    여러 HWP 작업을 한 번의 호출로 일괄 처리합니다.
    
    Args:
        operations (list): 실행할 작업 목록. 각 작업은 다음 형식의 딕셔너리입니다:
            {
                "operation": "작업명", # 예: "create", "set_font", "insert_text" 등
                "params": {파라미터 딕셔너리}  # 해당 작업에 필요한 파라미터
            }
    
    Returns:
        dict: 각 작업의 실행 결과
    """
    try:
        hwp = get_hwp_controller()
        if not hwp:
            return {"status": "error", "message": "Failed to connect to HWP program"}
        
        results = []
        
        for op in operations:
            operation = op.get("operation", "")
            params = op.get("params", {})
            
            result = {"operation": operation, "status": "success", "message": ""}
            
            try:
                if operation == "create":
                    if hwp.create_new_document():
                        result["message"] = "New document created successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to create new document"
                
                elif operation == "open":
                    path = params.get("path", "")
                    if not path:
                        result["status"] = "error"
                        result["message"] = "File path is required"
                    elif hwp.open_document(path):
                        result["message"] = f"Document opened: {path}"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to open document"
                
                elif operation == "save":
                    path = params.get("path", None)
                    if path and hwp.save_document(path):
                        result["message"] = f"Document saved to: {path}"
                    elif not path:
                        temp_path = os.path.join(os.getcwd(), "temp_document.hwp")
                        if hwp.save_document(temp_path):
                            result["message"] = f"Document saved to: {temp_path}"
                            result["path"] = temp_path
                        else:
                            result["status"] = "error"
                            result["message"] = "Failed to save document"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to save document"
                
                elif operation == "insert_text":
                    text = params.get("text", "")
                    preserve_linebreaks = params.get("preserve_linebreaks", True)
                    
                    if not text:
                        result["status"] = "error"
                        result["message"] = "Text is required"
                    elif preserve_linebreaks and ('\n' in text or '\\n' in text):
                        # 줄바꿈 보존 처리 개선
                        # 이스케이프된 줄바꿈 문자(\n)와 실제 줄바꿈 문자 모두 처리
                        # 먼저 이스케이프된 줄바꿈 문자를 실제 줄바꿈으로 변환
                        processed_text = text.replace('\\n', '\n')
                        lines = processed_text.split('\n')
                        
                        success = True
                        for i, line in enumerate(lines):
                            if not hwp.insert_text(line):
                                success = False
                                break
                            # 마지막 줄이 아니면 줄바꿈 삽입
                            if i < len(lines) - 1:
                                hwp.insert_paragraph()
                        
                        if success:
                            result["message"] = "Text with line breaks inserted successfully"
                        else:
                            result["status"] = "error"
                            result["message"] = "Failed to insert text with line breaks"
                    elif hwp.insert_text(text):
                        result["message"] = "Text inserted successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to insert text"
                
                elif operation == "set_font":
                    name = params.get("name", None)
                    size = params.get("size", None)
                    bold = params.get("bold", False)
                    italic = params.get("italic", False)
                    underline = params.get("underline", False)
                    select_previous_text = params.get("select_previous_text", False)
                    
                    if hwp.set_font_style(font_name=name, font_size=size, bold=bold, italic=italic, underline=underline, select_previous_text=select_previous_text):
                        result["message"] = "Font set successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to set font"
                
                elif operation == "insert_paragraph":
                    count = params.get("count", 1)  # 여러 줄 삽입 가능
                    success = True
                    for _ in range(count):
                        if not hwp.insert_paragraph():
                            success = False
                            break
                    
                    if success:
                        result["message"] = f"{count} paragraph(s) inserted successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to insert paragraph"
                
                elif operation == "insert_table":
                    rows = params.get("rows", 0)
                    cols = params.get("cols", 0)
                    data = params.get("data", [])
                    
                    if rows <= 0 or cols <= 0:
                        result["status"] = "error"
                        result["message"] = "Valid rows and cols are required"
                    elif hwp.insert_table(rows, cols):
                        # 데이터가 있으면 테이블 채우기 (향후 구현)
                        result["message"] = f"Table with {rows} rows and {cols} columns inserted successfully"
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to insert table"
                
                elif operation == "get_text":
                    text = hwp.get_text()
                    if text is not None:
                        result["message"] = "Text retrieved successfully"
                        result["text"] = text
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to retrieve text"
                
                elif operation == "close":
                    save = params.get("save", True)
                    if hwp.disconnect():
                        result["message"] = "Document closed successfully"
                        # 전역 변수 초기화
                        global hwp_controller
                        hwp_controller = None
                    else:
                        result["status"] = "error"
                        result["message"] = "Failed to close document"
                
                # 새로 추가: 문서 한 번에 생성
                elif operation == "create_document_from_text":
                    content = params.get("content", "")
                    title = params.get("title", None)
                    format_content = params.get("format_content", True)
                    save_filename = params.get("save_filename", None)
                    preserve_linebreaks = params.get("preserve_linebreaks", True)
                    
                    if not content:
                        result["status"] = "error"
                        result["message"] = "Document content is required"
                    else:
                        # 내부적으로 기존 함수 호출
                        doc_result = hwp_create_document_from_text(
                            content=content,
                            title=title,
                            format_content=format_content,
                            save_filename=save_filename,
                            preserve_linebreaks=preserve_linebreaks
                        )
                        
                        result["status"] = doc_result.get("status", "error")
                        result["message"] = doc_result.get("message", "Unknown error")
                        if "saved_path" in doc_result:
                            result["saved_path"] = doc_result["saved_path"]
                
                else:
                    result["status"] = "error"
                    result["message"] = f"Unknown operation: {operation}"
            
            except Exception as e:
                result["status"] = "error"
                result["message"] = f"Error in operation '{operation}': {str(e)}"
            
            results.append(result)
        
        return {"status": "success", "results": results}
    
    except Exception as e:
        logger.error(f"Error in batch operations: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}

if __name__ == "__main__":
    logger.info("Starting HWP MCP stdio server")
    try:
        # Run the FastMCP server with stdio transport
        mcp.run(transport="stdio")
    except Exception as e:
        logger.error(f"Error running server: {str(e)}", exc_info=True)
        print(f"Error: {str(e)}", file=sys.stderr)
        sys.exit(1) 