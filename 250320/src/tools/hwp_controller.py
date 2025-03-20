"""
한글(HWP) 문서를 제어하기 위한 컨트롤러 모듈
win32com을 이용하여 한글 프로그램을 자동화합니다.
"""

import os
import win32com.client
import win32gui
import win32con
import time
from typing import Optional, List, Dict, Any, Tuple


class HwpController:
    """한글 문서를 제어하는 클래스"""

    def __init__(self):
        """한글 애플리케이션 인스턴스를 초기화합니다."""
        self.hwp = None
        self.visible = True
        self.is_hwp_running = False
        self.current_document_path = None

    def connect(self, visible: bool = True, register_security_module: bool = True) -> bool:
        """
        한글 프로그램에 연결합니다.
        
        Args:
            visible (bool): 한글 창을 화면에 표시할지 여부
            register_security_module (bool): 보안 모듈을 등록할지 여부
            
        Returns:
            bool: 연결 성공 여부
        """
        try:
            self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            
            # 보안 모듈 등록 (파일 경로 체크 보안 경고창 방지)
            if register_security_module:
                try:
                    # 보안 모듈 DLL 경로 - 실제 파일이 위치한 경로로 수정 필요
                    module_path = os.path.abspath("D:/hwp-mcp/security_module/FilePathCheckerModuleExample.dll")
                    self.hwp.RegisterModule("FilePathCheckerModuleExample", module_path)
                    print("보안 모듈이 등록되었습니다.")
                except Exception as e:
                    print(f"보안 모듈 등록 실패 (무시하고 계속 진행): {e}")
            
            self.visible = visible
            self.hwp.XHwpWindows.Item(0).Visible = visible
            self.is_hwp_running = True
            return True
        except Exception as e:
            print(f"한글 프로그램 연결 실패: {e}")
            return False

    def disconnect(self) -> bool:
        """
        한글 프로그램 연결을 종료합니다.
        
        Returns:
            bool: 종료 성공 여부
        """
        try:
            if self.is_hwp_running:
                # HwpObject를 해제합니다
                self.hwp = None
                self.is_hwp_running = False
                
            return True
        except Exception as e:
            print(f"한글 프로그램 종료 실패: {e}")
            return False

    def create_new_document(self) -> bool:
        """
        새 문서를 생성합니다.
        
        Returns:
            bool: 생성 성공 여부
        """
        try:
            if not self.is_hwp_running:
                self.connect()
            
            self.hwp.Run("FileNew")
            self.current_document_path = None
            return True
        except Exception as e:
            print(f"새 문서 생성 실패: {e}")
            return False

    def open_document(self, file_path: str) -> bool:
        """
        문서를 엽니다.
        
        Args:
            file_path (str): 열 문서의 경로
            
        Returns:
            bool: 열기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                self.connect()
            
            abs_path = os.path.abspath(file_path)
            self.hwp.Open(abs_path)
            self.current_document_path = abs_path
            return True
        except Exception as e:
            print(f"문서 열기 실패: {e}")
            return False

    def save_document(self, file_path: Optional[str] = None) -> bool:
        """
        문서를 저장합니다.
        
        Args:
            file_path (str, optional): 저장할 경로. None이면 현재 경로에 저장.
            
        Returns:
            bool: 저장 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            if file_path:
                abs_path = os.path.abspath(file_path)
                # 파일 형식과 경로 모두 지정하여 저장
                self.hwp.SaveAs(abs_path, "HWP", "")
                self.current_document_path = abs_path
            else:
                if self.current_document_path:
                    self.hwp.Save()
                else:
                    # 저장 대화 상자 표시 (파라미터 없이 호출)
                    self.hwp.SaveAs()
                    # 대화 상자에서 사용자가 선택한 경로를 알 수 없으므로 None 유지
            
            return True
        except Exception as e:
            print(f"문서 저장 실패: {e}")
            return False

    def insert_text(self, text: str) -> bool:
        """
        현재 커서 위치에 텍스트를 삽입합니다.
        
        Args:
            text (str): 삽입할 텍스트
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            return True
        except Exception as e:
            print(f"텍스트 삽입 실패: {e}")
            return False

    def set_font(self, font_name: str, font_size: int, bold: bool = False, italic: bool = False) -> bool:
        """
        현재 선택된 텍스트의 글꼴을 설정합니다.
        
        Args:
            font_name (str): 글꼴 이름
            font_size (int): 글꼴 크기
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 간단한 방식으로 처리: 명령어로 직접 실행
            self.hwp.Run("SelectAll")  # 전체 선택
            
            # 매크로 명령 사용
            size_pt = font_size * 100  # 폰트 크기 단위 변환
            
            # 굵게 및 기울임꼴 설정
            boldValue = "1" if bold else "0"
            italicValue = "1" if italic else "0"
            
            # 직접 매크로 명령 실행
            self.hwp.Run(f'CharShape "{font_name}" {size_pt} {boldValue} {italicValue} 0 0 "" 0 "" 0')
            
            # 선택 취소 추가
            self.hwp.Run("Cancel")
            
            return True
        except Exception as e:
            print(f"글꼴 설정 실패: {e}")
            return False

    def insert_table(self, rows: int, cols: int) -> bool:
        """
        현재 커서 위치에 표를 삽입합니다.
        
        Args:
            rows (int): 행 수
            cols (int): 열 수
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.HAction.GetDefault("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            self.hwp.HParameterSet.HTableCreation.Rows = rows
            self.hwp.HParameterSet.HTableCreation.Cols = cols
            self.hwp.HParameterSet.HTableCreation.WidthType = 0  # 0: 단에 맞춤, 1: 절대값
            self.hwp.HParameterSet.HTableCreation.HeightType = 1  # 0: 자동, 1: 절대값
            self.hwp.HParameterSet.HTableCreation.WidthValue = 0  # 단에 맞춤이므로 무시됨
            self.hwp.HParameterSet.HTableCreation.HeightValue = 1000  # 셀 높이(hwpunit)
            
            # 각 열의 너비를 설정 (모두 동일하게)
            # PageWidth 대신 고정 값 사용
            col_width = 8000 // cols  # 전체 너비를 열 수로 나눔
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", cols)
            for i in range(cols):
                self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, col_width)
                
            self.hwp.HAction.Execute("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            return True
        except Exception as e:
            print(f"표 삽입 실패: {e}")
            return False

    def insert_image(self, image_path: str, width: int = 0, height: int = 0) -> bool:
        """
        현재 커서 위치에 이미지를 삽입합니다.
        
        Args:
            image_path (str): 이미지 파일 경로
            width (int): 이미지 너비(0이면 원본 크기)
            height (int): 이미지 높이(0이면 원본 크기)
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            abs_path = os.path.abspath(image_path)
            if not os.path.exists(abs_path):
                print(f"이미지 파일을 찾을 수 없습니다: {abs_path}")
                return False
                
            self.hwp.HAction.GetDefault("InsertPicture", self.hwp.HParameterSet.HInsertPicture.HSet)
            self.hwp.HParameterSet.HInsertPicture.FileName = abs_path
            self.hwp.HParameterSet.HInsertPicture.Width = width
            self.hwp.HParameterSet.HInsertPicture.Height = height
            self.hwp.HParameterSet.HInsertPicture.Embed = 1  # 0: 링크, 1: 파일 포함
            self.hwp.HAction.Execute("InsertPicture", self.hwp.HParameterSet.HInsertPicture.HSet)
            return True
        except Exception as e:
            print(f"이미지 삽입 실패: {e}")
            return False

    def find_text(self, text: str) -> bool:
        """
        문서에서 텍스트를 찾습니다.
        
        Args:
            text (str): 찾을 텍스트
            
        Returns:
            bool: 찾기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 간단한 매크로 명령 사용
            self.hwp.Run("MoveDocBegin")  # 문서 처음으로 이동
            
            # 찾기 명령 실행 (매크로 사용)
            result = self.hwp.Run(f'FindText "{text}" 1')  # 1=정방향검색
            return result  # True 또는 False 반환
        except Exception as e:
            print(f"텍스트 찾기 실패: {e}")
            return False

    def replace_text(self, find_text: str, replace_text: str, replace_all: bool = False) -> bool:
        """
        문서에서 텍스트를 찾아 바꿉니다.
        
        Args:
            find_text (str): 찾을 텍스트
            replace_text (str): 바꿀 텍스트
            replace_all (bool): 모두 바꾸기 여부
            
        Returns:
            bool: 바꾸기 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 매크로 명령 사용
            self.hwp.Run("MoveDocBegin")  # 문서 처음으로 이동
            
            if replace_all:
                # 모두 바꾸기 명령 실행
                result = self.hwp.Run(f'ReplaceAll "{find_text}" "{replace_text}" 0 0 0 0 0 0')
                return bool(result)
            else:
                # 하나만 바꾸기 (찾고 바꾸기)
                found = self.hwp.Run(f'FindText "{find_text}" 1')
                if found:
                    result = self.hwp.Run(f'Replace "{replace_text}"')
                    return bool(result)
                return False
        except Exception as e:
            print(f"텍스트 바꾸기 실패: {e}")
            return False

    def get_text(self) -> str:
        """
        현재 문서의 전체 텍스트를 가져옵니다.
        
        Returns:
            str: 문서 텍스트
        """
        try:
            if not self.is_hwp_running:
                return ""
            
            return self.hwp.GetTextFile("TEXT", "")
        except Exception as e:
            print(f"텍스트 가져오기 실패: {e}")
            return ""

    def set_page_setup(self, orientation: str = "portrait", margin_left: int = 1000, 
                     margin_right: int = 1000, margin_top: int = 1000, margin_bottom: int = 1000) -> bool:
        """
        페이지 설정을 변경합니다.
        
        Args:
            orientation (str): 용지 방향 ('portrait' 또는 'landscape')
            margin_left (int): 왼쪽 여백(hwpunit)
            margin_right (int): 오른쪽 여백(hwpunit)
            margin_top (int): 위쪽 여백(hwpunit)
            margin_bottom (int): 아래쪽 여백(hwpunit)
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 매크로 명령 사용
            orient_val = 0 if orientation.lower() == "portrait" else 1
            
            # 페이지 설정 매크로
            result = self.hwp.Run(f"PageSetup3 {orient_val} {margin_left} {margin_right} {margin_top} {margin_bottom}")
            return bool(result)
        except Exception as e:
            print(f"페이지 설정 실패: {e}")
            return False

    def insert_paragraph(self) -> bool:
        """
        새 단락을 삽입합니다.
        
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.HAction.Run("BreakPara")
            return True
        except Exception as e:
            print(f"단락 삽입 실패: {e}")
            return False

    def select_all(self) -> bool:
        """
        문서 전체를 선택합니다.
        
        Returns:
            bool: 선택 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            self.hwp.Run("SelectAll")
            return True
        except Exception as e:
            print(f"전체 선택 실패: {e}")
            return False

    def fill_cell_field(self, field_name: str, value: str, n: int = 1) -> bool:
        """
        동일한 이름의 셀필드 중 n번째에만 값을 채웁니다.
        위키독스 예제: https://wikidocs.net/261646
        
        Args:
            field_name (str): 필드 이름
            value (str): 채울 값
            n (int): 몇 번째 필드에 값을 채울지 (1부터 시작)
            
        Returns:
            bool: 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
                
            # 1. 필드 목록 가져오기
            # HGO_GetFieldList은 현재 문서에 있는 모든 필드 목록을 가져옵니다.
            self.hwp.HAction.GetDefault("HGo_GetFieldList", self.hwp.HParameterSet.HGo.HSet)
            self.hwp.HAction.Execute("HGo_GetFieldList", self.hwp.HParameterSet.HGo.HSet)
            
            # 2. 필드 이름이 동일한 모든 셀필드 찾기
            field_list = []
            field_count = self.hwp.HParameterSet.HGo.FieldList.Count
            
            for i in range(field_count):
                field_info = self.hwp.HParameterSet.HGo.FieldList.Item(i)
                if field_info.FieldName == field_name:
                    field_list.append((field_info.FieldName, i))
            
            # 3. n번째 필드가 존재하는지 확인 (인덱스는 0부터 시작하므로 n-1)
            if len(field_list) < n:
                print(f"해당 이름의 필드가 충분히 없습니다. 필요: {n}, 존재: {len(field_list)}")
                return False
                
            # 4. n번째 필드의 위치로 이동
            target_field_idx = field_list[n-1][1]
            
            # HGo_SetFieldText를 사용하여 해당 필드 위치로 이동한 후 텍스트 설정
            self.hwp.HAction.GetDefault("HGo_SetFieldText", self.hwp.HParameterSet.HGo.HSet)
            self.hwp.HParameterSet.HGo.HSet.SetItem("FieldIdx", target_field_idx)
            self.hwp.HParameterSet.HGo.HSet.SetItem("Text", value)
            self.hwp.HAction.Execute("HGo_SetFieldText", self.hwp.HParameterSet.HGo.HSet)
            
            return True
        except Exception as e:
            print(f"셀필드 값 채우기 실패: {e}")
            return False 