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

    def insert_text(self, text: str, preserve_linebreaks: bool = True) -> bool:
        """
        현재 커서 위치에 텍스트를 삽입합니다.
        
        Args:
            text (str): 삽입할 텍스트
            preserve_linebreaks (bool): 줄바꿈 유지 여부
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            if preserve_linebreaks and '\n' in text:
                # 줄바꿈이 포함된 경우 줄 단위로 처리
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    if i > 0:  # 첫 줄이 아니면 줄바꿈 추가
                        self.insert_paragraph()
                    if line.strip():  # 빈 줄이 아니면 텍스트 삽입
                        self._insert_text_direct(line)
                return True
            else:
                # 줄바꿈이 없거나 유지하지 않는 경우 한 번에 처리
                return self._insert_text_direct(text)
        except Exception as e:
            print(f"텍스트 삽입 실패: {e}")
            return False

    def _insert_text_direct(self, text: str) -> bool:
        """
        텍스트를 직접 삽입하는 내부 메서드입니다.
        
        Args:
            text (str): 삽입할 텍스트
            
        Returns:
            bool: 삽입 성공 여부
        """
        try:
            # 현재 위치 저장
            current_pos = self.hwp.GetPos()
            
            # 커서를 현재 위치의 시작점으로 이동 (이전 스타일의 영향을 받지 않도록)
            if current_pos:
                self.hwp.SetPos(*current_pos)
            
            # 텍스트 삽입을 위한 액션 초기화
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            
            return True
        except Exception as e:
            print(f"텍스트 직접 삽입 실패: {e}")
            return False

    def set_font(self, font_name: str, font_size: int, bold: bool = False, italic: bool = False, 
                select_previous_text: bool = False) -> bool:
        """
        글꼴 속성을 설정합니다. 현재 위치에서 다음에 입력할 텍스트에 적용됩니다.
        
        Args:
            font_name (str): 글꼴 이름
            font_size (int): 글꼴 크기
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            select_previous_text (bool): 이전에 입력한 텍스트를 선택할지 여부
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 새로운 구현: set_font_style 메서드 사용
            return self.set_font_style(
                font_name=font_name,
                font_size=font_size,
                bold=bold,
                italic=italic,
                underline=False,
                select_previous_text=select_previous_text
            )
        except Exception as e:
            print(f"글꼴 설정 실패: {e}")
            return False

    def set_font_style(self, font_name: str = None, font_size: int = None, 
                     bold: bool = False, italic: bool = False, underline: bool = False,
                     select_previous_text: bool = False) -> bool:
        """
        현재 선택된 텍스트의 글꼴 스타일을 설정합니다.
        선택된 텍스트가 없으면, 다음 입력될 텍스트에 적용됩니다.
        
        Args:
            font_name (str, optional): 글꼴 이름. None이면 현재 글꼴 유지.
            font_size (int, optional): 글꼴 크기. None이면 현재 크기 유지.
            bold (bool): 굵게 여부
            italic (bool): 기울임꼴 여부
            underline (bool): 밑줄 여부
            select_previous_text (bool): 이전에 입력한 텍스트를 선택할지 여부
            
        Returns:
            bool: 설정 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 이전 텍스트 선택 옵션이 활성화된 경우 현재 단락의 이전 텍스트 선택
            if select_previous_text:
                self.select_last_text()
            
            # 글꼴 설정을 위한 액션 초기화
            self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            
            # 글꼴 이름 설정
            if font_name:
                self.hwp.HParameterSet.HCharShape.FaceNameHangul = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameLatin = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameHanja = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameJapanese = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameOther = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameSymbol = font_name
                self.hwp.HParameterSet.HCharShape.FaceNameUser = font_name
            
            # 글꼴 크기 설정 (hwpunit, 10pt = 1000)
            if font_size:
                self.hwp.HParameterSet.HCharShape.Height = font_size * 100
            
            # 스타일 설정
            self.hwp.HParameterSet.HCharShape.Bold = bold
            self.hwp.HParameterSet.HCharShape.Italic = italic
            self.hwp.HParameterSet.HCharShape.UnderlineType = 1 if underline else 0
            
            # 변경사항 적용
            self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            
            return True
            
        except Exception as e:
            print(f"글꼴 스타일 설정 실패: {e}")
            return False

    def _get_current_position(self):
        """현재 커서 위치 정보를 가져옵니다."""
        try:
            # GetPos()는 현재 위치 정보를 (위치 유형, List ID, Para ID, CharPos)의 튜플로 반환
            return self.hwp.GetPos()
        except:
            # 실패 시 None 반환
            return None

    def _set_position(self, pos):
        """커서 위치를 지정된 위치로 변경합니다."""
        try:
            if pos:
                self.hwp.SetPos(*pos)
            return True
        except:
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

    def execute_script(self, script: str) -> str:
        """
        한글에서 JavaScript 코드를 실행합니다.
        
        Args:
            script (str): 실행할 JavaScript 코드
            
        Returns:
            str: 스크립트 실행 결과
        """
        try:
            if not self.is_hwp_running:
                return "Error: HWP is not running"
            
            # GetTextFile로 실행 및 결과 반환 (텍스트로 반환됨)
            result = self.hwp.GetTextFile("JS", script)
            return result if result else "Script executed without result"
        except Exception as e:
            error_msg = f"JavaScript 실행 실패: {e}"
            print(error_msg)
            return f"Error: {error_msg}"

    def select_last_text(self) -> bool:
        """
        현재 단락의 마지막으로 입력된 텍스트를 선택합니다.
        
        Returns:
            bool: 선택 성공 여부
        """
        try:
            if not self.is_hwp_running:
                return False
            
            # 현재 위치 저장
            current_pos = self.hwp.GetPos()
            if not current_pos:
                return False
                
            # 현재 단락의 시작으로 이동
            self.hwp.Run("MoveLineStart")
            start_pos = self.hwp.GetPos()
            
            # 이전 위치로 돌아가서 선택 영역 생성
            self.hwp.SetPos(*start_pos)
            self.hwp.SelectText(start_pos, current_pos)
            
            return True
        except Exception as e:
            print(f"텍스트 선택 실패: {e}")
            return False 