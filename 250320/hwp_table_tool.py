#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HWP 표 생성 및 조작 도구
pyhwpx를 이용한 표 생성, 커서 이동, 데이터 입력 기능을 제공합니다.
"""

import time
from typing import List, Tuple, Dict, Any, Optional, Union


class HwpTableTool:
    """
    한글 프로그램에서 표를 생성하고 조작하기 위한 도구 클래스
    커서 이동과 필드 조작을 함께 활용하여 표 작업을 자동화합니다.
    """

    def __init__(self, hwp_instance):
        """
        HwpTableTool 인스턴스를 초기화합니다.
        
        Args:
            hwp_instance: 한글 프로그램 인스턴스 (pyhwpx를 통해 생성된 객체)
        """
        self.hwp = hwp_instance
        self.current_table_pos = None  # 현재 작업 중인 표의 위치
        self.field_cache = {}  # 필드 이름과 위치를 캐싱하기 위한 딕셔너리

    def create_table(self, rows: int, cols: int, cell_width: int = None, cell_height: int = None) -> bool:
        """
        현재 커서 위치에 표를 생성합니다.
        
        Args:
            rows: 표의 행 수
            cols: 표의 열 수
            cell_width: 셀 너비 (None인 경우 기본값 사용)
            cell_height: 셀 높이 (None인 경우 기본값 사용)
            
        Returns:
            bool: 표 생성 성공 여부
        """
        try:
            # 현재 커서 위치 저장
            self.current_table_pos = self.hwp.get_pos()
            
            # 표 생성
            self.hwp.HAction.GetDefault("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            self.hwp.HParameterSet.HTableCreation.Rows = rows
            self.hwp.HParameterSet.HTableCreation.Cols = cols
            self.hwp.HParameterSet.HTableCreation.WidthType = 0  # 0: 단에 맞춤, 1: 절대값
            self.hwp.HParameterSet.HTableCreation.HeightType = 1  # 0: 자동, 1: 절대값
            
            # 셀 너비와 높이 설정
            if cell_width is None:
                cell_width = 8000 // cols  # 전체 너비를 열 수로 나눔
                
            if cell_height is None:
                cell_height = 1000  # 기본 높이
                
            self.hwp.HParameterSet.HTableCreation.WidthValue = cell_width
            self.hwp.HParameterSet.HTableCreation.HeightValue = cell_height
            
            # 각 열의 너비를 설정
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", cols)
            for i in range(cols):
                self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, cell_width)
                
            # 표 생성 실행
            self.hwp.HAction.Execute("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)
            
            # 생성된 표의 첫 번째 셀로 커서 이동
            table_list_idx = self.current_table_pos[0] + 2  # 표는 일반적으로 본문에서 List 인덱스가 2 증가
            self.move_to_cell(1, 1, table_list_idx)
            
            return True
        except Exception as e:
            print(f"표 생성 실패: {e}")
            return False

    def move_to_cell(self, row: int, col: int, table_list_idx: Optional[int] = None) -> bool:
        """
        표의 특정 셀로 커서를 이동합니다.
        
        Args:
            row: 이동할 행 번호 (1부터 시작)
            col: 이동할 열 번호 (1부터 시작)
            table_list_idx: 표의 List 인덱스 (None인 경우 현재 표 사용)
            
        Returns:
            bool: 이동 성공 여부
        """
        try:
            if table_list_idx is None and self.current_table_pos is not None:
                # 현재 작업 중인 표의 위치 사용
                table_list_idx = self.current_table_pos[0] + 2
            
            if table_list_idx is None:
                print("표의 위치를 알 수 없습니다.")
                return False
            
            # 표의 행과 열 번호는 0부터 시작하므로 1 감소
            row_idx = row - 1
            col_idx = col - 1
            
            # 표의 특정 셀로 이동
            # 각 셀은 고유한 List 인덱스를 가짐 (행과 열에 따라 계산)
            cell_list_idx = table_list_idx + row_idx * 2 + col_idx
            
            # 셀의 첫 번째 위치로 이동
            self.hwp.set_pos(cell_list_idx, 0, 0)
            
            return True
        except Exception as e:
            print(f"셀 이동 실패: {e}")
            return False

    def insert_text_to_cell(self, row: int, col: int, text: str, table_list_idx: Optional[int] = None) -> bool:
        """
        표의 특정 셀에 텍스트를 입력합니다.
        
        Args:
            row: 셀의 행 번호 (1부터 시작)
            col: 셀의 열 번호 (1부터 시작)
            text: 입력할 텍스트
            table_list_idx: 표의 List 인덱스 (None인 경우 현재 표 사용)
            
        Returns:
            bool: 입력 성공 여부
        """
        try:
            # 지정된 셀로 이동
            if not self.move_to_cell(row, col, table_list_idx):
                return False
            
            # 셀 내용 선택 (기존 내용 삭제를 위해)
            self.hwp.SelectAll()
            self.hwp.Delete()
            
            # 텍스트 입력
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            
            return True
        except Exception as e:
            print(f"셀에 텍스트 입력 실패: {e}")
            return False

    def fill_table_with_data(self, data: List[List[str]], start_row: int = 1, start_col: int = 1, 
                             has_header: bool = False, table_list_idx: Optional[int] = None) -> bool:
        """
        2차원 데이터로 표를 채웁니다.
        
        Args:
            data: 2차원 리스트 형태의 데이터 [[행1열1, 행1열2, ...], [행2열1, 행2열2, ...], ...]
            start_row: 시작 행 번호 (1부터 시작)
            start_col: 시작 열 번호 (1부터 시작)
            has_header: 첫 번째 행을 헤더로 처리할지 여부 (헤더인 경우 굵게 표시)
            table_list_idx: 표의 List 인덱스 (None인 경우 현재 표 사용)
            
        Returns:
            bool: 채우기 성공 여부
        """
        try:
            for row_idx, row_data in enumerate(data):
                for col_idx, cell_data in enumerate(row_data):
                    # 현재 셀 위치 계산
                    current_row = start_row + row_idx
                    current_col = start_col + col_idx
                    
                    # 셀에 데이터 입력
                    self.insert_text_to_cell(current_row, current_col, str(cell_data), table_list_idx)
                    
                    # 헤더 스타일 적용
                    if has_header and row_idx == 0:
                        # 현재 셀로 이동
                        self.move_to_cell(current_row, current_col, table_list_idx)
                        
                        # 셀 내용 선택
                        self.hwp.SelectAll()
                        
                        # 굵게 설정
                        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
                        self.hwp.HParameterSet.HCharShape.Bold = 1  # 굵게
                        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
                        
                        # 가운데 정렬
                        self.hwp.HAction.GetDefault("ParaShape", self.hwp.HParameterSet.HParaShape.HSet)
                        self.hwp.HParameterSet.HParaShape.Align = 1  # 가운데 정렬
                        self.hwp.HAction.Execute("ParaShape", self.hwp.HParameterSet.HParaShape.HSet)
            
            return True
        except Exception as e:
            print(f"표 데이터 채우기 실패: {e}")
            return False

    def create_field_in_cell(self, row: int, col: int, field_name: str, 
                             table_list_idx: Optional[int] = None) -> bool:
        """
        표의 특정 셀에 필드를 생성합니다.
        
        Args:
            row: 셀의 행 번호 (1부터 시작)
            col: 셀의 열 번호 (1부터 시작)
            field_name: 생성할 필드 이름
            table_list_idx: 표의 List 인덱스 (None인 경우 현재 표 사용)
            
        Returns:
            bool: 필드 생성 성공 여부
        """
        try:
            # 지정된 셀로 이동
            if not self.move_to_cell(row, col, table_list_idx):
                return False
            
            # 셀 내용 선택 (기존 내용 삭제를 위해)
            self.hwp.SelectAll()
            self.hwp.Delete()
            
            # 필드 생성
            self.hwp.HAction.GetDefault("FieldCreate", self.hwp.HParameterSet.HFieldCreate.HSet)
            self.hwp.HParameterSet.HFieldCreate.FieldName = field_name
            self.hwp.HParameterSet.HFieldCreate.FieldType = 0  # 0: 누름틀 필드
            self.hwp.HAction.Execute("FieldCreate", self.hwp.HParameterSet.HFieldCreate.HSet)
            
            # 필드 위치 캐싱
            current_pos = self.hwp.get_pos()
            self.field_cache[field_name] = {
                'row': row,
                'col': col,
                'pos': current_pos,
                'table_list_idx': table_list_idx
            }
            
            return True
        except Exception as e:
            print(f"셀에 필드 생성 실패: {e}")
            return False

    def fill_field(self, field_name: str, value: str) -> bool:
        """
        지정된 이름의 필드에 값을 채웁니다.
        
        Args:
            field_name: 필드 이름
            value: 채울 값
            
        Returns:
            bool: 필드 채우기 성공 여부
        """
        try:
            # 필드의 위치 정보가 캐시에 있으면 해당 위치로 이동
            if field_name in self.field_cache:
                field_info = self.field_cache[field_name]
                self.move_to_cell(field_info['row'], field_info['col'], field_info['table_list_idx'])
            
            # 필드 값 채우기
            self.hwp.PutFieldText(field_name, value)
            
            return True
        except Exception as e:
            print(f"필드 채우기 실패: {e}")
            return False

    def fill_fields_with_data(self, field_data: Dict[str, str]) -> bool:
        """
        여러 필드에 데이터를 한 번에 채웁니다.
        
        Args:
            field_data: 필드 이름과 값의 딕셔너리 {'필드이름': '값', ...}
            
        Returns:
            bool: 필드 채우기 성공 여부
        """
        try:
            for field_name, value in field_data.items():
                if not self.fill_field(field_name, value):
                    print(f"필드 '{field_name}' 채우기 실패")
            
            return True
        except Exception as e:
            print(f"필드 데이터 채우기 실패: {e}")
            return False

    def get_current_table_info(self) -> Tuple[int, int]:
        """
        현재 커서가 위치한 표의 행과 열 수를 반환합니다.
        
        Returns:
            Tuple[int, int]: (행 수, 열 수) 또는 표가 없는 경우 (0, 0)
        """
        try:
            self.hwp.HAction.GetDefault("TablePropertyDialog", self.hwp.HParameterSet.HShapeObject.HSet)
            rows = self.hwp.HParameterSet.HShapeObject.TreatAsChar
            cols = self.hwp.HParameterSet.HShapeObject.HeadingType
            return rows, cols
        except:
            return 0, 0

    def get_table_count(self) -> int:
        """
        문서 내 표의 개수를 반환합니다.
        
        Returns:
            int: 표의 개수
        """
        try:
            # 표 개수 확인을 위한 코드
            # 현재 커서 위치 저장
            original_pos = self.hwp.get_pos()
            
            # 문서 처음으로 이동
            self.hwp.set_pos(0, 0, 0)
            
            # 모든 컨트롤 검색
            count = 0
            while True:
                # 표 찾기
                found = self.hwp.HAction.GetDefault("TablePropertyDialog", self.hwp.HParameterSet.HShapeObject.HSet)
                if not found:
                    break
                
                count += 1
                
                # 다음 표로 이동
                self.hwp.HAction.Run("TableRightCell")
            
            # 원래 위치로 복귀
            self.hwp.set_pos(*original_pos)
            
            return count
        except Exception as e:
            print(f"표 개수 확인 실패: {e}")
            return 0

    def merge_cells(self, start_row: int, start_col: int, end_row: int, end_col: int, 
                    table_list_idx: Optional[int] = None) -> bool:
        """
        지정된 범위의 셀을 병합합니다.
        
        Args:
            start_row: 시작 행 번호 (1부터 시작)
            start_col: 시작 열 번호 (1부터 시작)
            end_row: 끝 행 번호 (1부터 시작)
            end_col: 끝 열 번호 (1부터 시작)
            table_list_idx: 표의 List 인덱스 (None인 경우 현재 표 사용)
            
        Returns:
            bool: 병합 성공 여부
        """
        try:
            # 시작 셀로 이동
            if not self.move_to_cell(start_row, start_col, table_list_idx):
                return False
            
            # 셀 선택 모드로 전환
            self.hwp.HAction.Run("TableCellBlock")
            
            # 끝 셀로 이동하면서 선택 영역 확장
            for _ in range(end_row - start_row):
                self.hwp.HAction.Run("TableLowerCell")
                
            for _ in range(end_col - start_col):
                self.hwp.HAction.Run("TableRightCell")
            
            # 셀 병합
            self.hwp.HAction.Run("TableCellBlockExtend")
            self.hwp.HAction.Run("TableCellBlock")
            self.hwp.HAction.Run("TableMergeCell")
            
            return True
        except Exception as e:
            print(f"셀 병합 실패: {e}")
            return False


# 사용 예제
def example_usage(hwp_instance):
    """
    HwpTableTool 사용 예제 함수
    
    Args:
        hwp_instance: 한글 프로그램 인스턴스
    """
    # 도구 인스턴스 생성
    table_tool = HwpTableTool(hwp_instance)
    
    # 새 문서 생성
    hwp_instance.Run("FileNew")
    
    # 3x3 표 생성
    table_tool.create_table(3, 3)
    
    # 표에 데이터 채우기
    data = [
        ["항목", "수량", "가격"],
        ["상품A", "2", "10,000원"],
        ["상품B", "1", "15,000원"]
    ]
    table_tool.fill_table_with_data(data, has_header=True)
    
    # 필드가 있는 표 생성 (새 문단 추가 후)
    hwp_instance.Run("BreakPara")
    hwp_instance.Run("BreakPara")
    table_tool.create_table(2, 2)
    
    # 셀에 필드 생성
    table_tool.create_field_in_cell(1, 1, "name")
    table_tool.create_field_in_cell(1, 2, "age")
    table_tool.create_field_in_cell(2, 1, "email")
    table_tool.create_field_in_cell(2, 2, "phone")
    
    # 필드에 데이터 채우기
    field_data = {
        "name": "홍길동",
        "age": "30",
        "email": "hong@example.com",
        "phone": "010-1234-5678"
    }
    table_tool.fill_fields_with_data(field_data)


if __name__ == "__main__":
    # pyhwpx 인스턴스를 직접 생성하여 테스트할 경우
    try:
        import win32com.client
        hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = True
        
        example_usage(hwp)
        print("예제 실행 완료")
    except Exception as e:
        print(f"예제 실행 실패: {e}") 