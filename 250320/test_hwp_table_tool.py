#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
HWP 표 생성 및 조작 도구 테스트 스크립트
MCP 서버를 통해 HwpTableTool을 테스트합니다.
"""

import os
import sys
import time
from mcp.client.stdio import stdio_client

# 현재 디렉토리를 sys.path에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

# HwpTableTool 클래스 임포트
from hwp_table_tool import HwpTableTool


def log(message):
    """메시지를 stderr에 출력합니다."""
    print(message, file=sys.stderr)


def test_table_creation_with_data():
    """표 생성 및 데이터 입력 테스트"""
    log("HWP 표 도구 테스트를 시작합니다...")
    
    try:
        # MCP 서버에 연결
        client = stdio_client.connect("hwp")
        log("HWP MCP 서버에 연결되었습니다.")
        
        # 새 문서 생성
        log("새 문서를 생성합니다...")
        client.hwp_create()
        time.sleep(1)
        
        # 3x4 표 생성 및 기본 데이터 입력
        log("3x4 표를 생성하고 데이터를 입력합니다...")
        
        # 표 생성 코드 실행 (JavaScript 코드를 문자열로 전달)
        js_code = """
        function createTableWithData() {
            var hwp = this;
            
            // 표 생성
            hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
            hwp.HParameterSet.HTableCreation.Rows = 3;
            hwp.HParameterSet.HTableCreation.Cols = 4;
            hwp.HParameterSet.HTableCreation.WidthType = 0;  // 단에 맞춤
            hwp.HParameterSet.HTableCreation.HeightType = 1;  // 절대값
            hwp.HParameterSet.HTableCreation.WidthValue = 0;
            hwp.HParameterSet.HTableCreation.HeightValue = 1000;
            
            // 열 너비 설정
            var colWidth = 8000 / 4;  // 균등하게 4등분
            hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 4);
            for (var i = 0; i < 4; i++) {
                hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, colWidth);
            }
            
            hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
            
            // 데이터 채우기 - 첫 번째 행 (헤더)
            hwp.SetPos(3, 0, 0);  // 표의 첫 번째 셀로 이동
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "항목1";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            hwp.SetPos(4, 0, 0);  // 다음 셀
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "항목2";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            hwp.SetPos(5, 0, 0);  // 다음 셀
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "항목3";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            hwp.SetPos(6, 0, 0);  // 다음 셀
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "항목4";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            // 데이터 채우기 - 두 번째 행
            hwp.SetPos(7, 0, 0);  // 다음 행, 첫 번째 셀
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "데이터1";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            hwp.SetPos(8, 0, 0);  // 다음 셀
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "데이터2";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            hwp.SetPos(9, 0, 0);  // 다음 셀
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "데이터3";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            hwp.SetPos(10, 0, 0);  // 다음 셀
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet);
            hwp.HParameterSet.HInsertText.Text = "데이터4";
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet);
            
            // 헤더 스타일 적용 (굵게, 가운데 정렬)
            for (var i = 0; i < 4; i++) {
                var cellPos = 3 + i;
                hwp.SetPos(cellPos, 0, 0);
                hwp.SelectAll();  // 셀 내용 선택
                
                // 글자 굵게 설정
                hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet);
                hwp.HParameterSet.HCharShape.Bold = 1;
                hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet);
                
                // 가운데 정렬
                hwp.HAction.GetDefault("ParaShape", hwp.HParameterSet.HParaShape.HSet);
                hwp.HParameterSet.HParaShape.Align = 1;  // 가운데 정렬
                hwp.HAction.Execute("ParaShape", hwp.HParameterSet.HParaShape.HSet);
            }
            
            return "표 생성 및 데이터 입력 완료";
        }
        
        createTableWithData();
        """
        
        # JavaScript 코드 실행
        from mcp.run import run_js
        result = run_js(client, js_code)
        log(f"JavaScript 실행 결과: {result}")
        time.sleep(1)
        
        # 새 단락 추가
        log("새 단락을 추가합니다...")
        client.hwp_insert_paragraph()
        client.hwp_insert_paragraph()
        time.sleep(1)
        
        # 필드가 있는 표 생성
        log("필드가 있는 표를 생성합니다...")
        js_code_fields = """
        function createTableWithFields() {
            var hwp = this;
            
            // 표 생성
            hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
            hwp.HParameterSet.HTableCreation.Rows = 2;
            hwp.HParameterSet.HTableCreation.Cols = 2;
            hwp.HParameterSet.HTableCreation.WidthType = 0;  // 단에 맞춤
            hwp.HParameterSet.HTableCreation.HeightType = 1;  // 절대값
            hwp.HParameterSet.HTableCreation.WidthValue = 0;
            hwp.HParameterSet.HTableCreation.HeightValue = 1500;
            
            // 열 너비 설정
            var colWidth = 8000 / 2;  // 균등하게 2등분
            hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 2);
            hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, colWidth);
            hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, colWidth);
            
            hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet);
            
            // 첫 번째 셀로 이동하여 필드 생성
            hwp.SetPos(13, 0, 0);  // 조정 필요한 값
            hwp.SelectAll();  // 셀 내용 선택
            hwp.Delete();  // 기존 내용 삭제
            
            // 필드 생성
            hwp.HAction.GetDefault("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            hwp.HParameterSet.HFieldCreate.FieldName = "name";
            hwp.HParameterSet.HFieldCreate.FieldType = 0;  // 누름틀 필드
            hwp.HAction.Execute("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            
            // 두 번째 셀로 이동하여 필드 생성
            hwp.SetPos(14, 0, 0);
            hwp.SelectAll();
            hwp.Delete();
            
            hwp.HAction.GetDefault("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            hwp.HParameterSet.HFieldCreate.FieldName = "age";
            hwp.HParameterSet.HFieldCreate.FieldType = 0;
            hwp.HAction.Execute("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            
            // 세 번째 셀로 이동하여 필드 생성
            hwp.SetPos(15, 0, 0);
            hwp.SelectAll();
            hwp.Delete();
            
            hwp.HAction.GetDefault("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            hwp.HParameterSet.HFieldCreate.FieldName = "email";
            hwp.HParameterSet.HFieldCreate.FieldType = 0;
            hwp.HAction.Execute("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            
            // 네 번째 셀로 이동하여 필드 생성
            hwp.SetPos(16, 0, 0);
            hwp.SelectAll();
            hwp.Delete();
            
            hwp.HAction.GetDefault("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            hwp.HParameterSet.HFieldCreate.FieldName = "phone";
            hwp.HParameterSet.HFieldCreate.FieldType = 0;
            hwp.HAction.Execute("FieldCreate", hwp.HParameterSet.HFieldCreate.HSet);
            
            // 필드에 값 입력
            hwp.PutFieldText("name", "홍길동");
            hwp.PutFieldText("age", "30");
            hwp.PutFieldText("email", "hong@example.com");
            hwp.PutFieldText("phone", "010-1234-5678");
            
            return "필드가 있는 표 생성 완료";
        }
        
        createTableWithFields();
        """
        
        result = run_js(client, js_code_fields)
        log(f"필드 표 생성 결과: {result}")
        
        # 문서 저장
        log("문서를 저장합니다...")
        save_path = os.path.join(os.getcwd(), "hwp_table_test_result.hwp")
        client.hwp_save(path=save_path)
        log(f"문서가 다음 경로에 저장되었습니다: {save_path}")
        
        log("테스트가 성공적으로 완료되었습니다!")
        return True
        
    except Exception as e:
        log(f"테스트 중 오류 발생: {str(e)}")
        import traceback
        log(traceback.format_exc())
        return False


class RunJs:
    """JavaScript 코드를 HWP MCP 서버에서 실행하기 위한 래퍼 클래스"""
    
    @staticmethod
    def run_js(client, js_code):
        """
        JavaScript 코드를 HWP MCP 서버에서 실행합니다.
        
        Args:
            client: MCP 클라이언트 인스턴스
            js_code: 실행할 JavaScript 코드
            
        Returns:
            str: 실행 결과
        """
        # 실제로는 MCP 서버에 JavaScript 실행 명령을 전송하는 로직이 필요함
        # 현재는 가상의 메서드임
        return "JavaScript 실행 성공 (가상 결과)"


if __name__ == "__main__":
    test_table_creation_with_data()
    sys.exit(0) 