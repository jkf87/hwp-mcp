#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
MCP 서버를 위한 확장 기능 모듈
JavaScript 코드 실행과 같은 추가 기능을 제공합니다.
"""

import os
import sys
import json
import base64
import tempfile


def run_js(client, js_code):
    """
    MCP 클라이언트를 통해 JavaScript 코드를 HWP 서버에서 실행합니다.
    
    Args:
        client: MCP 클라이언트 인스턴스
        js_code: 실행할 JavaScript 코드
        
    Returns:
        str: 실행 결과
    """
    try:
        # JavaScript 코드를 실행하는 방법은 MCP 서버 구현에 따라 달라질 수 있음
        # 여기서는 hwp_execute_script라는 가상의 메서드를 사용한다고 가정
        
        if hasattr(client, 'hwp_execute_script'):
            # 직접 스크립트 실행 메서드가 있는 경우
            return client.hwp_execute_script(script=js_code)
        else:
            # 스크립트 실행 메서드가 없는 경우 대안으로 임시 파일을 통한 실행
            
            # 임시 JS 파일 생성
            fd, temp_path = tempfile.mkstemp(suffix='.js')
            try:
                with os.fdopen(fd, 'w', encoding='utf-8') as f:
                    f.write(js_code)
                
                # hwp_run_jscript 메서드가 있는지 확인
                if hasattr(client, 'hwp_run_jscript'):
                    return client.hwp_run_jscript(file_path=temp_path)
                else:
                    # 마지막 대안으로 hwp_exec 함수가 있다고 가정
                    if hasattr(client, 'hwp_exec'):
                        # 코드를 base64로 인코딩하여 직접 전달
                        encoded_js = base64.b64encode(js_code.encode('utf-8')).decode('utf-8')
                        return client.hwp_exec(command=f"RunJScript", parameter=encoded_js)
                    else:
                        return "Error: No JavaScript execution method available"
            finally:
                # 임시 파일 삭제
                try:
                    os.unlink(temp_path)
                except:
                    pass
    except Exception as e:
        return f"Error executing JavaScript: {str(e)}"


def add_tool_to_server(server, name, func):
    """
    MCP 서버에 새 도구를 동적으로 추가합니다.
    
    Args:
        server: MCP 서버 인스턴스
        name: 추가할 도구 이름
        func: 도구 함수
        
    Returns:
        bool: 추가 성공 여부
    """
    try:
        # 서버에 새 도구 등록
        server.add_tool(name, func)
        return True
    except Exception as e:
        print(f"Error adding tool to server: {str(e)}", file=sys.stderr)
        return False


def execute_hwp_action(hwp, action_name, parameter_set=None):
    """
    HWP 액션을 실행합니다.
    
    Args:
        hwp: HWP 인스턴스
        action_name: 실행할 액션 이름
        parameter_set: 액션에 필요한 파라미터 세트 (선택적)
        
    Returns:
        bool: 실행 성공 여부
    """
    try:
        if parameter_set:
            # 파라미터가 있는 경우
            hwp.HAction.GetDefault(action_name, parameter_set)
            return hwp.HAction.Execute(action_name, parameter_set)
        else:
            # 파라미터 없이 실행
            return hwp.HAction.Run(action_name)
    except Exception as e:
        print(f"Error executing HWP action '{action_name}': {str(e)}", file=sys.stderr)
        return False 