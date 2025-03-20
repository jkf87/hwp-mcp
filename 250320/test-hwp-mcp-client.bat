@echo off
setlocal enabledelayedexpansion

:: HWP MCP 클라이언트 테스트 스크립트
echo HWP MCP 클라이언트 테스트를 시작합니다...

:: Python 가상환경 활성화
if not exist venv (
    echo 가상환경이 존재하지 않습니다. 먼저 서버를 실행해주세요.
    exit /b 1
)

call venv\Scripts\activate

:: 테스트 실행
python test_hwp_mcp_client.py

:: 완료
echo HWP MCP 클라이언트 테스트가 완료되었습니다.
endlocal 