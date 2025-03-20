@echo off
setlocal enabledelayedexpansion

:: HWP-MCP 서버 구동 스크립트
echo HWP MCP 서버를 시작합니다...

:: Python 가상환경 준비
if exist venv (
    echo 기존 가상환경을 사용합니다.
    call venv\Scripts\activate
) else (
    echo 새 가상환경을 만듭니다.
    python -m venv venv
    call venv\Scripts\activate
    
    echo 필수 패키지를 설치합니다.
    pip install mcp pywin32 python-dotenv > pip_install.log 2>&1
    if %errorlevel% neq 0 (
        echo 패키지 설치 실패! pip_install.log 확인하세요.
        exit /b 1
    )
)

:: 서버 실행
echo HWP MCP 서버를 실행합니다...
python hwp_mcp_stdio_server.py

:: 완료
echo HWP MCP 서버가 종료되었습니다.
endlocal 