# HWP MCP (Message Control Protocol) Server

한글 프로그램(HWP)을 제어하기 위한 MCP(Message Control Protocol) 서버입니다.

## 개요

이 프로젝트는 한글 프로그램을 프로그래밍 방식으로 제어할 수 있는 MCP 서버를 제공합니다. 이를 통해 다양한 클라이언트 응용 프로그램에서 한글 문서 생성, 편집, 저장 등의 기능을 자동화할 수 있습니다.

## 주요 기능

- 새 한글 문서 생성
- 기존 한글 문서 열기
- 문서 저장
- 텍스트 삽입 및 서식 지정
- 표 삽입
- 단락 삽입
- 문서 텍스트 추출
- **pyhwpx 기반 표 생성 및 조작 도구 추가**

## 새로운 표 도구 기능

이 프로젝트에는 pyhwpx를 이용하여 표를 쉽게 생성하고 조작할 수 있는 도구를 추가했습니다:

- 표 생성 및 조작 (`hwp_table_tool.py`)
  - 표 생성 및 크기 조정
  - 특정 셀로 커서 이동
  - 셀에 텍스트 입력
  - 셀 병합
  - 표 스타일 설정
  
- 필드 관련 기능
  - 셀에 필드 생성
  - 필드 값 설정
  - 여러 필드를 한 번에 채우기

- 표 데이터 관리
  - 2차원 데이터로 표 채우기
  - 헤더 행 설정 (굵게, 가운데 정렬 등)
  
자세한 사용법은 `hwp_table_tool.py` 및 `test_hwp_table_tool.py` 파일을 참조하세요.

## 필수 요구사항

- Python 3.8 이상
- 한글(HWP) 프로그램 설치
- Windows 운영체제

## 설치 방법

1. 필요한 패키지 설치:
   ```
   pip install -r requirements.txt
   ```

2. 한글 프로그램이 설치되어 있는지 확인하세요.

## 사용 방법

### 서버 실행

```
run-hwp-mcp-server.bat
```

또는 직접 Python으로 실행:

```
python hwp_mcp_stdio_server.py
```

### 클라이언트 테스트

테스트 클라이언트를 실행하여 서버 기능을 확인할 수 있습니다:

```
test-hwp-mcp-client.bat
```

또는 직접 Python으로 실행:

```
python test_hwp_mcp_client.py
```

### 표 도구 테스트

표 생성 및 조작 도구 테스트:

```
python test_hwp_table_tool.py
```

## 파일 구조

- `hwp_mcp_stdio_server.py`: MCP 서버 메인 파일
- `src/tools/hwp_controller.py`: 한글 프로그램 제어 클래스
- `hwp_table_tool.py`: pyhwpx 기반 표 생성 및 조작 도구
- `test_hwp_table_tool.py`: 표 도구 테스트 스크립트
- `mcp/run.py`: JavaScript 실행 등 확장 기능 제공
- `requirements.txt`: 필요한 Python 패키지 목록
- `run-hwp-mcp-server.bat`: 서버 실행 스크립트
- `test-hwp-mcp-client.bat`: 클라이언트 테스트 스크립트
- `test_hwp_mcp_client.py`: 클라이언트 테스트 코드

## 라이센스

이 프로젝트는 MIT 라이센스를 따릅니다. 