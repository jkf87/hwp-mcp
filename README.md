# HWP MCP Server (Node.js)

한글(HWP) 문서를 제어하기 위한 Model Context Protocol(MCP) 서버입니다. 이 서버는 Node.js로 구현되었으며, Claude AI와 같은 AI 어시스턴트가 한글 문서를 조작할 수 있도록 합니다.

## 주요 기능

- 한글 문서 생성 및 열기
- 텍스트 삽입 및 편집
- 글꼴 설정 변경
- 표 삽입
- 문서 저장
- 그 외 다양한 한글 문서 편집 기능

## 설치 방법

### 전제 조건

- Node.js v12 이상 설치
- 한글(HWP) 프로그램 설치
- Windows 운영체제 (COM 인터페이스 사용)

### 설치 단계

1. 이 저장소를 클론 또는 다운로드합니다.

2. 필요한 NPM 패키지를 설치합니다:

```bash
cd hwp-mcp-node
npm install
```

3. 만약 `winax` 설치에 문제가 있다면, 다음과 같이 변경하여 설치해보세요:

```bash
npm install --save-optional winax
```

## 사용 방법

### 서버 실행

다음 명령어를 사용하여 서버를 실행합니다:

```bash
npm start
```

또는 제공된 배치 파일을 실행할 수도 있습니다:

```bash
run-hwp-mcp-server.cmd
```

### 클라이언트 테스트

서버가 실행 중인 상태에서 테스트 클라이언트를 실행하여 기능을 확인할 수 있습니다:

```bash
node test_hwp_mcp_client.js
```

또는 제공된 배치 파일을 실행할 수도 있습니다:

```bash
test-hwp-mcp-client.cmd
```

## Claude AI 설정

Claude AI와 함께 사용하려면 다음과 같이 설정합니다:

1. Claude Desktop 앱을 설치합니다.
2. `%APPDATA%\Claude` 폴더에 있는 `claude_desktop_config.json` 파일을 열거나 생성합니다.
3. 다음 설정을 추가합니다:

```json
{
  "mcpServers": {
    "hwp": {
      "command": "node",
      "args": ["D:\\경로\\hwp-mcp-node\\hwp_mcp_server.js"]
    }
  }
}
```

* 경로 부분은 실제 hwp_mcp_server.js 파일의 전체 경로로 변경하세요.

## API 사용 예시

### 새 문서 만들기
```javascript
client.call('hwp_create');
```

### 텍스트 삽입
```javascript
client.call('hwp_insert_text', { text: '안녕하세요! HWP MCP 테스트입니다.' });
```

### 글꼴 설정
```javascript
client.call('hwp_set_font', { name: '맑은 고딕', size: 12, bold: true });
```

### 표 삽입
```javascript
client.call('hwp_insert_table', { rows: 3, cols: 3 });
```

### 문서 저장
```javascript
client.call('hwp_save', { path: 'D:\\경로\\문서.hwp' });
```

## 문제 해결

### 자주 발생하는 문제

1. **winax 모듈 설치 실패**
   
   해결방법: Visual Studio 빌드 도구를 설치하거나, 다음과 같이 선택적 종속성으로 설치합니다:
   ```
   npm install --save-optional winax
   ```

2. **한글 프로그램 연결 실패**
   
   해결방법: 한글 프로그램이 설치되어 있는지 확인하고, 관리자 권한으로 실행해보세요.

3. **글꼴 설정 후 텍스트 덮어쓰기 문제**
   
   해결방법: 최신 버전에서는 이 문제가 수정되었습니다. 업데이트하거나 소스코드를 확인하세요.

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 