/**
 * 한글(HWP) MCP 서버
 * Node.js로 구현된 Model Context Protocol 서버
 */

const { MCPServer } = require('mcp-toolkit');
const fs = require('fs');
const path = require('path');
const HwpController = require('./hwp_controller');

// 로깅 설정
function log(message) {
  console.error(`[${new Date().toISOString()}] ${message}`);
  fs.appendFileSync('hwp_mcp_node.log', `${new Date().toISOString()} - ${message}\n`);
}

// 글로벌 HwpController 인스턴스
let hwpController = null;

// HwpController 인스턴스 가져오기
function getHwpController() {
  if (!hwpController) {
    log('HwpController 인스턴스 생성 중...');
    try {
      hwpController = new HwpController();
      if (!hwpController.connect(true)) {
        log('한글 프로그램 연결 실패');
        return null;
      }
      log('한글 프로그램 연결 성공');
    } catch (e) {
      log(`HwpController 인스턴스 생성 오류: ${e.message}`);
      return null;
    }
  }
  return hwpController;
}

// MCP 서버 인스턴스 생성
const server = new MCPServer({
  name: 'hwp',
  version: '0.1.0',
  description: 'HWP MCP Server for controlling Hangul Word Processor',
  transport: 'stdio',
});

// 도구 정의 및 등록
server.defineTools({
  // 새 문서 생성
  hwp_create: {
    description: '새 한글 문서를 생성합니다.',
    parameters: {},
    handler: async () => {
      try {
        const hwp = getHwpController();
        if (!hwp) {
          return '한글 프로그램 연결 실패';
        }
        
        if (hwp.createNewDocument()) {
          log('새 문서 생성 성공');
          return '새 문서가 생성되었습니다';
        } else {
          log('새 문서 생성 실패');
          return '새 문서 생성 실패';
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  },
  
  // 문서 열기
  hwp_open: {
    description: '기존 한글 문서를 엽니다.',
    parameters: {
      path: {
        type: 'string',
        description: '열 문서의 경로',
        required: true
      }
    },
    handler: async ({ path: filePath }) => {
      try {
        if (!filePath) {
          return '파일 경로가 필요합니다';
        }
        
        const hwp = getHwpController();
        if (!hwp) {
          return '한글 프로그램 연결 실패';
        }
        
        if (hwp.openDocument(filePath)) {
          log(`문서 열기 성공 - ${filePath}`);
          return `문서를 열었습니다: ${filePath}`;
        } else {
          log(`문서 열기 실패 - ${filePath}`);
          return '문서 열기 실패';
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  },
  
  // 문서 저장
  hwp_save: {
    description: '현재 문서를 저장합니다.',
    parameters: {
      path: {
        type: 'string',
        description: '저장할 경로 (선택사항)',
        required: false
      }
    },
    handler: async ({ path: filePath }) => {
      try {
        const hwp = getHwpController();
        if (!hwp) {
          return '한글 프로그램 연결 실패';
        }
        
        if (filePath) {
          if (hwp.saveDocument(filePath)) {
            log(`문서 저장 성공 - ${filePath}`);
            return `문서가 저장되었습니다: ${filePath}`;
          } else {
            log(`문서 저장 실패 - ${filePath}`);
            return '문서 저장 실패';
          }
        } else {
          const tempPath = path.join(process.cwd(), 'temp_document.hwp');
          if (hwp.saveDocument(tempPath)) {
            log(`문서 임시 저장 성공 - ${tempPath}`);
            return `문서가 임시 저장되었습니다: ${tempPath}`;
          } else {
            log('문서 임시 저장 실패');
            return '문서 임시 저장 실패';
          }
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  },
  
  // 텍스트 삽입
  hwp_insert_text: {
    description: '문서에 텍스트를 삽입합니다.',
    parameters: {
      text: {
        type: 'string',
        description: '삽입할 텍스트',
        required: true
      }
    },
    handler: async ({ text }) => {
      try {
        if (!text) {
          return '텍스트가 필요합니다';
        }
        
        const hwp = getHwpController();
        if (!hwp) {
          return '한글 프로그램 연결 실패';
        }
        
        if (hwp.insertText(text)) {
          log('텍스트 삽입 성공');
          return '텍스트가 삽입되었습니다';
        } else {
          log('텍스트 삽입 실패');
          return '텍스트 삽입 실패';
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  },
  
  // 글꼴 설정
  hwp_set_font: {
    description: '글꼴을 설정합니다.',
    parameters: {
      name: {
        type: 'string',
        description: '글꼴 이름',
        required: false
      },
      size: {
        type: 'number',
        description: '글꼴 크기',
        required: false
      },
      bold: {
        type: 'boolean',
        description: '굵게',
        required: false
      },
      italic: {
        type: 'boolean',
        description: '기울임꼴',
        required: false
      }
    },
    handler: async ({ name, size, bold, italic }) => {
      try {
        const hwp = getHwpController();
        if (!hwp) {
          return '한글 프로그램 연결 실패';
        }
        
        if (hwp.setFont(name, size, bold, italic)) {
          log('글꼴 설정 성공');
          return '글꼴이 설정되었습니다';
        } else {
          log('글꼴 설정 실패');
          return '글꼴 설정 실패';
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  },
  
  // 표 삽입
  hwp_insert_table: {
    description: '표를 삽입합니다.',
    parameters: {
      rows: {
        type: 'number',
        description: '행 수',
        required: true
      },
      cols: {
        type: 'number',
        description: '열 수',
        required: true
      }
    },
    handler: async ({ rows, cols }) => {
      try {
        const hwp = getHwpController();
        if (!hwp) {
          return '한글 프로그램 연결 실패';
        }
        
        if (hwp.insertTable(rows, cols)) {
          log(`${rows}x${cols} 표 삽입 성공`);
          return `${rows}x${cols} 크기의 표가 삽입되었습니다`;
        } else {
          log('표 삽입 실패');
          return '표 삽입 실패';
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  },
  
  // 단락 추가
  hwp_insert_paragraph: {
    description: '새 단락을 추가합니다.',
    parameters: {},
    handler: async () => {
      try {
        const hwp = getHwpController();
        if (!hwp) {
          return '한글 프로그램 연결 실패';
        }
        
        if (hwp.insertParagraph()) {
          log('단락 추가 성공');
          return '단락이 추가되었습니다';
        } else {
          log('단락 추가 실패');
          return '단락 추가 실패';
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  },
  
  // 문서 텍스트 가져오기
  hwp_get_text: {
    description: '문서의 텍스트를 가져옵니다.',
    parameters: {},
    handler: async () => {
      try {
        const hwp = getHwpController();
        if (!hwp) {
          return '한글 프로그램 연결 실패';
        }
        
        const text = hwp.getText();
        if (text !== null) {
          log('문서 텍스트 가져오기 성공');
          return text;
        } else {
          log('문서 텍스트 가져오기 실패');
          return '문서 텍스트 가져오기 실패';
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  },
  
  // 문서 닫기
  hwp_close: {
    description: '한글 문서를 닫고 연결을 종료합니다.',
    parameters: {
      save: {
        type: 'boolean',
        description: '저장 여부',
        required: false
      }
    },
    handler: async ({ save = true }) => {
      try {
        if (hwpController && hwpController.isHwpRunning) {
          if (hwpController.disconnect()) {
            log('한글 프로그램 연결 종료 성공');
            hwpController = null;
            return '한글 프로그램 연결이 종료되었습니다';
          } else {
            log('한글 프로그램 연결 종료 실패');
            return '한글 프로그램 연결 종료 실패';
          }
        } else {
          return '이미 한글 프로그램이 종료되었습니다';
        }
      } catch (e) {
        log(`오류: ${e.message}`);
        return `오류: ${e.message}`;
      }
    }
  }
});

// 서버 시작
log('HWP MCP 서버 시작');
server.start();

// 종료 처리
process.on('exit', () => {
  log('서버 종료 중...');
  if (hwpController && hwpController.isHwpRunning) {
    hwpController.disconnect();
    log('한글 프로그램 연결 종료');
  }
}); 