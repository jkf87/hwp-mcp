/**
 * 한글(HWP) MCP 클라이언트 테스트
 */

const { MCPClient } = require('mcp-toolkit');
const path = require('path');

// 로깅 함수
function log(message) {
  console.error(`[${new Date().toISOString()}] ${message}`);
}

async function main() {
  log('HWP MCP 클라이언트 테스트 시작...');
  
  try {
    // MCP 서버에 연결 (서버가 실행 중이어야 함)
    const client = await MCPClient.connect({
      name: 'hwp-mcp',
      transport: 'stdio'
    });
    
    log('HWP MCP 서버에 연결되었습니다');
    
    // 새 문서 생성 테스트
    log('테스트: 새 문서 생성');
    const createResult = await client.call('hwp_create');
    log(`결과: ${createResult}`);
    await sleep(1000);
    
    // 텍스트 삽입 테스트
    log('테스트: 텍스트 삽입');
    const textResult = await client.call('hwp_insert_text', {
      text: '안녕하세요! HWP MCP 테스트입니다.'
    });
    log(`결과: ${textResult}`);
    await sleep(1000);
    
    // 글꼴 설정 테스트
    log('테스트: 글꼴 설정');
    const fontResult = await client.call('hwp_set_font', {
      name: '맑은 고딕',
      size: 12,
      bold: true
    });
    log(`결과: ${fontResult}`);
    await sleep(1000);
    
    // 단락 추가 테스트
    log('테스트: 단락 추가');
    const paraResult = await client.call('hwp_insert_paragraph');
    log(`결과: ${paraResult}`);
    await sleep(1000);
    
    // 표 삽입 테스트
    log('테스트: 표 삽입');
    const tableResult = await client.call('hwp_insert_table', {
      rows: 3,
      cols: 3
    });
    log(`결과: ${tableResult}`);
    await sleep(1000);
    
    // 텍스트 가져오기 테스트
    log('테스트: 텍스트 가져오기');
    const getTextResult = await client.call('hwp_get_text');
    log(`결과: ${getTextResult.substring(0, 100)}...`);  // 처음 100자만 표시
    await sleep(1000);
    
    // 문서 저장 테스트
    log('테스트: 문서 저장');
    const testFilePath = path.join(process.cwd(), 'hwp_mcp_test_result.hwp');
    const saveResult = await client.call('hwp_save', {
      path: testFilePath
    });
    log(`결과: ${saveResult}`);
    
    log('모든 테스트가 성공적으로 완료되었습니다!');
    log(`테스트 문서 저장 위치: ${testFilePath}`);
    
    // 문서는 열어둔 상태로 종료 (수동으로 닫기 위해)
    log('테스트 완료. 한글 문서는 열린 상태로 유지됩니다.');
    
    return 0;
  } catch (e) {
    log(`테스트 중 오류 발생: ${e.message}`);
    console.error(e.stack);
    return 1;
  }
}

// 지정한 시간(ms) 동안 대기하는 함수
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 메인 함수 실행
main().then(exitCode => {
  process.exit(exitCode);
}).catch(err => {
  console.error('예상치 못한 오류:', err);
  process.exit(1);
}); 