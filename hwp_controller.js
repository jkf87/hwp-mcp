/**
 * 한글(HWP) 문서를 제어하기 위한 컨트롤러 모듈
 * Windows COM을 통해 한글 프로그램을 자동화합니다.
 */

const winax = require('winax');

class HwpController {
  constructor() {
    this.hwp = null;
    this.visible = true;
    this.isHwpRunning = false;
    this.currentDocumentPath = null;
  }

  /**
   * 한글 프로그램에 연결합니다.
   * @param {boolean} visible 한글 창을 화면에 표시할지 여부
   * @returns {boolean} 연결 성공 여부
   */
  connect(visible = true) {
    try {
      this.hwp = new winax.Object('HWPFrame.HwpObject');
      
      this.visible = visible;
      if (this.hwp.XHwpWindows && this.hwp.XHwpWindows.Item) {
        this.hwp.XHwpWindows.Item(0).Visible = visible;
      }
      
      this.isHwpRunning = true;
      return true;
    } catch (e) {
      console.error(`한글 프로그램 연결 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 한글 프로그램 연결을 종료합니다.
   * @returns {boolean} 종료 성공 여부
   */
  disconnect() {
    try {
      if (this.isHwpRunning) {
        // COM 객체 해제
        this.hwp = null;
        this.isHwpRunning = false;
      }
      return true;
    } catch (e) {
      console.error(`한글 프로그램 종료 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 새 문서를 생성합니다.
   * @returns {boolean} 생성 성공 여부
   */
  createNewDocument() {
    try {
      if (!this.isHwpRunning) {
        this.connect();
      }
      
      this.hwp.Run('FileNew');
      this.currentDocumentPath = null;
      return true;
    } catch (e) {
      console.error(`새 문서 생성 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 문서를 엽니다.
   * @param {string} filePath 열 문서의 경로
   * @returns {boolean} 열기 성공 여부
   */
  openDocument(filePath) {
    try {
      if (!this.isHwpRunning) {
        this.connect();
      }
      
      this.hwp.Open(filePath);
      this.currentDocumentPath = filePath;
      return true;
    } catch (e) {
      console.error(`문서 열기 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 문서를 저장합니다.
   * @param {string} filePath 저장할 경로. 없으면 현재 경로에 저장.
   * @returns {boolean} 저장 성공 여부
   */
  saveDocument(filePath = null) {
    try {
      if (!this.isHwpRunning) {
        return false;
      }
      
      if (filePath) {
        // 파일 형식과 경로 모두 지정하여 저장
        this.hwp.SaveAs(filePath, 'HWP', '');
        this.currentDocumentPath = filePath;
      } else {
        if (this.currentDocumentPath) {
          this.hwp.Save();
        } else {
          // 저장 대화 상자 표시
          this.hwp.SaveAs();
        }
      }
      
      return true;
    } catch (e) {
      console.error(`문서 저장 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 현재 커서 위치에 텍스트를 삽입합니다.
   * @param {string} text 삽입할 텍스트
   * @returns {boolean} 삽입 성공 여부
   */
  insertText(text) {
    try {
      if (!this.isHwpRunning) {
        return false;
      }
      
      this.hwp.HAction.GetDefault('InsertText', this.hwp.HParameterSet.HInsertText.HSet);
      this.hwp.HParameterSet.HInsertText.Text = text;
      this.hwp.HAction.Execute('InsertText', this.hwp.HParameterSet.HInsertText.HSet);
      return true;
    } catch (e) {
      console.error(`텍스트 삽입 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 현재 선택된 텍스트의 글꼴을 설정합니다.
   * @param {string} fontName 글꼴 이름
   * @param {number} fontSize 글꼴 크기
   * @param {boolean} bold 굵게 여부
   * @param {boolean} italic 기울임꼴 여부
   * @returns {boolean} 설정 성공 여부
   */
  setFont(fontName, fontSize, bold = false, italic = false) {
    try {
      if (!this.isHwpRunning) {
        return false;
      }
      
      // 전체 선택
      this.hwp.Run('SelectAll');
      
      // 글꼴 설정
      const sizePt = fontSize * 100; // 폰트 크기 단위 변환
      const boldValue = bold ? '1' : '0';
      const italicValue = italic ? '1' : '0';
      
      // 매크로 명령 실행
      this.hwp.Run(`CharShape "${fontName}" ${sizePt} ${boldValue} ${italicValue} 0 0 "" 0 "" 0`);
      
      return true;
    } catch (e) {
      console.error(`글꼴 설정 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 현재 커서 위치에 표를 삽입합니다.
   * @param {number} rows 행 수
   * @param {number} cols 열 수
   * @returns {boolean} 삽입 성공 여부
   */
  insertTable(rows, cols) {
    try {
      if (!this.isHwpRunning) {
        return false;
      }
      
      this.hwp.HAction.GetDefault('TableCreate', this.hwp.HParameterSet.HTableCreation.HSet);
      this.hwp.HParameterSet.HTableCreation.Rows = rows;
      this.hwp.HParameterSet.HTableCreation.Cols = cols;
      this.hwp.HParameterSet.HTableCreation.WidthType = 0; // 0: 단에 맞춤, 1: 절대값
      this.hwp.HParameterSet.HTableCreation.HeightType = 1; // 0: 자동, 1: 절대값
      this.hwp.HParameterSet.HTableCreation.WidthValue = 0; // 단에 맞춤이므로 무시됨
      this.hwp.HParameterSet.HTableCreation.HeightValue = 1000; // 셀 높이
      
      // 각 열의 너비를 설정 (모두 동일하게)
      const colWidth = 8000 / cols; // 전체 너비를 열 수로 나눔
      this.hwp.HParameterSet.HTableCreation.CreateItemArray('ColWidth', cols);
      
      for (let i = 0; i < cols; i++) {
        this.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(i, colWidth);
      }
      
      this.hwp.HAction.Execute('TableCreate', this.hwp.HParameterSet.HTableCreation.HSet);
      return true;
    } catch (e) {
      console.error(`표 삽입 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 새 단락을 추가합니다.
   * @returns {boolean} 추가 성공 여부
   */
  insertParagraph() {
    try {
      if (!this.isHwpRunning) {
        return false;
      }
      
      this.hwp.Run('BreakPara');
      return true;
    } catch (e) {
      console.error(`단락 추가 실패: ${e.message}`);
      return false;
    }
  }

  /**
   * 문서의 텍스트를 가져옵니다.
   * @returns {string} 문서 텍스트
   */
  getText() {
    try {
      if (!this.isHwpRunning) {
        return null;
      }
      
      const text = this.hwp.GetTextFile('TEXT', '');
      return text;
    } catch (e) {
      console.error(`텍스트 가져오기 실패: ${e.message}`);
      return null;
    }
  }
}

module.exports = HwpController; 