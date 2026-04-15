/**
 * Google Slides 전체 폰트 일괄 변경 스크립트
 *
 * 사용법:
 * 1. https://script.google.com/ → 새 프로젝트
 * 2. 이 파일 전체를 붙여넣기
 * 3. FILE_ID를 업로드한 Slides 파일 ID로 변경
 *    (Drive URL /presentation/d/<FILE_ID>/edit 에서 추출)
 * 4. 저장 → 실행 → 권한 허용
 * 5. 실행 로그 확인 후 Slides 파일 새로고침
 *
 * 처리 범위:
 * - SHAPE 내부 텍스트 (도형, 텍스트박스)
 * - TABLE 셀 텍스트
 * - GROUP 내부 요소 (재귀)
 * - 현재는 페이지 요소만 처리. 노트/마스터/레이아웃은 미처리.
 */

function replaceAllFonts() {
  // ====== 설정 ======
  const FILE_ID = 'REPLACE_WITH_YOUR_FILE_ID';
  const TARGET_FONT = 'Noto Sans KR';
  // ===================

  const pres = SlidesApp.openById(FILE_ID);
  const slides = pres.getSlides();
  let textCount = 0;
  let cellCount = 0;
  let groupDepth = 0;

  function applyFont(textRange) {
    try {
      if (textRange && textRange.asString().trim()) {
        textRange.getTextStyle().setFontFamily(TARGET_FONT);
        return true;
      }
    } catch (e) {
      Logger.log('텍스트 스타일 적용 실패: ' + e.message);
    }
    return false;
  }

  function processElement(el) {
    const t = el.getPageElementType();

    if (t === SlidesApp.PageElementType.SHAPE) {
      if (applyFont(el.asShape().getText())) textCount++;

    } else if (t === SlidesApp.PageElementType.TABLE) {
      const tbl = el.asTable();
      const rows = tbl.getNumRows();
      const cols = tbl.getNumColumns();
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          try {
            const cell = tbl.getCell(r, c);
            if (applyFont(cell.getText())) cellCount++;
          } catch (e) {
            // 병합된 셀 등은 skip
          }
        }
      }

    } else if (t === SlidesApp.PageElementType.GROUP) {
      groupDepth++;
      el.asGroup().getChildren().forEach(processElement);
      groupDepth--;
    }
    // WORD_ART, LINE, IMAGE, VIDEO, SHEETS_CHART 등은 텍스트 없음
  }

  const total = slides.length;
  slides.forEach((slide, i) => {
    slide.getPageElements().forEach(processElement);
    if ((i + 1) % 10 === 0 || i + 1 === total) {
      Logger.log('진행 ' + (i + 1) + '/' + total);
    }
  });

  Logger.log('=================================');
  Logger.log('완료');
  Logger.log('  SHAPE 텍스트 : ' + textCount);
  Logger.log('  TABLE 셀     : ' + cellCount);
  Logger.log('  적용 폰트    : ' + TARGET_FONT);
  Logger.log('=================================');
  Logger.log('Slides 파일 탭으로 돌아가서 F5로 새로고침하세요.');
}
