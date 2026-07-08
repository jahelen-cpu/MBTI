/*
  배포 방법
  1) sheets.google.com 에서 새 스프레드시트 생성 (첫 행에 헤더는 자동 생성됨)
  2) 확장 프로그램 > Apps Script 메뉴 열기
  3) 기본 코드를 지우고 이 파일 내용 전체를 붙여넣기
  4) 스프레드시트 주소창의 URL에서 /d/ 와 /edit 사이의 긴 문자열을 복사해
     아래 SHEET_ID 값에 붙여넣기
  5) 배포 > 새 배포 > 유형: 웹 앱
     - 실행할 계정: 나
     - 액세스 권한이 있는 사용자: 모든 사용자
  6) 배포 후 나오는 웹 앱 URL을 index.html 상단의 GAS_URL 값에 붙여넣기
*/

const SHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID_HERE';
const HEADER = ['제출시각', '반', '학번', '이름', 'MBTI'];

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
  }
  return sheet;
}

function getAllRows_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(r => r[2])
    .map(r => ({
      timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      classNum: String(r[1]),
      studentId: String(r[2]),
      name: String(r[3]),
      mbti: String(r[4]),
    }));
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'createSlides') {
    try {
      const url = createSlidesDeck_(getAllRows_());
      return jsonOutput_({ ok: true, url });
    } catch (err) {
      return jsonOutput_({ ok: false, error: String(err) });
    }
  }

  const sheet = getSheet_();
  const classNum = String(data.classNum || '').trim();
  const studentId = String(data.studentId || '').trim();
  const name = String(data.name || '').trim();
  const mbti = String(data.mbti || '').trim().toUpperCase();

  if (!classNum || !studentId || !name || !mbti) {
    return jsonOutput_({ ok: false, error: 'missing fields' });
  }

  // 같은 반+학번이 이미 제출한 경우 새 값으로 덮어쓰기 (재제출 허용)
  const values = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === classNum && String(values[i][2]) === studentId) { rowIndex = i + 1; break; }
  }

  const row = [new Date(), classNum, studentId, name, mbti];
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return jsonOutput_({ ok: true });
}

function doGet(e) {
  return jsonOutput_({ ok: true, rows: getAllRows_() });
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 구글 슬라이드 생성 (Slides 서비스는 기본 제공, 별도 설정 불필요)
// ============================================================
const TYPE_GROUPS_ = [
  { label: '분석가', color: '#8E44AD', types: ['INTJ', 'INTP', 'ENTJ', 'ENTP'] },
  { label: '외교관', color: '#2E9E6B', types: ['INFJ', 'INFP', 'ENFJ', 'ENFP'] },
  { label: '관리자', color: '#2C6FBB', types: ['ISTJ', 'ISFJ', 'ESTJ', 'ESFJ'] },
  { label: '탐험가', color: '#D4A017', types: ['ISTP', 'ISFP', 'ESTP', 'ESFP'] },
];
const TYPE_NICK_ = {
  INTJ: '전략가', INTP: '논리술사', ENTJ: '통솔자', ENTP: '변론가',
  INFJ: '옹호자', INFP: '중재자', ENFJ: '선도자', ENFP: '활동가',
  ISTJ: '현실주의자', ISFJ: '수호자', ESTJ: '경영자', ESFJ: '집정관',
  ISTP: '장인', ISFP: '모험가', ESTP: '사업가', ESFP: '연예인',
};

function createSlidesDeck_(rows) {
  function countsFor(filterFn) {
    const counts = {};
    TYPE_GROUPS_.forEach(g => g.types.forEach(t => counts[t] = 0));
    rows.filter(filterFn).forEach(r => { if (counts.hasOwnProperty(r.mbti)) counts[r.mbti]++; });
    return counts;
  }

  function addSummarySlide(deck, label, counts, total) {
    const slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 960, 50).getFill().setSolidFill('#2E7D32');
    slide.insertTextBox(`${label}  ·  총 ${total}명`, 20, 5, 700, 40)
      .getText().getTextStyle().setFontSize(20).setBold(true).setForegroundColor('#FFFFFF');

    TYPE_GROUPS_.forEach((g, gi) => {
      g.types.forEach((t, ti) => {
        const x = 40 + ti * 220;
        const y = 90 + gi * 110;
        const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, 200, 90);
        box.getFill().setSolidFill(g.color);
        box.getBorder().setTransparent();
        const tb = slide.insertTextBox(`${t} · ${TYPE_NICK_[t]}\n${counts[t]}명`, x, y, 200, 90);
        tb.getText().getTextStyle().setFontSize(13).setBold(true).setForegroundColor('#FFFFFF');
        tb.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      });
    });
  }

  const deck = SlidesApp.create('MBTI 조사 결과 - ' + Utilities.formatDate(new Date(), 'GMT+9', 'yyyy.MM.dd'));
  const titleSlide = deck.getSlides()[0];
  titleSlide.getShapes().forEach(sh => { try { sh.remove(); } catch (err) {} });
  titleSlide.insertTextBox('🧭 우리 반 MBTI 조사 결과', 40, 180, 880, 70)
    .getText().getTextStyle().setFontSize(32).setBold(true);
  titleSlide.insertTextBox(`총 ${rows.length}명 응답 · 8개 학급`, 40, 260, 880, 40)
    .getText().getTextStyle().setFontSize(16).setForegroundColor('#5A705A');

  for (let c = 1; c <= 8; c++) {
    const classNum = String(c);
    const counts = countsFor(r => r.classNum === classNum);
    const total = Object.values(counts).reduce((a, b) => a + b, 0);
    addSummarySlide(deck, `${c}반`, counts, total);
  }

  addSummarySlide(deck, '전체 학급 합계', countsFor(() => true), rows.length);

  return deck.getUrl();
}

// ============================================================
// 스프레드시트 메뉴 (시트를 열면 자동으로 생김)
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MBTI 도구')
    .addItem('📊 결과 슬라이드 만들기', 'createSlidesFromMenu_')
    .addToUi();
}

function createSlidesFromMenu_() {
  const ui = SpreadsheetApp.getUi();
  try {
    const url = createSlidesDeck_(getAllRows_());
    const html = HtmlService
      .createHtmlOutput(`<p>슬라이드가 생성되었습니다.</p><p><a href="${url}" target="_blank">${url}</a></p>`)
      .setWidth(420)
      .setHeight(120);
    ui.showModalDialog(html, 'MBTI 결과 슬라이드');
  } catch (err) {
    ui.alert('오류가 발생했습니다: ' + err);
  }
}
