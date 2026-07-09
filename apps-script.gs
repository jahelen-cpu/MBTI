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
const ROSTER_SHEET_NAME = '전체명단';
const ROSTER_HEADER = ['반', '번호', '이름'];

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

// 반별 전체 학생 명단 시트 (미제출자 확인용). 없으면 헤더만 있는 빈 시트를 만들어둔다.
function getRosterSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ROSTER_SHEET_NAME);
    sheet.appendRow(ROSTER_HEADER);
  }
  return sheet;
}

function getAllRoster_() {
  const sheet = getRosterSheet_();
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(r => r[1])
    .map(r => ({
      classNum: String(r[0]),
      studentId: String(r[1]),
      name: String(r[2]),
    }));
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'createSlides') {
    try {
      const classFilter = data.classFilter ? String(data.classFilter) : 'all';
      const url = createSlidesDeck_(getAllRows_(), classFilter);
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
  return jsonOutput_({ ok: true, rows: getAllRows_(), roster: getAllRoster_() });
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
const TYPE_EMOJI_ = {
  INTJ: '🦉', INTP: '🧠', ENTJ: '🦁', ENTP: '🦊',
  INFJ: '🌙', INFP: '🦋', ENFJ: '🌻', ENFP: '🌟',
  ISTJ: '🐢', ISFJ: '🐰', ESTJ: '🦅', ESFJ: '🐝',
  ISTP: '🔧', ISFP: '🎨', ESTP: '🐆', ESFP: '🎉',
};
const TYPE_CAREERS_ = {
  INTJ: ['분석가', '회계사', '인류학자', '파일럿', '경영 컨설턴트', '제약회사 연구원', '웹 개발자', '최고 재무 책임자'],
  INTP: ['경제학자', '심리학자', '경찰', '프로그래머', '천문학자', '비평가', '아트디렉터', '연구원'],
  ENTJ: ['경영 컨설턴트', '공인중개사', '관리자', '변호사', '재무 상담사', '경제 분석가', '벤처 투자가', '판사'],
  ENTP: ['발명가', '벤처 사업가', '에이전트', '배우', '가수', '영화감독', '칼럼니스트', '정치인'],
  INFJ: ['직업상담사', '특수교사', '노인복지사', '아트 디렉터', '프리랜서 기획자', '저널리스트', '상품기획 MD'],
  INFP: ['예술가', '소설가', '시인', '음악가', '미술치료사', '사회복지사', '작곡가', '사서'],
  ENFJ: ['아나운서', '리포터', '방송 MC', '언어교사', '아동복지사', 'CEO', '취업 컨설턴트', '동시통역가'],
  ENFP: ['크리에이티브 디렉터', '디자이너', '시나리오 작가', '방송 프로듀서', '홍보 컨설턴트', '상담사', '상품 기획자'],
  ISTJ: ['통계학자', '바이어', '기상학자', '법률연구원', '보험 심사관', '형사', '감정평가사', '세관조사관'],
  ISFJ: ['행정보조원', '인사관리자', '신용상담가', '보호감찰관', '물리치료사', '정신과의사', '방사선기사'],
  ESTJ: ['감독관', '예산분석가', '은행장', '정책책임자', '보안요원', '기관사', '교육전문가'],
  ESFJ: ['홍보책임자', '호텔지배인', '마케팅책임자', '초등학교교사', '특수교사', '비서', '유치원교사'],
  ISTP: ['파일럿', '카레이서', '범죄학자', '사진작가', '판매원', '운동선수', '항공기정비사', '네트워크관리자'],
  ISFP: ['보석세공사', '음향디자이너', '만화가', '지질학자', '사육사', '수의사', '법률비서', '약사'],
  ESTP: ['경찰관', '소방관', '군장교', '펀드매니저', '은행원', '기자', '여행가이드', '건축엔지니어'],
  ESFP: ['코미디언', '의상디자이너', '일러스트레이터', '애니메이터', '여행상품기획자', '놀이치료사'],
};
const TYPE_DESC_ = {
  INTJ: '창의적이고 독립적이며, 장기적 계획을 세우고 효율적으로 목표를 달성하려는 전략적 성격입니다.',
  INTP: '이론적이고 분석적이며, 지적 호기심이 강하고 독립적으로 복잡한 문제를 탐구하는 성격입니다.',
  ENTJ: '지도력 있고 결단력 있으며, 효율적인 체계를 구축하고 목표 달성을 위해 주도적으로 행동하는 성격입니다.',
  ENTP: '논쟁적이고 혁신적이며, 도전 과제를 분석하고 다양한 관점에서 문제를 해결하려는 성격입니다.',
  INFJ: '이상과 의미를 추구하며 타인의 성장을 돕고자 하고, 깊은 통찰력과 확신 있는 신념을 가진 성격입니다.',
  INFP: '이상주의적이고 개인의 신념을 따르며, 타인을 진심으로 돕고 창의적 표현을 추구하는 성격입니다.',
  ENFJ: '리더십 있고 영감을 주며, 타인의 성장을 도모하고 공동의 목표를 위해 헌신하는 성격입니다.',
  ENFP: '열정적이고 창의적이며, 가능성을 탐색하고 사람들을 독려하며 새로운 경험을 추구하는 성격입니다.',
  ISTJ: '책임감 있고 신뢰할 수 있으며, 전통과 규칙을 중시하고 체계적으로 일을 처리하는 성격입니다.',
  ISFJ: '따뜻하고 배려심 깊으며, 타인의 필요를 먼저 챙기고 조화로운 관계를 중요시하는 성격입니다.',
  ESTJ: '조직적이고 리더십 있으며, 질서를 유지하고 효율적으로 목표를 달성하려는 책임감 있는 성격입니다.',
  ESFJ: '따뜻하고 사교적이며, 타인의 필요를 챙기고 화합을 추구하는 배려심 깊은 성격입니다.',
  ISTP: '현실적이고 논리적이며, 문제를 실질적으로 분석하고 즉흥적으로 대응하는 실용주의 성격입니다.',
  ISFP: '온화하고 친절하며, 자신의 가치관을 소중히 여기고 현재의 순간을 즐기는 감성적 성격입니다.',
  ESTP: '대담하고 현실적이며, 즉각적인 행동을 선호하고 변화와 도전을 즐기는 활동적 성격입니다.',
  ESFP: '외향적이고 친근하며, 타인과 즐거움을 나누고 현재 순간을 충분히 즐기는 명랑한 성격입니다.',
};

function countsFor_(rows, filterFn) {
  const counts = {};
  TYPE_GROUPS_.forEach(g => g.types.forEach(t => counts[t] = 0));
  rows.filter(filterFn).forEach(r => { if (counts.hasOwnProperty(r.mbti)) counts[r.mbti]++; });
  return counts;
}

function addSlideHeader_(slide, title, pageWidth) {
  const barHeight = 36;
  slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, pageWidth, barHeight).getFill().setSolidFill('#2E7D32');
  slide.insertTextBox(title, 12, 3, pageWidth - 24, barHeight - 6)
    .getText().getTextStyle().setFontSize(13).setBold(true).setForegroundColor('#FFFFFF');
}

// 유형별 요약(명수+이름) 슬라이드 — 페이지 실제 폭에 맞춰 16개 박스를 한 화면에 배치
function addSummarySlide_(deck, label, scopeRows) {
  const pw = deck.getPageWidth();
  const ph = deck.getPageHeight();
  const headerHeight = 36;
  const margin = 8;
  const gap = 5;
  const cols = 4;
  const rows = 4;
  const boxW = (pw - margin * 2 - gap * (cols - 1)) / cols;
  const top = headerHeight + 6;
  const boxH = (ph - top - margin - gap * (rows - 1)) / rows;

  const slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  addSlideHeader_(slide, `${label}  ·  총 ${scopeRows.length}명`, pw);
  const counts = countsFor_(scopeRows, () => true);

  TYPE_GROUPS_.forEach((g, gi) => {
    g.types.forEach((t, ti) => {
      const x = margin + ti * (boxW + gap);
      const y = top + gi * (boxH + gap);
      const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, boxW, boxH);
      box.getFill().setSolidFill(g.color);
      box.getBorder().setTransparent();

      const names = scopeRows.filter(r => r.mbti === t).map(r => r.name);
      const namesText = names.length ? names.join(', ') : '-';
      const title = `${TYPE_EMOJI_[t]} ${t} (${counts[t]}명)`;
      const text = `${title}\n${namesText}`;

      const tb = slide.insertTextBox(text, x + 3, y + 2, boxW - 6, boxH - 4);
      const range = tb.getText();
      range.getTextStyle().setForegroundColor('#FFFFFF').setBold(true);
      range.getRange(0, title.length).getTextStyle().setFontSize(14);
      range.getRange(title.length + 1, text.length).getTextStyle().setFontSize(11).setBold(false);
      range.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });
  });
}

// 유형별 어울리는 직업 + 학생 이름 슬라이드 (반 단위로 4장씩, 한 장에 4개 유형만 크게 배치)
function addCareerSlides_(deck, label, scopeRows) {
  const pw = deck.getPageWidth();
  const ph = deck.getPageHeight();
  const headerHeight = 36;
  const margin = 8;
  const gap = 8;
  const cols = 2;
  const rows = 2;
  const boxW = (pw - margin * 2 - gap * (cols - 1)) / cols;
  const top = headerHeight + 6;
  const boxH = (ph - top - margin - gap * (rows - 1)) / rows;

  const allTypes = TYPE_GROUPS_.reduce((acc, g) => acc.concat(g.types.map(t => ({ t: t, color: g.color }))), []);
  const chunkSize = 4;
  const chunks = [];
  for (let i = 0; i < allTypes.length; i += chunkSize) chunks.push(allTypes.slice(i, i + chunkSize));

  chunks.forEach((chunk, idx) => {
    const slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    addSlideHeader_(slide, `${label} · 유형별 어울리는 직업 (${idx + 1}/${chunks.length})`, pw);

    chunk.forEach((item, i) => {
      const col = i % cols;
      const row = Math.floor(i / cols);
      const x = margin + col * (boxW + gap);
      const y = top + row * (boxH + gap);
      const box = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, boxW, boxH);
      box.getFill().setSolidFill(item.color);
      box.getBorder().setTransparent();

      const title = `${TYPE_EMOJI_[item.t]} ${item.t} · ${TYPE_NICK_[item.t]}`;
      const desc = TYPE_DESC_[item.t] || '';
      const jobs = (TYPE_CAREERS_[item.t] || []).join(' · ');
      const infoText = `${title}\n${desc}\n${jobs}`;

      const l2Start = title.length + 1;
      const l2End = l2Start + desc.length;
      const l3Start = l2End + 1;

      const nameBoxH = boxH * 0.22;
      const infoH = boxH - nameBoxH - 4;
      const tb = slide.insertTextBox(infoText, x + 10, y + 8, boxW - 20, infoH - 12);
      const range = tb.getText();
      range.getTextStyle().setForegroundColor('#FFFFFF');
      range.getRange(0, title.length).getTextStyle().setFontSize(24).setBold(true);
      range.getRange(l2Start, l2End).getTextStyle().setFontSize(15).setBold(false);
      range.getRange(l3Start, infoText.length).getTextStyle().setFontSize(14).setBold(false);

      // 학생 이름은 아래쪽에 작은 띠로 살짝 넣어서, 위 설명 영역을 가리지 않게 함
      const names = scopeRows.filter(r => r.mbti === item.t).map(r => r.name);
      const namesText = names.length ? names.join(', ') : '(제출자 없음)';
      const nameBoxY = y + infoH + 4;
      const nameBox = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 6, nameBoxY, boxW - 12, nameBoxH);
      nameBox.getFill().setSolidFill('#1B2A56');
      nameBox.getBorder().setWeight(2);
      nameBox.getBorder().getLineFill().setSolidFill('#FFEB3B');
      const nameTb = slide.insertTextBox(namesText, x + 10, nameBoxY + 2, boxW - 20, nameBoxH - 4);
      nameTb.getText().getTextStyle().setFontSize(12).setBold(true).setForegroundColor('#FFEB3B');
      nameTb.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });
  });
}

function createSlidesDeck_(rows, classFilter) {
  const deck = SlidesApp.create('MBTI 조사 결과 - ' + Utilities.formatDate(new Date(), 'GMT+9', 'yyyy.MM.dd'));
  const pw = deck.getPageWidth();
  const textW = pw - 80;
  const titleSlide = deck.getSlides()[0];
  titleSlide.getShapes().forEach(sh => { try { sh.remove(); } catch (err) {} });
  const titleText = (classFilter && classFilter !== 'all') ? `🧭 우리 ${classFilter}반 MBTI 조사 결과` : '🧭 우리 반 MBTI 조사 결과 (전체 학급)';
  titleSlide.insertTextBox(titleText, 40, 150, textW, 70)
    .getText().getTextStyle().setFontSize(36).setBold(true);
  titleSlide.insertTextBox(`총 ${rows.length}명 응답`, 40, 230, textW, 40)
    .getText().getTextStyle().setFontSize(18).setForegroundColor('#5A705A');
  titleSlide.insertTextBox('⚠️ MBTI는 성격을 이해하는 참고 자료일 뿐, 사람을 규정짓는 절대적인 기준이 아닙니다. 검사 시점과 상황에 따라 얼마든지 변할 수 있어요.', 40, 300, textW, 90)
    .getText().getTextStyle().setFontSize(14).setBold(true).setForegroundColor('#8A6D00');

  if (!classFilter || classFilter === 'all') {
    for (let c = 1; c <= 8; c++) {
      const classNum = String(c);
      const scopeRows = rows.filter(r => r.classNum === classNum);
      addSummarySlide_(deck, `${c}반`, scopeRows);
      addCareerSlides_(deck, `${c}반`, scopeRows);
    }
    addSummarySlide_(deck, '전체 학급 합계', rows);
  } else {
    const scopeRows = rows.filter(r => r.classNum === classFilter);
    addSummarySlide_(deck, `${classFilter}반`, scopeRows);
    addCareerSlides_(deck, `${classFilter}반`, scopeRows);
  }

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
    const url = createSlidesDeck_(getAllRows_(), 'all');
    const html = HtmlService
      .createHtmlOutput(`<p>슬라이드가 생성되었습니다.</p><p><a href="${url}" target="_blank">${url}</a></p>`)
      .setWidth(420)
      .setHeight(120);
    ui.showModalDialog(html, 'MBTI 결과 슬라이드');
  } catch (err) {
    ui.alert('오류가 발생했습니다: ' + err);
  }
}
