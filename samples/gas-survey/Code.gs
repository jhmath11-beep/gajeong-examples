/**
 * 우리 반 식습관 실태 조사 + 실시간 결과 — Code.gs
 *
 * 학생이 휴대폰으로 문항에 답하면 이 구글 시트 '응답' 탭에 1명 = 1줄로 쌓이고,
 * 교사 화면에서 문항별 그래프·반별 비교·서술형 모아보기·CSV를 봅니다.
 * 설치 순서는 같은 폴더의 '설치방법.md'를 보세요.
 *
 * 문항을 바꾸려면 아래 QUESTIONS 배열만 고치면 화면도 그대로 바뀝니다.
 * (이미 응답이 쌓인 뒤 문항을 바꾸면 '응답' 시트 이름을 '응답(이전)'처럼 바꿔 두세요. 새 시트가 자동으로 생깁니다.)
 */

// ===================== 설정 (여기만 고치세요) =====================
const TEACHER_PIN = '0000';          // ← 바꾸세요! 숫자 6자리 이상 권장. '0000' 그대로면 교사 화면이 열리지 않습니다.
const APP_TITLE = '우리 반 식습관 조사';
const SURVEY_INTRO = '정답은 없어요. 요즘 나의 모습 그대로 골라 주세요. 이름은 받지 않아요.';
const CLASSES = ['1반', '2반', '3반', '4반'];
const MAX_NUMBER = 40;               // 번호(선택) 최댓값. 번호를 적으면 같은 반·번호의 이전 응답을 덮어씁니다.

// type: 'single' = 하나 고르기, 'multi' = 여러 개 고르기(max 개까지), 'text' = 글로 쓰기(max 글자까지)
// 선택지 글자에는 쉼표(,)를 넣지 마세요(시트에 쉼표로 이어 적기 때문).
const QUESTIONS = [
  { id: 'breakfast', type: 'single', title: '일주일에 아침밥을 며칠 먹나요?',
    options: ['매일', '5~6일', '3~4일', '1~2일', '거의 안 먹음'] },
  { id: 'skipReason', type: 'multi', max: 2, title: '아침을 거른다면 그 이유는? (2개까지)',
    options: ['시간이 없어서', '입맛이 없어서', '살을 빼려고', '차려진 음식이 없어서', '습관이 돼서', '거르지 않음'] },
  { id: 'snackFreq', type: 'single', title: '하루에 간식을 몇 번 먹나요?',
    options: ['안 먹음', '1번', '2번', '3번 이상'] },
  { id: 'snackKind', type: 'multi', max: 3, title: '자주 먹는 간식은? (3개까지)',
    options: ['과자', '빵·떡', '과일', '우유·요구르트', '탄산·단 음료', '라면', '견과류'] },
  { id: 'water', type: 'single', title: '하루에 물을 몇 컵(200mL) 마시나요?',
    options: ['1컵 이하', '2~3컵', '4~5컵', '6컵 이상'] },
  { id: 'veg', type: 'single', title: '채소 반찬을 얼마나 자주 먹나요?',
    options: ['끼니마다', '하루 1번', '이틀에 1번', '거의 안 먹음'] },
  { id: 'lateNight', type: 'single', title: '일주일에 야식(밤 9시 이후)을 몇 번 먹나요?',
    options: ['안 먹음', '1~2번', '3~4번', '5번 이상'] },
  { id: 'fastFood', type: 'single', title: '일주일에 패스트푸드(햄버거·피자·치킨)를 몇 번 먹나요?',
    options: ['안 먹음', '1번', '2~3번', '4번 이상'] },
  { id: 'pledge', type: 'text', max: 100, title: '나의 식습관 개선 다짐 한 가지',
    placeholder: '예: 아침에 우유 한 잔이라도 마시기' }
];
// =================================================================

const SHEET_RESPONSES = '응답';
const FIXED_HEADERS = ['제출시각', '반', '번호'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 시트를 미리 만들어 둡니다. (안 해도 첫 사용 때 자동 생성) */
function setup() {
  const headers = headers_();
  const sheet = getSheet_(SHEET_RESPONSES, headers);
  const now = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (now.join('\u0001') !== headers.join('\u0001')) {
    if (sheet.getLastRow() > 1) {
      fail_("문항이 바뀌었어요. 구글 시트에서 '응답' 탭 이름을 '응답(이전)'으로 바꾸면 새 '응답' 탭이 만들어져요.");
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

// ---------------- 학생 화면 ----------------

function getConfig() {
  return { appTitle: APP_TITLE, intro: SURVEY_INTRO, classes: CLASSES, maxNumber: MAX_NUMBER, questions: QUESTIONS };
}

/** 응답 제출. 번호를 적었으면 같은 반·번호의 이전 응답을 덮어씁니다. */
function submitResponse(resp) {
  const v = validateResponse_(resp);
  return withLock_(function () {
    const sheet = setup();
    let replaced = false;
    if (v.num !== '') {
      const values = sheet.getDataRange().getValues();
      for (let r = values.length - 1; r >= 1; r--) {
        if (String(values[r][1]) === v.cls && String(values[r][2]) === String(v.num)) { sheet.deleteRow(r + 1); replaced = true; }
      }
    }
    const now = new Date();
    const row = [now, v.cls, v.num].concat(QUESTIONS.map(function (q) {
      const a = v.answers[q.id];
      return Array.isArray(a) ? a.join(', ') : a;
    }));
    const at = sheet.getLastRow() + 1;
    sheet.getRange(at, 2, 1, row.length - 1).setNumberFormat('@');   // 반·번호·답을 글자 그대로(시트가 '3-4' 같은 답을 날짜로 바꾸지 않게)
    sheet.getRange(at, 1, 1, row.length).setValues([row]);
    return { ok: true, replaced: replaced, savedAt: now.toISOString() };
  });
}

// ---------------- 교사 화면 (PIN 확인) ----------------

/** PIN이 맞을 때만 모든 응답을 돌려줍니다. 그래프 계산은 화면에서 합니다. */
function getTeacherData(pin) {
  checkPin_(pin);
  const values = setup().getDataRange().getValues().slice(1);
  const responses = values.filter(function (r) { return String(r[1]) !== ''; }).map(function (r) {
    const answers = {};
    QUESTIONS.forEach(function (q, i) {
      const cell = String(r[FIXED_HEADERS.length + i] == null ? '' : r[FIXED_HEADERS.length + i]);
      answers[q.id] = q.type === 'multi' ? (cell ? cell.split(', ') : []) : cell;
    });
    return { savedAt: toIso_(r[0]), cls: String(r[1]), num: r[2] === '' ? '' : Number(r[2]), answers: answers };
  });
  return { config: getConfig(), responses: responses };
}

// ---------------- 내부 도우미 ----------------

function headers_() {
  return FIXED_HEADERS.concat(QUESTIONS.map(function (q) { return q.title; }));
}

/** 서버 쪽 입력 검사 — 선택지에 없는 값, 너무 긴 글, 범위를 벗어난 번호를 막습니다. */
function validateResponse_(resp) {
  if (!resp || typeof resp !== 'object') fail_('응답이 비어 있어요.');
  const cls = String(resp.cls || '');
  if (CLASSES.indexOf(cls) < 0) fail_('반을 골라 주세요.');
  let num = '';
  if (resp.num !== '' && resp.num != null) {
    num = Number(resp.num);
    if (!isInt_(num, 1, MAX_NUMBER)) fail_('번호는 1~' + MAX_NUMBER + ' 사이로 적거나 비워 두세요.');
  }
  const input = resp.answers && typeof resp.answers === 'object' ? resp.answers : {};
  const answers = {};
  QUESTIONS.forEach(function (q, i) {
    const label = (i + 1) + '번 문항';
    const a = input[q.id];
    if (q.type === 'single') {
      if (q.options.indexOf(a) < 0) fail_(label + '에 답해 주세요.');
      answers[q.id] = a;
    } else if (q.type === 'multi') {
      if (!Array.isArray(a) || a.length === 0) fail_(label + '에서 1개 이상 골라 주세요.');
      if (a.length > (q.max || q.options.length)) fail_(label + '은 ' + (q.max || q.options.length) + '개까지 고를 수 있어요.');
      a.forEach(function (x) { if (q.options.indexOf(x) < 0) fail_(label + '에 없는 선택지예요.'); });
      if (new Set(a).size !== a.length) fail_(label + '에 같은 선택지가 두 번 들어갔어요.');
      answers[q.id] = q.options.filter(function (o) { return a.indexOf(o) >= 0; });   // 선택지 순서대로 정리
    } else {
      const s = clean_(a, q.max || 100, label);
      if (!s) fail_(label + '에 한 줄이라도 적어 주세요.');
      answers[q.id] = s;
    }
  });
  return { cls: cls, num: num, answers: answers };
}

// ---- 세 샘플 공통 도우미 ----

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  const isNew = !sheet;
  if (isNew) sheet = ss.insertSheet(name);
  // 문항이 많아 기본 열 수(26)를 넘으면 열을 늘림
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  if (isNew) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) fail_('지금 제출하는 사람이 많아요. 잠시 뒤 다시 눌러 주세요.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function checkPin_(pin) {
  if (String(TEACHER_PIN) === '0000') fail_('먼저 Code.gs 맨 위의 TEACHER_PIN을 바꾸고, 저장한 뒤 새 버전으로 배포하세요.');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= 20) fail_('PIN을 너무 많이 틀렸어요. 10분 뒤 다시 시도하세요.');
  if (String(pin == null ? '' : pin).trim() !== String(TEACHER_PIN)) {
    cache.put('pinFails', String(fails + 1), 600);
    fail_('PIN이 맞지 않아요.');
  }
}

function clean_(value, max, label) {
  let s = String(value == null ? '' : value).replace(/[\u0000-\u001f\u007f]/g, ' ').trim();
  if (s.length > max) fail_(label + ' 글자 수는 ' + max + '자 이내로 적어 주세요.');
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 시트 수식 주입 방지
  return s;
}

function isInt_(n, min, max) { return Number.isInteger(n) && n >= min && n <= max; }
function toIso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
function fail_(msg) { throw new Error(msg); }
