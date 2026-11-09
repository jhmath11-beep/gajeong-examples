/**
 * 실습 재료 신청 & 장보기 목록 자동 집계 — Code.gs
 *
 * 모둠이 휴대폰으로 실습 메뉴·재료를 신청하면 이 구글 시트에 모이고,
 * 교사 화면에서 반별 제출 현황·장보기 목록(품목별 합계)·모둠별 신청서를 봅니다.
 * 설치 순서는 같은 폴더의 '설치방법.md'를 보세요.
 *
 * 시트 3개를 자동으로 만듭니다.
 *   - 품목 : 신청할 수 있는 재료 표(품목·단위·1인분량·분류). 시트에서 직접 고치면 바로 반영됩니다.
 *   - 신청 : 모둠이 낸 신청(품목 1개 = 1줄). 같은 반·모둠이 다시 내면 덮어씁니다.
 */

// ===================== 설정 (여기만 고치세요) =====================
const TEACHER_PIN = '0000';          // ← 바꾸세요! 숫자 6자리 이상 권장. '0000' 그대로면 교사 화면이 열리지 않습니다.
const APP_TITLE = '실습 재료 신청';
const PRACTICE_TITLE = '2학기 조리실습';           // 화면 맨 위에 보이는 실습 이름
const CLASSES = ['1반', '2반', '3반', '4반'];      // 반 목록
const GROUPS_PER_CLASS = 6;                        // 반마다 모둠 수
const PRACTICE_DATES = {                           // 반별 실습 날짜 (모르면 '' 로 두기)
  '1반': '2026-11-20', '2반': '2026-11-20', '3반': '2026-11-23', '4반': '2026-11-23'
};
const MAX_PEOPLE = 8;                              // 모둠 최대 인원

// '품목' 시트가 없을 때 처음 한 번 넣는 예시 재료입니다(1인분량은 예시값 — 메뉴에 맞게 시트에서 고치세요).
const DEFAULT_ITEMS = [
  // 품목, 단위, 1인분량, 분류
  ['쌀', 'g', 100, '곡류'],
  ['식빵', '장', 2, '곡류'],
  ['감자', '개', 0.5, '채소'],
  ['양파', '개', 0.25, '채소'],
  ['당근', '개', 0.2, '채소'],
  ['애호박', '개', 0.2, '채소'],
  ['대파', '대', 0.2, '채소'],
  ['배추김치', 'g', 50, '채소'],
  ['달걀', '개', 1, '고기·달걀·콩'],
  ['돼지고기(앞다리)', 'g', 60, '고기·달걀·콩'],
  ['두부', '모', 0.25, '고기·달걀·콩'],
  ['우유', 'mL', 200, '우유·유제품'],
  ['슬라이스 치즈', '장', 1, '우유·유제품'],
  ['식용유', 'mL', 10, '양념'],
  ['간장', 'mL', 10, '양념']
];
// =================================================================

const SHEET_ITEMS = '품목';
const SHEET_ORDERS = '신청';
const ITEM_HEADERS = ['품목', '단위', '1인분량', '분류'];
const ORDER_HEADERS = ['제출시각', '반', '모둠', '메뉴', '인원', '품목', '수량', '단위', '모둠이 가져옴', '그 밖의 준비물'];

/** 웹 앱 주소로 들어오면 화면(index.html)을 보여 줍니다. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 시트를 미리 만들어 둡니다. (편집기 위쪽에서 setup 고르고 ▶실행 — 안 해도 첫 사용 때 자동 생성) */
function setup() {
  getSheet_(SHEET_ORDERS, ORDER_HEADERS);
  const items = getSheet_(SHEET_ITEMS, ITEM_HEADERS);
  if (items.getLastRow() < 2) {
    items.getRange(2, 1, DEFAULT_ITEMS.length, ITEM_HEADERS.length).setValues(DEFAULT_ITEMS);
  }
}

// ---------------- 학생(모둠) 화면에서 부르는 함수 ----------------

/** 화면에 필요한 설정과 품목 표 */
function getConfig() {
  setup();
  return {
    appTitle: APP_TITLE,
    practiceTitle: PRACTICE_TITLE,
    classes: CLASSES,
    groups: GROUPS_PER_CLASS,
    dates: PRACTICE_DATES,
    maxPeople: MAX_PEOPLE,
    items: readItems_()
  };
}

/** 우리 모둠이 전에 낸 신청 불러오기(없으면 null) */
function getOrder(cls, group) {
  const key = String(cls) + '|' + Number(group);
  const orders = readOrders_();
  return orders.filter(function (o) { return o.cls + '|' + o.group === key; })[0] || null;
}

/** 신청 제출. 같은 반·모둠이 이미 냈으면 덮어씁니다. */
function submitOrder(order) {
  setup();
  const o = validateOrder_(order, readItems_());
  return withLock_(function () {
    const sheet = getSheet_(SHEET_ORDERS, ORDER_HEADERS);
    const values = sheet.getDataRange().getValues();
    let replaced = false;
    for (let r = values.length - 1; r >= 1; r--) {          // 아래에서 위로 지워야 줄 번호가 안 밀림
      if (String(values[r][1]) === o.cls && Number(values[r][2]) === o.group) {
        sheet.deleteRow(r + 1);
        replaced = true;
      }
    }
    const now = new Date();
    const rows = o.items.map(function (it) {
      return [now, o.cls, o.group, o.menu, o.people, it.name, it.qty, it.unit, it.bring ? 'O' : '', o.note];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ORDER_HEADERS.length).setValues(rows);
    return { ok: true, replaced: replaced, savedAt: now.toISOString() };
  });
}

// ---------------- 교사 화면에서 부르는 함수 (PIN 확인) ----------------

/** PIN이 맞을 때만 모든 신청을 돌려줍니다. 집계(장보기 목록)는 화면에서 계산합니다. */
function getTeacherData(pin) {
  checkPin_(pin);
  setup();
  return { config: getConfig(), orders: readOrders_() };
}

// ---------------- 내부 도우미 ----------------

function readItems_() {
  const values = getSheet_(SHEET_ITEMS, ITEM_HEADERS).getDataRange().getValues().slice(1);
  return values
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      return { name: String(r[0]).trim(), unit: String(r[1]).trim(), per: Number(r[2]) || 0, cat: String(r[3]).trim() || '기타' };
    });
}

function readOrders_() {
  const values = getSheet_(SHEET_ORDERS, ORDER_HEADERS).getDataRange().getValues().slice(1);
  const map = {};
  const list = [];
  values.forEach(function (r) {
    if (String(r[1]) === '') return;
    const key = String(r[1]) + '|' + Number(r[2]);
    if (!map[key]) {
      map[key] = { cls: String(r[1]), group: Number(r[2]), menu: String(r[3]), people: Number(r[4]),
                   note: String(r[9] || ''), savedAt: toIso_(r[0]), items: [] };
      list.push(map[key]);
    }
    map[key].items.push({ name: String(r[5]), qty: Number(r[6]), unit: String(r[7]), bring: String(r[8]) === 'O' });
  });
  return list;
}

/** 서버 쪽 입력 검사 — 화면을 우회한 요청도 여기서 막습니다. */
function validateOrder_(order, items) {
  if (!order || typeof order !== 'object') fail_('신청 내용이 비어 있어요.');
  const cls = String(order.cls || '');
  if (CLASSES.indexOf(cls) < 0) fail_('반을 골라 주세요.');
  const group = Number(order.group);
  if (!isInt_(group, 1, GROUPS_PER_CLASS)) fail_('모둠 번호는 1~' + GROUPS_PER_CLASS + ' 사이로 골라 주세요.');
  const menu = clean_(order.menu, 40, '실습 메뉴');
  if (!menu) fail_('실습 메뉴를 적어 주세요.');
  const people = Number(order.people);
  if (!isInt_(people, 1, MAX_PEOPLE)) fail_('인원은 1~' + MAX_PEOPLE + '명 사이로 적어 주세요.');
  const note = clean_(order.note, 100, '그 밖의 준비물');
  if (!Array.isArray(order.items) || order.items.length === 0) fail_('필요한 재료를 1개 이상 골라 주세요.');
  if (order.items.length > 50) fail_('재료는 50개까지만 신청할 수 있어요.');
  const byName = {};
  items.forEach(function (it) { byName[it.name] = it; });
  const seen = {};
  const list = order.items.map(function (it) {
    const name = String(it && it.name || '');
    const item = byName[name];
    if (!item) fail_('품목 표에 없는 재료예요: ' + name.slice(0, 20));
    if (seen[name]) fail_('같은 재료가 두 번 들어갔어요: ' + name);
    seen[name] = true;
    const qty = Math.round(Number(it.qty) * 100) / 100;
    if (!isFinite(qty) || qty <= 0 || qty > 9999) fail_(name + ' 수량은 0보다 크고 9999 이하로 적어 주세요.');
    return { name: name, unit: item.unit, qty: qty, bring: it.bring === true };
  });
  return { cls: cls, group: group, menu: menu, people: people, note: note, items: list };
}

// ---- 세 샘플 공통 도우미 ----

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 동시에 여러 모둠이 제출해도 줄이 섞이지 않게 한 번에 한 명씩 씁니다. */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) fail_('지금 제출하는 사람이 많아요. 잠시 뒤 다시 눌러 주세요.');
  try { return fn(); } finally { lock.releaseLock(); }
}

/** 교사 PIN 확인. 틀리면 데이터를 주지 않고 오류를 냅니다. 여러 번 틀리면 10분 잠금. */
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

/** 글자 정리: 줄바꿈·제어문자 제거, 길이 제한, 시트 수식으로 실행되지 않게 막기 */
function clean_(value, max, label) {
  let s = String(value == null ? '' : value).replace(/[\u0000-\u001f\u007f]/g, ' ').trim();
  if (s.length > max) fail_(label + ' 글자 수는 ' + max + '자 이내로 적어 주세요.');
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // =IMPORTXML(...) 같은 수식 주입 방지
  return s;
}

function isInt_(n, min, max) { return Number.isInteger(n) && n >= min && n <= max; }
function toIso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
function fail_(msg) { throw new Error(msg); }
