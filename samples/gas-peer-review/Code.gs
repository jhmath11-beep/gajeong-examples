/**
 * 모둠 활동 자기·동료평가 수합 — Code.gs
 *
 * 학생이 반·모둠·(모둠 안) 번호를 고르고, 자기평가 + 같은 모둠원 각자 평가 + 한 줄 칭찬/조언을 냅니다.
 * 이 구글 시트 '평가' 탭에 (평가한 사람 → 평가받은 사람) 1쌍 = 1줄로 쌓이고,
 * 교사 화면에서 자기 점수 vs 동료 평균, 항목별 평균, 차이 큰 학생, 미제출자, 피드백, CSV를 봅니다.
 * 설치 순서는 같은 폴더의 '설치방법.md'를 보세요.
 *
 * 새 평가를 할 때는 EVAL_NAME만 바꾸면 됩니다. 이전 평가 기록은 시트에 남고, 화면에는 지금 평가만 보입니다.
 */

// ===================== 설정 (여기만 고치세요) =====================
const TEACHER_PIN = '0000';          // ← 바꾸세요! 숫자 6자리 이상 권장. '0000' 그대로면 교사 화면이 열리지 않습니다.
const APP_TITLE = '모둠 활동 자기·동료평가';
const EVAL_NAME = '조리실습 3차';     // 평가 이름 (바꾸면 새 평가로 따로 모입니다)
const CLASSES = ['1반', '2반', '3반', '4반'];
const GROUPS_PER_CLASS = 6;
const MAX_GROUP_SIZE = 6;
const DIFF_ALERT = 1.0;               // 자기 평균 - 동료 평균 차이가 이 값 이상이면 교사 화면에 표시
const SCORE_LABELS = ['전혀 아니다', '조금 아니다', '보통이다', '대체로 그렇다', '매우 그렇다'];   // 1점~5점 뜻

// 루브릭 항목 (4~5개 권장). id는 영문, name·desc는 자유롭게 고치세요.
// (평가를 받기 시작한 뒤 항목을 바꾸면 '평가' 시트 이름을 '평가(이전)'으로 바꿔 두세요.)
const CRITERIA = [
  { id: 'role',   name: '역할 수행', desc: '맡은 역할(조리·계량·정리 등)을 끝까지 책임 있게 했다.' },
  { id: 'coop',   name: '협력',     desc: '모둠원과 의견을 나누고 어려운 친구를 도왔다.' },
  { id: 'safety', name: '위생·안전', desc: '손 씻기, 칼·불 사용 규칙을 지켰다.' },
  { id: 'tidy',   name: '정리정돈', desc: '뒷정리와 설거지에 끝까지 참여했다.' },
  { id: 'part',   name: '참여 태도', desc: '활동에 적극적으로 참여하고 시간을 지켰다.' }
];
// =================================================================

const SHEET_REVIEWS = '평가';
const FIXED_HEADERS = ['제출시각', '평가명', '반', '모둠', '평가한 번호', '평가받은 번호', '구분', '모둠 인원'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 시트를 미리 만들어 둡니다. (안 해도 첫 사용 때 자동 생성) */
function setup() {
  const headers = headers_();
  const sheet = getSheet_(SHEET_REVIEWS, headers);
  const now = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (now.join('\u0001') !== headers.join('\u0001')) {
    if (sheet.getLastRow() > 1) {
      fail_("평가 항목이 바뀌었어요. 구글 시트에서 '평가' 탭 이름을 '평가(이전)'으로 바꾸면 새 '평가' 탭이 만들어져요.");
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

// ---------------- 학생 화면 ----------------

function getConfig() {
  return { appTitle: APP_TITLE, evalName: EVAL_NAME, classes: CLASSES, groups: GROUPS_PER_CLASS,
           maxGroupSize: MAX_GROUP_SIZE, diffAlert: DIFF_ALERT, scoreLabels: SCORE_LABELS, criteria: CRITERIA };
}

/** 평가 제출. 같은 평가·반·모둠·번호가 다시 내면 이전 것을 지우고 새로 씁니다. */
function submitReview(review) {
  const v = validateReview_(review);
  return withLock_(function () {
    const sheet = setup();
    const values = sheet.getDataRange().getValues();
    let replaced = false;
    for (let r = values.length - 1; r >= 1; r--) {
      const row = values[r];
      if (String(row[1]) === EVAL_NAME && String(row[2]) === v.cls && Number(row[3]) === v.group && Number(row[4]) === v.num) {
        sheet.deleteRow(r + 1); replaced = true;
      }
    }
    const now = new Date();
    const rows = v.targets.map(function (t) {
      return [now, EVAL_NAME, v.cls, v.group, v.num, t.num, t.num === v.num ? '자기' : '동료', v.size]
        .concat(CRITERIA.map(function (c) { return t.scores[c.id]; }))
        .concat([t.comment]);
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { ok: true, replaced: replaced, savedAt: now.toISOString() };
  });
}

// ---------------- 교사 화면 (PIN 확인) ----------------

/** PIN이 맞을 때만 지금 평가(EVAL_NAME)의 모든 줄을 돌려줍니다. 평균·차이 계산은 화면에서 합니다. */
function getTeacherData(pin) {
  checkPin_(pin);
  const n = FIXED_HEADERS.length;
  const rows = setup().getDataRange().getValues().slice(1)
    .filter(function (r) { return String(r[1]) === EVAL_NAME; })
    .map(function (r) {
      const scores = {};
      CRITERIA.forEach(function (c, i) { scores[c.id] = Number(r[n + i]); });
      return { savedAt: toIso_(r[0]), cls: String(r[2]), group: Number(r[3]), from: Number(r[4]), to: Number(r[5]),
               self: String(r[6]) === '자기', size: Number(r[7]), scores: scores, comment: String(r[n + CRITERIA.length] || '') };
    });
  return { config: getConfig(), rows: rows };
}

// ---------------- 내부 도우미 ----------------

function headers_() {
  return FIXED_HEADERS.concat(CRITERIA.map(function (c) { return c.name; })).concat(['한마디']);
}

/** 서버 쪽 입력 검사 — 모둠원 전원을 빠짐없이, 점수는 1~5 정수만. */
function validateReview_(review) {
  if (!review || typeof review !== 'object') fail_('평가 내용이 비어 있어요.');
  const cls = String(review.cls || '');
  if (CLASSES.indexOf(cls) < 0) fail_('반을 골라 주세요.');
  const group = Number(review.group);
  if (!isInt_(group, 1, GROUPS_PER_CLASS)) fail_('모둠 번호는 1~' + GROUPS_PER_CLASS + ' 사이로 골라 주세요.');
  const size = Number(review.size);
  if (!isInt_(size, 2, MAX_GROUP_SIZE)) fail_('모둠 인원은 2~' + MAX_GROUP_SIZE + '명 사이로 골라 주세요.');
  const num = Number(review.num);
  if (!isInt_(num, 1, size)) fail_('내 번호는 1~' + size + ' 사이로 골라 주세요.');
  const list = Array.isArray(review.targets) ? review.targets : [];
  if (list.length !== size) fail_('나를 포함해 모둠원 ' + size + '명을 모두 평가해 주세요.');
  const seen = {};
  const targets = list.map(function (t) {
    const to = Number(t && t.num);
    if (!isInt_(to, 1, size) || seen[to]) fail_('평가 대상 번호가 잘못됐어요. 처음부터 다시 해 주세요.');
    seen[to] = true;
    const who = to === num ? '자기평가' : to + '번 모둠원';
    const scores = {};
    CRITERIA.forEach(function (c) {
      const s = Number(t.scores && t.scores[c.id]);
      if (!isInt_(s, 1, 5)) fail_(who + '의 ‘' + c.name + '’ 점수를 1~5점 중에서 골라 주세요.');
      scores[c.id] = s;
    });
    const comment = clean_(t.comment, 100, who + ' 한마디');
    if (to !== num && comment.length < 2) fail_(to + '번 모둠원에게 칭찬이나 조언을 한 줄 적어 주세요.');
    return { num: to, scores: scores, comment: comment };
  });
  targets.sort(function (a, b) { return a.num - b.num; });
  return { cls: cls, group: group, size: size, num: num, targets: targets };
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
