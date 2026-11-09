/**
 * 업사이클링 모둠 프로젝트 진행 보드 — Code.gs
 *
 * 모둠(반·모둠)이 모둠 비밀번호 4자리로 들어와 단계별 체크(조사 → 디자인 구상 → 재료 준비 → 제작 → 발표),
 * 진행 메모·다음 할 일·막힌 점을 저장합니다. 같은 모둠은 한 줄을 덮어씁니다.
 * 교사 화면(PIN)에서 반 전체 진행 현황판(단계별 열), 막힌 점 모아보기, 모둠별 피드백 한 줄, 오래 멈춘 모둠, CSV를 봅니다.
 * 설치 순서는 같은 폴더의 '설치방법.md'를 보세요.
 *
 * 새 프로젝트를 할 때는 PROJECT_NAME만 바꾸면 됩니다. 이전 기록은 시트에 남고, 화면에는 지금 프로젝트만 보입니다.
 */

// ===================== 설정 (여기만 고치세요) =====================
const TEACHER_PIN = '0000';          // ← 바꾸세요! 숫자 6자리 이상 권장. '0000' 그대로면 교사 화면이 열리지 않습니다.
const APP_TITLE = '업사이클링 모둠 프로젝트';
const PROJECT_NAME = '헌 옷 업사이클링 2학기';   // 바꾸면 새 프로젝트로 따로 모입니다
const CLASSES = ['1반', '2반', '3반', '4반'];
const GROUPS_PER_CLASS = 6;
const STALE_DAYS = 3;                // 이 날수 넘게 저장이 없으면 교사 화면에 '멈춤' 표시

// 단계와 단계별 체크 항목. id는 영문(바꾸지 않기), name·checks는 수업에 맞게 고치세요.
// (진행 중에 단계 수를 바꾸면 '모둠' 시트 이름을 '모둠(이전)'으로 바꿔 두세요.)
const STAGES = [
  { id: 'research', name: '조사', checks: ['업사이클링 사례 3가지 찾기', '우리 모둠 헌 옷 소재·상태 살펴보기', '만들 물건 후보 2개 정하기'] },
  { id: 'design',   name: '디자인 구상', checks: ['완성 모습 스케치', '치수 재고 옷본(본) 그리기', '역할 나누기'] },
  { id: 'prep',     name: '재료 준비', checks: ['헌 옷 세탁·다림질', '실·바늘·단추 등 부자재 챙기기', '재단 전 선생님 확인'] },
  { id: 'make',     name: '제작', checks: ['재단하기', '손바느질·재봉하기', '마무리(실밥 정리·장식)'] },
  { id: 'present',  name: '발표', checks: ['사진 찍기', '발표 자료 만들기', '환경에 준 도움 한 줄 정리'] }
];
// =================================================================

const SHEET_GROUPS = '모둠';
const FIXED_HEADERS = ['프로젝트', '반', '모둠', '비밀번호 해시', '마지막 저장'];
const TAIL_HEADERS = ['진행 메모', '다음 할 일', '막힌 점', '교사 피드백', '피드백 시각'];
const ON = '■', OFF = '□';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 시트를 미리 만들어 둡니다. (안 해도 첫 사용 때 자동 생성) */
function setup() {
  const headers = headers_();
  const sheet = getSheet_(SHEET_GROUPS, headers);
  const now = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (now.join('\u0001') !== headers.join('\u0001')) {
    if (sheet.getLastRow() > 1) fail_("단계 설정이 바뀌었어요. 구글 시트에서 '모둠' 탭 이름을 '모둠(이전)'으로 바꾸면 새 '모둠' 탭이 만들어져요.");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

// ---------------- 학생(모둠) 화면 ----------------

function getConfig() {
  return { appTitle: APP_TITLE, projectName: PROJECT_NAME, classes: CLASSES, groups: GROUPS_PER_CLASS, staleDays: STALE_DAYS, stages: STAGES };
}

/** 모둠 열기. 처음이면 빈 보드(isNew)를 돌려주고, 첫 저장 때 이 비밀번호가 모둠 비밀번호가 됩니다. */
function openGroup(p) {
  const who = validateWho_(p);
  const g = findGroup_(who);
  if (!g) return { isNew: true, group: emptyGroup_(who) };
  checkGroupPw_(g, who);
  return { isNew: false, group: publicGroup_(g) };
}

/** 모둠 진행 저장(덮어쓰기). base = 불러올 때의 '마지막 저장' — 그 사이 다른 모둠원이 저장했으면 거절해 덮어쓰기 사고를 막습니다. */
function saveGroup(p) {
  const who = validateWho_(p);
  const v = validateProgress_(p);
  const base = String(p.base == null ? '' : p.base);
  return withLock_(function () {
    const sheet = setup();
    const g = findGroup_(who);
    const now = new Date();
    const stageCols = STAGES.map(function (s) { return v.checks[s.id].map(function (b) { return b ? ON : OFF; }).join(''); });
    if (g) {
      checkGroupPw_(g, who);
      if (base !== g.updatedAt) fail_('다른 모둠원이 방금 저장했어요. ‘다시 불러오기’를 누른 뒤 고쳐서 저장해 주세요.');
      const hash = g.hash || hash_(who);   // 교사가 비밀번호를 초기화했으면 지금 번호로 새로 정함
      sheet.getRange(g.row, 4, 1, 2 + STAGES.length + 3).setValues([[hash, now].concat(stageCols).concat([v.memo, v.next, v.blocker])]);
    } else {
      if (base) fail_('모둠 기록을 찾지 못했어요. 새로고침해 주세요.');
      sheet.appendRow([PROJECT_NAME, who.cls, who.group, hash_(who), now].concat(stageCols).concat([v.memo, v.next, v.blocker, '', '']));
    }
    // 시트에서 다시 읽은 값을 돌려줘야 다음 저장 때 base 비교가 정확합니다(시트가 밀리초를 버릴 수 있음).
    SpreadsheetApp.flush();
    return { ok: true, updatedAt: findGroup_(who).updatedAt };
  });
}

// ---------------- 교사 화면 (PIN 확인) ----------------

/** PIN이 맞을 때만 지금 프로젝트의 모든 모둠을 돌려줍니다(비밀번호 해시 제외). 단계 계산은 화면에서 합니다. */
function getTeacherData(pin) {
  checkPin_(pin);
  return { config: getConfig(), groups: readGroups_().map(publicGroup_), now: new Date().toISOString() };
}

/** 모둠 피드백 한 줄(학생 화면에 보임). 빈칸으로 저장하면 지워집니다. */
function setFeedback(pin, cls, group, text) {
  checkPin_(pin);
  const who = validateClsGroup_(cls, group);
  const fb = clean_(text, 80, '피드백');
  return withLock_(function () {
    const g = findGroup_(who);
    if (!g) fail_('아직 저장한 적 없는 모둠이에요. 모둠이 한 번 저장한 뒤에 피드백을 남길 수 있어요.');
    const now = new Date();
    setup().getRange(g.row, FIXED_HEADERS.length + STAGES.length + 4, 1, 2).setValues([[fb, fb ? now : '']]);
    return { ok: true, at: fb ? now.toISOString() : '' };
  });
}

/** 모둠 비밀번호 초기화 — 그 모둠이 다음에 저장할 때 쓰는 번호가 새 비밀번호가 됩니다. */
function resetGroupPw(pin, cls, group) {
  checkPin_(pin);
  const who = validateClsGroup_(cls, group);
  return withLock_(function () {
    const g = findGroup_(who);
    if (!g) fail_('아직 저장한 적 없는 모둠이에요.');
    setup().getRange(g.row, 4).setValue('');
    CacheService.getScriptCache().remove(failKey_(who));
    return { ok: true };
  });
}

// ---------------- 내부 도우미 ----------------

function headers_() {
  return FIXED_HEADERS.concat(STAGES.map(function (s) { return s.name; })).concat(TAIL_HEADERS);
}

function readGroups_() {
  const n = FIXED_HEADERS.length, S = STAGES.length;
  return setup().getDataRange().getValues().slice(1)
    .map(function (r, i) {
      const checks = {};
      STAGES.forEach(function (s, k) {
        const marks = Array.from(String(r[n + k] || ''));
        checks[s.id] = s.checks.map(function (_, j) { return marks[j] === ON; });
      });
      return { row: i + 2, project: String(r[0]), cls: String(r[1]), group: Number(r[2]), hash: String(r[3] || ''), updatedAt: toIso_(r[4]),
               checks: checks, memo: plain_(r[n + S]), next: plain_(r[n + S + 1]), blocker: plain_(r[n + S + 2]),
               feedback: plain_(r[n + S + 3]), feedbackAt: toIso_(r[n + S + 4]) };
    })
    .filter(function (g) { return g.project === PROJECT_NAME; });
}

function findGroup_(who) {
  return readGroups_().filter(function (g) { return g.cls === who.cls && g.group === who.group; })[0] || null;
}

function publicGroup_(g) {
  return { cls: g.cls, group: g.group, updatedAt: g.updatedAt, checks: g.checks, memo: g.memo, next: g.next, blocker: g.blocker,
           feedback: g.feedback, feedbackAt: g.feedbackAt, pwReset: !g.hash };
}

function emptyGroup_(who) {
  const checks = {};
  STAGES.forEach(function (s) { checks[s.id] = s.checks.map(function () { return false; }); });
  return { cls: who.cls, group: who.group, updatedAt: '', checks: checks, memo: '', next: '', blocker: '', feedback: '', feedbackAt: '', pwReset: false };
}

/** 모둠 비밀번호 확인. 10번 틀리면 그 모둠은 10분 잠금. 교사가 초기화한 모둠(해시 빈칸)은 통과. */
function checkGroupPw_(g, who) {
  if (!g.hash) return;
  const cache = CacheService.getScriptCache();
  const key = failKey_(who);
  const fails = Number(cache.get(key) || 0);
  if (fails >= 10) fail_('비밀번호를 너무 많이 틀렸어요. 10분 뒤 다시 해 보거나 선생님께 초기화를 부탁하세요.');
  if (g.hash !== hash_(who)) {
    cache.put(key, String(fails + 1), 600);
    fail_('모둠 비밀번호가 맞지 않아요. (처음 저장할 때 정한 4자리)');
  }
}

function failKey_(who) { return 'gFail:' + who.cls + '|' + who.group; }

/** 비밀번호 해시: 이 스프레드시트 ID를 소금(salt)으로 섞은 SHA-256. 시트에는 이 값만 저장됩니다. */
function hash_(who) {
  const raw = SpreadsheetApp.getActive().getId() + '|' + PROJECT_NAME + '|' + who.cls + '|' + who.group + '|' + who.pw;
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function validateClsGroup_(cls, group) {
  cls = String(cls || '');
  if (CLASSES.indexOf(cls) < 0) fail_('반을 골라 주세요.');
  group = Number(group);
  if (!isInt_(group, 1, GROUPS_PER_CLASS)) fail_('모둠은 1~' + GROUPS_PER_CLASS + ' 사이로 골라 주세요.');
  return { cls: cls, group: group };
}

function validateWho_(p) {
  if (!p || typeof p !== 'object') fail_('입력이 비어 있어요.');
  const who = validateClsGroup_(p.cls, p.group);
  const pw = String(p.pw == null ? '' : p.pw).trim();
  if (!/^\d{4}$/.test(pw)) fail_('모둠 비밀번호는 숫자 4자리로 적어 주세요.');
  who.pw = pw;
  return who;
}

/** 서버 쪽 입력 검사 — 단계별 체크 개수가 설정과 정확히 같아야 하고, 글자 수 제한. */
function validateProgress_(p) {
  const src = p.checks && typeof p.checks === 'object' ? p.checks : null;
  if (!src) fail_('체크 내용이 비어 있어요. 새로고침해 주세요.');
  const checks = {};
  STAGES.forEach(function (s) {
    const a = src[s.id];
    if (!Array.isArray(a) || a.length !== s.checks.length || a.some(function (b) { return typeof b !== 'boolean'; })) {
      fail_('‘' + s.name + '’ 단계 체크가 설정과 달라요. 새로고침해 주세요.');
    }
    checks[s.id] = a.slice();
  });
  return { checks: checks, memo: clean_(p.memo, 200, '진행 메모'), next: clean_(p.next, 80, '다음 할 일'), blocker: clean_(p.blocker, 100, '막힌 점') };
}

/** 시트에 넣을 때 붙인 수식 방지용 ' 를 화면에 보낼 때 뗍니다. */
function plain_(v) { return String(v == null ? '' : v).replace(/^'(?=[=+\-@])/, ''); }

// ---- 샘플 공통 도우미 ----

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
  if (!lock.tryLock(15000)) fail_('지금 저장하는 모둠이 많아요. 잠시 뒤 다시 눌러 주세요.');
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

/** 글자 정리: 제어문자 제거(메모는 줄바꿈 허용), 길이 제한, 시트 수식 주입 방지. */
function clean_(value, max, label) {
  let s = String(value == null ? '' : value).replace(/\r\n?/g, '\n').replace(/[\u0000-\u0009\u000b-\u001f\u007f]/g, ' ').trim();
  if (s.length > max) fail_(label + ' 글자 수는 ' + max + '자 이내로 적어 주세요.');
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 시트 수식 주입 방지
  return s;
}

function isInt_(n, min, max) { return Number.isInteger(n) && n >= min && n <= max; }
function toIso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
function fail_(msg) { throw new Error(msg); }
