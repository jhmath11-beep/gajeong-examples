/**
 * 학급 나눔 장터 (물물교환 게시판) — Code.gs
 *
 * 학생이 반 + 별칭 + 4자리 비밀번호로 안 쓰는 물건을 올리고, 다른 친구 물건에 '교환 신청'을 합니다.
 * 물건 주인이 신청을 수락하면 두 물건 모두 '교환 완료'가 되고, 게시판에 도장이 찍힙니다.
 * 교사 화면(PIN)에서 부적절한 글 숨김·삭제, 분류별 등록·완료 통계, '버려질 뻔한 물건 재사용' 수, CSV를 봅니다.
 * 설치 순서는 같은 폴더의 '설치방법.md'를 보세요.
 *
 * 개인정보: 실명·연락처는 받지 않습니다(반 + 별칭만). 비밀번호는 그대로 저장하지 않고 해시(암호화한 값)만 저장합니다.
 */

// ===================== 설정 (여기만 고치세요) =====================
const TEACHER_PIN = '0000';          // ← 바꾸세요! 숫자 6자리 이상 권장. '0000' 그대로면 교사 화면이 열리지 않습니다.
const APP_TITLE = '학급 나눔 장터';
const CLASSES = ['1반', '2반', '3반', '4반'];
const CATEGORIES = ['학용품', '책', '옷·잡화', '생활용품', '취미·놀이', '기타'];
const CONDITIONS = [                 // 물건 상태 3단계 (name은 상·중·하 그대로 두고 desc만 고치세요)
  { name: '상', desc: '거의 새것' },
  { name: '중', desc: '사용감 있지만 멀쩡' },
  { name: '하', desc: '흠집 있지만 쓸 만함' }
];
const MAX_ITEMS_PER_PERSON = 5;      // 한 사람이 동시에 올려 둘 수 있는 물건 수 (교환 끝난 물건은 빼고 셈)
// =================================================================

const SHEET_ITEMS = '물건';
const SHEET_REQS = '신청';
const SHEET_PEOPLE = '사람';
const ITEM_HEADERS = ['등록시각', '물건ID', '반', '별칭', '물건', '분류', '상태', '원하는 교환', '한 줄 이야기', '거래', '숨김', '완료시각'];
const REQ_HEADERS = ['신청시각', '신청ID', '받을 물건ID', '신청 반', '신청 별칭', '내놓은 물건ID', '메시지', '처리', '처리시각'];
const PEOPLE_HEADERS = ['처음 사용', '반', '별칭', '비밀번호 해시'];
const I = { at: 0, id: 1, cls: 2, alias: 3, name: 4, cat: 5, cond: 6, want: 7, story: 8, deal: 9, hidden: 10, doneAt: 11 };
const R = { at: 0, id: 1, target: 2, cls: 3, alias: 4, offer: 5, msg: 6, status: 7, doneAt: 8 };
const OPEN = '교환 가능', DONE = '교환 완료', HIDDEN = '숨김';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 시트 3개를 미리 만들어 둡니다. (안 해도 첫 사용 때 자동 생성) */
function setup() {
  getSheet_(SHEET_ITEMS, ITEM_HEADERS);
  getSheet_(SHEET_REQS, REQ_HEADERS);
  getSheet_(SHEET_PEOPLE, PEOPLE_HEADERS);
}

// ---------------- 학생 화면 ----------------

function getConfig() {
  return { appTitle: APP_TITLE, classes: CLASSES, categories: CATEGORIES, conditions: CONDITIONS, maxItems: MAX_ITEMS_PER_PERSON };
}

/** 게시판: 숨김이 아닌 물건 전체(최신순) + 재사용 수. 비밀번호 해시는 절대 내보내지 않습니다. */
function listItems() {
  const items = readItems_();
  const pending = {};
  readReqs_().forEach(function (q) { if (q.status === '대기') pending[q.target] = (pending[q.target] || 0) + 1; });
  const shown = items.filter(function (it) { return !it.hidden; })
    .map(function (it) { return publicItem_(it, pending[it.id] || 0); })
    .reverse();
  return { items: shown, reused: items.filter(function (it) { return it.done; }).length };
}

/** 물건 올리기. 처음 쓰는 별칭이면 그 비밀번호로 등록됩니다. */
function registerItem(p) {
  const who = validateWho_(p);
  const v = validateItem_(p);
  return withLock_(function () {
    auth_(who, true);
    const mine = readItems_().filter(function (it) { return isMe_(it, who) && !it.done && !it.hidden; });
    if (mine.length >= MAX_ITEMS_PER_PERSON) fail_('한 사람이 동시에 올릴 수 있는 물건은 ' + MAX_ITEMS_PER_PERSON + '개예요. 교환이 끝나면 더 올릴 수 있어요.');
    const id = newId_();
    getSheet_(SHEET_ITEMS, ITEM_HEADERS).appendRow([new Date(), id, who.cls, who.alias, v.name, v.cat, v.cond, v.want, v.story, OPEN, '', '']);
    return { ok: true, id: id };
  });
}

/** 교환 신청. 내 물건(myItemId)을 내놓습니다. 상대가 '그냥 나눔'(원하는 교환 빈칸)이면 내 물건 없이도 신청할 수 있어요. */
function requestSwap(p) {
  const who = validateWho_(p);
  const targetId = cleanId_(p && p.targetId);
  const myItemId = p && p.myItemId ? cleanId_(p.myItemId) : '';
  const msg = clean_(p && p.message, 60, '메시지');
  if (msg.length < 2) fail_('물건 주인에게 보낼 메시지를 한 줄 적어 주세요.');
  return withLock_(function () {
    auth_(who, true);
    const items = readItems_();
    const target = find_(items, targetId);
    if (!target || target.hidden) fail_('그 물건은 지금 게시판에 없어요. 새로고침해 보세요.');
    if (target.done) fail_('이미 교환이 끝난 물건이에요.');
    if (isMe_(target, who)) fail_('내가 올린 물건에는 신청할 수 없어요.');
    if (myItemId) {
      const mine = find_(items, myItemId);
      if (!mine || !isMe_(mine, who)) fail_('내놓을 물건은 내가 올린 물건 중에서 골라 주세요.');
      if (mine.done || mine.hidden) fail_('내놓을 물건이 이미 교환됐거나 숨겨졌어요. 다른 물건을 골라 주세요.');
    } else if (target.want) {
      fail_('이 물건은 교환을 원해요. 먼저 내 물건을 올리고, 내놓을 물건을 골라 주세요.');
    }
    const dup = readReqs_().some(function (q) { return q.target === targetId && q.status === '대기' && q.cls === who.cls && q.alias === who.alias; });
    if (dup) fail_('이미 이 물건에 신청했어요. 내 장터에서 답을 기다려 주세요.');
    const id = newId_();
    getSheet_(SHEET_REQS, REQ_HEADERS).appendRow([new Date(), id, targetId, who.cls, who.alias, myItemId, msg, '대기', '']);
    return { ok: true, id: id };
  });
}

/** 내 장터: 내 물건 + 내 물건에 온 신청 + 내가 보낸 신청. */
function getMine(p) {
  const who = validateWho_(p);
  auth_(who, false);
  const items = readItems_();
  const reqs = readReqs_();
  const byId = {};
  items.forEach(function (it) { byId[it.id] = it; });
  const brief = function (id) { const it = byId[id]; return it ? { id: it.id, name: it.name, cond: it.cond, cat: it.cat, done: it.done, hidden: it.hidden } : null; };
  const mineItems = items.filter(function (it) { return isMe_(it, who); }).reverse();
  const mineIds = {};
  mineItems.forEach(function (it) { mineIds[it.id] = true; });
  return {
    items: mineItems.map(function (it) { return Object.assign(publicItem_(it, 0), { hidden: it.hidden }); }),
    received: reqs.filter(function (q) { return mineIds[q.target]; }).reverse().map(function (q) {
      return { id: q.id, at: q.at, target: brief(q.target), fromCls: q.cls, fromAlias: q.alias, offer: brief(q.offer), message: q.msg, status: q.status };
    }),
    sent: reqs.filter(function (q) { return q.cls === who.cls && q.alias === who.alias; }).reverse().map(function (q) {
      const t = byId[q.target];
      return { id: q.id, at: q.at, target: brief(q.target), owner: t ? t.cls + ' ' + t.alias : '', offer: brief(q.offer), message: q.msg, status: q.status };
    })
  };
}

/** 내 물건에 온 신청에 답하기. accept=true면 두 물건 모두 '교환 완료', 같은 물건에 걸린 다른 신청은 '종료'. */
function respondRequest(p) {
  const who = validateWho_(p);
  const reqId = cleanId_(p && p.reqId);
  const accept = p && p.accept === true;
  return withLock_(function () {
    auth_(who, false);
    const items = readItems_();
    const reqs = readReqs_();
    const q = find_(reqs, reqId);
    if (!q) fail_('그 신청을 찾지 못했어요. 새로고침해 보세요.');
    const target = find_(items, q.target);
    if (!target || !isMe_(target, who)) fail_('내 물건에 온 신청만 답할 수 있어요.');
    if (q.status !== '대기') fail_('이미 처리된 신청이에요.');
    const now = new Date();
    const reqSheet = getSheet_(SHEET_REQS, REQ_HEADERS);
    if (!accept) {
      reqSheet.getRange(q.row, R.status + 1, 1, 2).setValues([['거절', now]]);
      return { ok: true, status: '거절' };
    }
    if (target.done || target.hidden) fail_('이 물건은 이미 교환됐거나 숨겨졌어요.');
    const offer = q.offer ? find_(items, q.offer) : null;
    if (q.offer && (!offer || offer.done || offer.hidden)) fail_('상대가 내놓은 물건이 이미 다른 친구와 교환됐어요. 이 신청은 거절해 주세요.');
    const itemSheet = getSheet_(SHEET_ITEMS, ITEM_HEADERS);
    [target, offer].forEach(function (it) { if (it) itemSheet.getRange(it.row, I.deal + 1).setValue(DONE); });
    [target, offer].forEach(function (it) { if (it) itemSheet.getRange(it.row, I.doneAt + 1).setValue(now); });
    const closed = {};
    closed[target.id] = true;
    if (offer) closed[offer.id] = true;
    reqs.forEach(function (o) {
      if (o.status !== '대기') return;
      if (o.id === q.id) reqSheet.getRange(o.row, R.status + 1, 1, 2).setValues([['수락', now]]);
      else if (closed[o.target] || closed[o.offer]) reqSheet.getRange(o.row, R.status + 1, 1, 2).setValues([['종료', now]]);
    });
    return { ok: true, status: '수락' };
  });
}

// ---------------- 교사 화면 (PIN 확인) ----------------

/** PIN이 맞을 때만 숨긴 물건까지 전부 + 신청 전부를 돌려줍니다. 통계는 화면에서 계산합니다. */
function getTeacherData(pin) {
  checkPin_(pin);
  return {
    config: getConfig(),
    items: readItems_().map(function (it) { return Object.assign(publicItem_(it, 0), { hidden: it.hidden }); }),
    reqs: readReqs_().map(function (q) { return { id: q.id, at: q.at, target: q.target, cls: q.cls, alias: q.alias, offer: q.offer, message: q.msg, status: q.status }; })
  };
}

/** 부적절한 글 숨기기/다시 보이기. 숨긴 물건에 걸린 대기 신청은 그대로 둡니다(다시 보이면 이어서 처리). */
function setHidden(pin, itemId, hidden) {
  checkPin_(pin);
  const id = cleanId_(itemId);
  return withLock_(function () {
    const it = find_(readItems_(), id);
    if (!it) fail_('그 물건을 찾지 못했어요.');
    getSheet_(SHEET_ITEMS, ITEM_HEADERS).getRange(it.row, I.hidden + 1).setValue(hidden === true ? HIDDEN : '');
    return { ok: true };
  });
}

/** 물건 줄을 시트에서 지웁니다(되돌릴 수 없음). 이 물건이 걸린 대기 신청은 '종료'. */
function deleteItem(pin, itemId) {
  checkPin_(pin);
  const id = cleanId_(itemId);
  return withLock_(function () {
    const it = find_(readItems_(), id);
    if (!it) fail_('그 물건을 찾지 못했어요.');
    const reqSheet = getSheet_(SHEET_REQS, REQ_HEADERS);
    const now = new Date();
    readReqs_().forEach(function (q) {
      if (q.status === '대기' && (q.target === id || q.offer === id)) reqSheet.getRange(q.row, R.status + 1, 1, 2).setValues([['종료', now]]);
    });
    getSheet_(SHEET_ITEMS, ITEM_HEADERS).deleteRow(it.row);
    return { ok: true };
  });
}

// ---------------- 내부 도우미 ----------------

function readItems_() {
  return getSheet_(SHEET_ITEMS, ITEM_HEADERS).getDataRange().getValues().slice(1)
    .map(function (r, i) {
      return { row: i + 2, at: toIso_(r[I.at]), id: String(r[I.id]), cls: String(r[I.cls]), alias: String(r[I.alias]), name: plain_(r[I.name]),
               cat: String(r[I.cat]), cond: String(r[I.cond]), want: plain_(r[I.want]), story: plain_(r[I.story]),
               done: String(r[I.deal]) === DONE, hidden: String(r[I.hidden]) === HIDDEN, doneAt: toIso_(r[I.doneAt]) };
    })
    .filter(function (it) { return it.id !== ''; });
}

function readReqs_() {
  return getSheet_(SHEET_REQS, REQ_HEADERS).getDataRange().getValues().slice(1)
    .map(function (r, i) {
      return { row: i + 2, at: toIso_(r[R.at]), id: String(r[R.id]), target: String(r[R.target]), cls: String(r[R.cls]), alias: String(r[R.alias]),
               offer: String(r[R.offer] || ''), msg: plain_(r[R.msg]), status: String(r[R.status]) };
    })
    .filter(function (q) { return q.id !== ''; });
}

function publicItem_(it, requests) {
  return { id: it.id, at: it.at, cls: it.cls, alias: it.alias, name: it.name, cat: it.cat, cond: it.cond,
           want: it.want, story: it.story, done: it.done, doneAt: it.doneAt, requests: requests };
}

/** 반+별칭+비밀번호 확인. 처음 보는 별칭이고 create=true면 그 비밀번호로 등록. 틀리면 10번까지, 그 뒤 10분 잠금. */
function auth_(who, create) {
  const cache = CacheService.getScriptCache();
  const failKey = 'idFail:' + who.cls + '|' + who.alias;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 10) fail_('비밀번호를 너무 많이 틀렸어요. 10분 뒤 다시 해 보세요.');
  const sheet = getSheet_(SHEET_PEOPLE, PEOPLE_HEADERS);
  const rows = sheet.getDataRange().getValues().slice(1);
  const hash = hash_(who);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]) === who.cls && String(rows[i][2]) === who.alias) {
      if (String(rows[i][3]) !== hash) {
        cache.put(failKey, String(fails + 1), 600);
        fail_('비밀번호가 맞지 않아요. (이 별칭을 처음 쓸 때 정한 4자리)');
      }
      return who;
    }
  }
  if (!create) fail_(who.cls + '에 ‘' + who.alias + '’ 별칭으로 올리거나 신청한 기록이 없어요. 반과 별칭을 확인해 주세요.');
  sheet.appendRow([new Date(), who.cls, who.alias, hash]);
  return who;
}

/** 비밀번호 해시: 이 스프레드시트 ID를 소금(salt)으로 섞은 SHA-256. 시트에는 이 값만 저장됩니다. */
function hash_(who) {
  const raw = SpreadsheetApp.getActive().getId() + '|' + who.cls + '|' + who.alias + '|' + who.pw;
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function validateWho_(p) {
  if (!p || typeof p !== 'object') fail_('입력이 비어 있어요.');
  const cls = String(p.cls || '');
  if (CLASSES.indexOf(cls) < 0) fail_('반을 골라 주세요.');
  const alias = clean_(p.alias, 10, '별칭').replace(/\s+/g, ' ');
  if (alias.length < 2) fail_('별칭을 2~10자로 적어 주세요. (실명 말고)');
  if (/^'/.test(alias)) fail_('별칭은 글자로 시작해 주세요.');
  const pw = String(p.pw == null ? '' : p.pw).trim();
  if (!/^\d{4}$/.test(pw)) fail_('비밀번호는 숫자 4자리로 적어 주세요.');
  return { cls: cls, alias: alias, pw: pw };
}

function validateItem_(p) {
  const name = clean_(p.name, 30, '물건 이름');
  if (!name) fail_('물건 이름을 적어 주세요.');
  const cat = String(p.cat || '');
  if (CATEGORIES.indexOf(cat) < 0) fail_('분류를 골라 주세요.');
  const cond = String(p.cond || '');
  if (!CONDITIONS.some(function (c) { return c.name === cond; })) fail_('물건 상태(상·중·하)를 골라 주세요.');
  return { name: name, cat: cat, cond: cond, want: clean_(p.want, 30, '원하는 교환'), story: clean_(p.story, 60, '한 줄 이야기') };
}

function isMe_(it, who) { return it.cls === who.cls && it.alias === who.alias; }
function find_(list, id) { return list.filter(function (x) { return x.id === id; })[0] || null; }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function cleanId_(v) {
  const s = String(v == null ? '' : v);
  if (!/^[0-9a-f]{6,32}$/i.test(s)) fail_('잘못된 요청이에요. 새로고침해 보세요.');
  return s;
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
  if (!lock.tryLock(15000)) fail_('지금 올리는 사람이 많아요. 잠시 뒤 다시 눌러 주세요.');
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

/** 글자 정리: 제어문자 제거, 길이 제한, 전화번호처럼 보이면 거절, 시트 수식 주입 방지. */
function clean_(value, max, label) {
  let s = String(value == null ? '' : value).replace(/[\u0000-\u001f\u007f]/g, ' ').trim();
  if (s.length > max) fail_(label + ' 글자 수는 ' + max + '자 이내로 적어 주세요.');
  if (/01[016789][\s.-]?\d{3,4}[\s.-]?\d{4}/.test(s)) fail_(label + '에 전화번호는 적지 마세요. 연락은 이 장터 안에서 해요.');
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 시트 수식 주입 방지
  return s;
}

function toIso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
function fail_(msg) { throw new Error(msg); }
