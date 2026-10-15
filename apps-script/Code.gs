/**
 * YOUTH CLUB SEOUL — 참석 신청 접수
 *
 * 이 스크립트가 하는 일은 둘이다.
 *   1. 방문자가 남긴 이름·휴대폰 뒷자리 4자리·카카오톡 닉네임을 시트에 받아두고,
 *      바깥에는 이름과 순번만 돌려준다.
 *   2. 신청자가 YCLB Index(회원명부)에 있는지 이름+뒷자리로 맞춰 보고,
 *      맞으면 공개 명단에 verified 표시만 붙인다.
 * 전체 번호는 받지 않는다. 뒷자리는 동명이인을 가리는 데만 쓴다.
 * 뒷자리와 닉네임은 어떤 공개 응답에도 실리지 않는다 — 시트와 관리자만 본다.
 *
 * 정원 판정도 여기서 한다. 브라우저에서 세면 얼마든지 우회할 수 있지만,
 * 접수 시점에 시트가 순번을 매기면 실제 선착순이 보장된다.
 *
 * 관리자 열쇠: 이 코드에 적지 않는다(저장소가 공개라서).
 *   Apps Script 왼쪽 ⚙️ 프로젝트 설정 → 스크립트 속성 → 속성 추가
 *   속성 이름 ADMIN_KEY / 값 = 관리자 화면에 넣는 것과 같은 열쇠
 *   스크립트 속성은 코드를 새로 붙여넣어도 그대로 남는다.
 *
 * 배포: 확장 프로그램 → Apps Script → 이 코드 붙여넣기 →
 *      배포 → 새 배포 → 유형 '웹 앱' →
 *      실행: 나 / 액세스: 모든 사용자 → 배포 → /exec URL 복사
 * 코드를 바꾼 뒤에는: 배포 → 배포 관리 → 연필 → 버전 '새 버전' → 배포
 * (그래야 주소가 그대로 유지된다)
 */

/* 관리자 화면에서 신청자 정보를 불러올 때 쓰는 열쇠. 스크립트 속성에만 둔다.
   비어 있으면 관리자 조회는 아예 닫힌다. */
function adminKey_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_KEY') || '';
}

/* 정원을 여기서 따로 관리하지 않는다. 공개된 이벤트 데이터가 유일한 출처다. */
/* 도메인 연결 상태와 무관하게 읽히도록 GitHub 원본을 직접 읽는다(최대 5분 지연). */
var EVENTS_URL = 'https://raw.githubusercontent.com/jikchaa/seoulyouthclub/main/data/events.json';

var SHEET_NAME = 'signups';
/* 열 순서를 바꾸지 말 것 — 이미 쌓인 행이 이 순서로 들어 있다. 새 열은 끝에만 붙인다. */
var HEADERS = ['접수시각', '이벤트ID', '이벤트명', '이름', '휴대폰 뒷자리', '취소', '카카오톡 닉네임'];
var COL = { at: 0, eventId: 1, title: 2, name: 3, phone: 4, cancel: 5, kakao: 6 };

/* 회원명부 탭. 이 이름 중 하나로 된 탭을 찾고, 없으면 첫 번째 이름으로 만든다. */
var INDEX_SHEETS = ['YCLB Index', 'index', 'Index'];
var INDEX_HEADERS = ['이름', '휴대폰 뒷자리', '카카오톡 닉네임', '이메일', '인스타그램', '소속', '직책', '리멤버/링크드인'];

/* ------------------------------------------------------------------ */

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    return sh;
  }
  /* 예전 버전으로 만든 시트에는 뒤쪽 머리글이 없다. 비어 있는 칸만 채운다. */
  var head = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  for (var i = 0; i < HEADERS.length; i++) {
    if (!head[i]) sh.getRange(1, i + 1).setValue(HEADERS[i]);
  }
  return sh;
}

function rows_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(function (r) { return r[COL.eventId] && !r[COL.cancel]; });   // 이벤트ID 있고 취소 아님
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 같은 사람인지 비교하기 위한 정리 ---------------- */

/* 휴대폰 뒷자리 4자리. 하이픈·공백 상관없이 숫자만 보고 끝 네 자리를 쓴다.
   명부에 전체 번호가 있어도, 뒷자리만 있어도 똑같이 맞춰진다.
   시트가 숫자로 읽어 앞자리 0 을 먹은 경우(0123 → 123)는 0 을 채워 되살린다. */
function last4_(v) {
  var d = String(v == null ? '' : v).replace(/[^0-9]/g, '');
  if (!d) return '';
  if (d.length < 4) d = ('0000' + d).slice(-4);
  return d.slice(-4);
}

/* 띄어쓰기와 대소문자 차이는 같은 이름으로 본다. (김 민수 = 김민수, Amy = amy) */
function normName_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '').toLowerCase();
}

function key_(name, phone) { return normName_(name) + '|' + last4_(phone); }

/* ---------------- 회원명부 ---------------- */

function indexSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < INDEX_SHEETS.length; i++) {
    var sh = ss.getSheetByName(INDEX_SHEETS[i]);
    if (sh) return sh;
  }
  var created = ss.insertSheet(INDEX_SHEETS[0]);
  created.appendRow(INDEX_HEADERS);
  created.setFrozenRows(1);
  return created;
}

/* 머리글 이름으로 열을 찾는다 — 구글폼 응답 시트처럼 열 순서가 달라도 된다. */
function findCol_(head, words) {
  for (var i = 0; i < head.length; i++) {
    var h = String(head[i] || '').replace(/\s+/g, '').toLowerCase();
    for (var j = 0; j < words.length; j++) if (h.indexOf(words[j]) >= 0) return i;
  }
  return -1;
}

/* 명부의 이름+뒷자리 목록. 매 요청마다 시트를 읽으면 느리니 1분 캐시한다. */
function indexKeys_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('indexKeys');
  if (hit) return JSON.parse(hit);

  var keys = {};
  var sh = indexSheet_();
  var last = sh.getLastRow(), width = sh.getLastColumn();
  if (last >= 2 && width >= 1) {
    var all = sh.getRange(1, 1, last, width).getValues();
    var head = all[0];
    var cName = findCol_(head, ['이름', 'name']);
    var cPhone = findCol_(head, ['휴대폰', '전화', '연락처', 'phone', 'mobile']);
    if (cName >= 0 && cPhone >= 0) {
      for (var r = 1; r < all.length; r++) {
        var n = all[r][cName], p = all[r][cPhone];
        if (n && p) keys[key_(n, p)] = 1;
      }
    }
  }
  /* Index 신청서 응답 시트가 연결돼 있으면 그쪽 신청자도 인증으로 본다. */
  if (prop_('INDEX_SHEET_ID')) {
    try {
      indexPeople_().forEach(function (p) { if (p._phone) keys[key_(p.name, p._phone)] = 1; });
    } catch (err) { /* 시트를 못 열어도 기존 명부 탭으로 계속 */ }
  }
  cache.put('indexKeys', JSON.stringify(keys), 60);
  return keys;
}

/* ---------------- 이벤트 정원표 ---------------- */

/* 매 요청마다 받아오면 느리니 잠깐 캐시한다. */
function capacities_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('caps');
  if (hit) return JSON.parse(hit);

  var caps = {};
  try {
    var res = UrlFetchApp.fetch(EVENTS_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var data = JSON.parse(res.getContentText());
      (data.events || []).forEach(function (ev) {
        caps[ev.id] = {
          capacity: ev.capacity || 0,
          title: ev.title || '',
          status: ev.status || 'auto',
          date: ev.date || '',
          /* 카톡에서 받아 적어 events.json 에 들어간 사람들. 이들이 먼저 신청한
             것이므로 순번을 셀 때 반드시 앞에 놓아야 화면과 숫자가 맞는다. */
          seedCount: (ev.attendees || []).length
        };
      });
    }
  } catch (err) {
    /* 정원표를 못 받아도 접수는 계속된다 — 정원 없는 모임처럼 처리한다. */
  }
  cache.put('caps', JSON.stringify(caps), 60);
  return caps;
}

/* 한 이벤트의 접수 명단을 순서대로. */
function rosterOf_(eventId, all) {
  return (all || rows_())
    .filter(function (r) { return String(r[COL.eventId]) === String(eventId); })
    .map(function (r) {
      return { at: r[COL.at], name: String(r[COL.name]), key: key_(r[COL.name], r[COL.phone]) };
    });
}

/* ------------------------------------------------------------------ *
 * GET — 공개. 이름·순번·인증 여부만 나간다.
 *   ?action=list                     모든 이벤트의 명단
 *   ?action=phones&key=ADMIN_KEY     관리자 전용, 뒷자리·닉네임 포함
 * ------------------------------------------------------------------ */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'list';
  var index = indexKeys_();

  if (action === 'phones') {
    var k = adminKey_();
    if (!k || params.key !== k) {
      return json_({ ok: false, error: 'not authorized' });
    }
    var full = {};
    rows_().forEach(function (r) {
      var id = String(r[COL.eventId]);
      if (!full[id]) full[id] = [];
      full[id].push({
        name: String(r[COL.name]),
        phone: last4_(r[COL.phone]),
        kakao: String(r[COL.kakao] || ''),
        verified: !!index[key_(r[COL.name], r[COL.phone])],
        at: r[COL.at] ? new Date(r[COL.at]).toISOString() : ''
      });
    });
    return json_({ ok: true, signups: full });
  }

  if (action !== 'list') return json_({ ok: false, error: 'unknown action' });

  /* 지난 모임의 명단은 공개 응답에 싣지 않는다(관리자 조회 ?action=phones 에는 그대로 있다). */
  var caps = capacities_();
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var out = {};
  rows_().forEach(function (r) {
    var id = String(r[COL.eventId]);
    if (caps[id] && caps[id].date && caps[id].date < today) return;
    if (!out[id]) out[id] = [];
    /* 뒷자리·닉네임은 여기에 절대 넣지 않는다. 인증 여부는 결과만 싣는다. */
    out[id].push({
      name: String(r[COL.name]),
      verified: !!index[key_(r[COL.name], r[COL.phone])]
    });
  });
  return json_({ ok: true, signups: out });
}

/* ------------------------------------------------------------------ *
 * POST — 신청 접수.  { action:'signup', eventId, name, phone, kakao }
 * ------------------------------------------------------------------ */

function doPost(e) {
  var body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '요청을 읽지 못했습니다.' });
  }

  if (body.action === 'index') return indexPost_(body);
  if (body.action !== 'signup') return json_({ ok: false, error: 'unknown action' });

  var eventId = String(body.eventId || '').trim();
  var name = String(body.name || '').trim().slice(0, 20);
  /* 뒷자리 4자리만 받는다. 예전 화면이 전체 번호를 보내도 끝 네 자리만 남긴다. */
  var digits = String(body.phone || '').replace(/[^0-9]/g, '');
  var phone = digits.slice(-4);
  var kakao = String(body.kakao || '').trim().slice(0, 30);

  if (!eventId) return json_({ ok: false, error: '어느 모임인지 알 수 없습니다.' });
  if (name.length < 2) return json_({ ok: false, error: '이름을 두 글자 이상 입력해 주세요.' });
  if (!(digits.length === 4 || digits.length === 10 || digits.length === 11)) {
    return json_({ ok: false, error: '휴대폰 번호 뒷자리 4자리를 입력해 주세요.' });
  }
  if (!kakao) return json_({ ok: false, error: '카카오톡 닉네임을 입력해 주세요.' });

  /* 동시에 두 명이 마지막 자리를 누르면 순번이 겹친다. 한 줄로 세운다. */
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return json_({ ok: false, error: '잠시 뒤 다시 시도해 주세요.' });
  }

  try {
    var caps = capacities_();
    var meta = caps[eventId];

    if (meta && meta.status === 'closed') {
      return json_({ ok: false, error: '마감된 모임입니다.' });
    }
    if (meta && meta.date) {
      var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
      if (meta.date < today) return json_({ ok: false, error: '이미 지난 모임입니다.' });
    }

    var all = rows_();
    var roster = rosterOf_(eventId, all);

    /* 이름과 뒷자리가 모두 같으면 같은 사람이다. 동명이인은 뒷자리로 갈린다. */
    var me = key_(name, phone);
    for (var i = 0; i < roster.length; i++) {
      if (roster[i].key === me) {
        var at = ((meta && meta.seedCount) || 0) + i + 1;
        return json_({ ok: false, error: '이미 신청하셨습니다 — ' + at + '번째입니다.' });
      }
    }

    sheet_().appendRow([
      new Date(),
      eventId,
      (meta && meta.title) || '',
      name,
      "'" + phone,          // 0123 처럼 0으로 시작해도 날아가지 않게 문자열로 넣는다
      '',
      kakao
    ]);

    var seedCount = (meta && meta.seedCount) || 0;
    var position = seedCount + roster.length + 1;   // 화면에 찍히는 번호와 같은 기준
    var capacity = (meta && meta.capacity) || 0;
    var waitlisted = capacity > 0 && position > capacity;

    return json_({
      ok: true,
      position: waitlisted ? position - capacity : position,
      waitlisted: waitlisted,
      verified: !!indexKeys_()[key_(name, phone)]
    });
  } finally {
    lock.releaseLock();
  }
}

/* ================================================================== *
 * YCLB Index — 회원명부 열람
 *
 * 명부 화면(yclb-index.html)에는 데이터가 없다. 이 스크립트가 본인 확인을 한
 * 사람에게만 명부를 내려준다.
 *
 *   본인 확인: Index 신청서에 적은 이름 + 휴대폰 뒷자리 4자리 + 이메일이
 *             모두 같고, 두 동의 문항에 동의한 행이 있을 때만 통과.
 *   명부 내용: Index 신청자 = 공개 항목 전체(이메일·휴대폰·거주지 제외)
 *             그 외 가입자 = 성 + ** , 소속, 산업
 *   같은 사람: 휴대폰 번호가 같으면 같은 사람. Index 신청자는 가입 명단에서 뺀다.
 *
 * 스크립트 속성 (⚙️ 프로젝트 설정 → 스크립트 속성):
 *   INDEX_SHEET_ID    Index 신청서 응답 스프레드시트 ID  (주소의 /d/ 와 /edit 사이)
 *   MEMBERS_SHEET_ID  가입 신청서 응답 스프레드시트 ID
 *   INDEX_ADMIN_KEY   명부 관리자 비밀번호 (호스트 체크용)
 *   INDEX_SECRET      비워 두면 처음 실행 때 자동으로 만든다. 바꾸면 모두 다시 로그인해야 한다.
 *
 * 가입 신청서 응답 스프레드시트에 탭 두 개가 자동으로 생긴다(응답 탭은 건드리지 않는다):
 *   index_hosts  호스트 경험자. 명부 화면의 관리자 모드에서 체크하면 여기 쌓인다.
 *   org_map      가입 신청서 '소속' 원문 → 화면에 보일 소속 / 산업.
 *                처음 보는 소속은 '표시 소속'이 빈 채로 자동 추가되고, 채울 때까지 숨긴다.
 *                (소속 칸에 실명을 같이 적는 사람이 있어서다.)
 * ================================================================== */

var IDX_HOSTS_SHEET = 'index_hosts';
var IDX_HOSTS_HEADERS = ['이름', '휴대폰', '체크 시각'];
var IDX_ORGMAP_SHEET = 'org_map';
var IDX_ORGMAP_HEADERS = ['소속 원문', '표시 소속 (비우면 숨김)', '산업', '추가된 시각'];
var IDX_TOKEN_DAYS = 7;
/* 명부를 서버 캐시에 미리 만들어 두는 시간. 5분마다 도는 트리거(installIndexRefresh)가
   다시 채우므로, 로그인은 시트를 읽지 않고 캐시에서 바로 답한다. */
var IDX_CACHE_SEC = 600;
var IDX_EMPTY = { '': 1, '-': 1, 'na': 1, 'n/a': 1, '없음': 1, '네': 1, 'x': 1, '비공개': 1, 'none': 1 };

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }

function idxSecret_() {
  var s = prop_('INDEX_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('INDEX_SECRET', s);
  }
  return s;
}

function hmac_(text) {
  var raw = Utilities.computeHmacSha256Signature(text, idxSecret_());
  return Utilities.base64EncodeWebSafe(raw).replace(/=+$/, '');
}

/* 휴대폰 전체 번호(숫자만). +82, 앞자리 0 빠짐을 되살린다. */
function phoneDigits_(v) {
  var d = String(v == null ? '' : v).replace(/[^0-9]/g, '');
  if (d.indexOf('82') === 0) d = '0' + d.slice(2);
  if (d.length === 10 && d.indexOf('10') === 0) d = '0' + d;
  return d;
}

/* 사람 id: 휴대폰 번호의 HMAC. 화면에는 번호 대신 이 값만 나간다. */
function personId_(phone, fallback) {
  var d = phoneDigits_(phone);
  return hmac_(d ? 'p:' + d : 'f:' + fallback).slice(0, 16);
}

function clean_(v) {
  var s = String(v == null ? '' : v).trim();
  return IDX_EMPTY[s.toLowerCase()] ? '' : s;
}
/* '(비공개) ...' 처럼 적은 칸은 통째로 숨긴다. */
function priv_(v) { var s = clean_(v); return s.indexOf('비공개') >= 0 ? '' : s; }

function igId_(v) {
  var s = clean_(v).replace(/^@/, '');
  var m = s.match(/instagram\.com\/([^\/?#]+)/);
  if (m) return m[1];
  return /^[A-Za-z0-9._]{1,30}$/.test(s) ? s : '';
}

function linkedin_(v) {
  var s = clean_(v);
  if (s.indexOf('linkedin.com/') < 0) return '';
  s = s.split('?')[0].replace(/\/+$/, '');
  return /^https?:\/\//.test(s) ? s : 'https://' + s;
}

function nick_(v, name) {
  var s = clean_(v);
  if (!s || s === name) return '';
  if (s.indexOf('->') >= 0) s = s.split('->').pop().trim();
  return s === name ? '' : s;
}

function multi_(v) {
  return clean_(v).split(',').map(function (x) { return x.trim(); }).filter(String);
}

function mask_(name) {
  var s = String(name || '').trim();
  return s ? s.charAt(0).toUpperCase() + '**' : '';
}

function consented_(v) { return String(v || '').indexOf('동의') >= 0; }

/* 머리글로 열 찾기. 공백 무시, 앞에서부터 처음 맞는 열. */
function colOf_(head, words) {
  var norm = head.map(function (h) { return String(h || '').replace(/\s+/g, '').toLowerCase(); });
  for (var j = 0; j < words.length; j++) {
    for (var i = 0; i < norm.length; i++) if (norm[i].indexOf(words[j]) >= 0) return i;
  }
  return -1;
}

/* 응답 탭: 이 스크립트가 만든 탭(org_map, index_hosts)을 뺀 첫 탭. */
function readSheet_(id) {
  if (!id) return null;
  var sh = SpreadsheetApp.openById(id).getSheets().filter(function (s) {
    var n = s.getName(); return n !== IDX_ORGMAP_SHEET && n !== IDX_HOSTS_SHEET;
  })[0];
  if (!sh) return { head: [], rows: [] };
  var last = sh.getLastRow(), width = sh.getLastColumn();
  if (last < 1 || width < 1) return { head: [], rows: [] };
  var all = sh.getRange(1, 1, last, width).getValues();
  return { head: all[0], rows: all.slice(1) };
}

/* 운영 탭은 가입 신청서 응답 스프레드시트에 둔다. 맨 뒤에 붙인다. */
function tab_(name, headers) {
  var ss = SpreadsheetApp.openById(prop_('MEMBERS_SHEET_ID'));
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name, ss.getSheets().length);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* Index 신청자 목록. 로그인·토큰 확인·명부 만들기가 모두 이걸 쓰므로 1분 캐시한다.
   (스크립트 캐시는 서버 안에만 있다. 화면으로 나가지 않는다.) */
function indexPeople_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('indexPeople');
  if (hit) return JSON.parse(hit);
  var people = indexPeopleFromSheet_();
  try { cache.put('indexPeople', JSON.stringify(people), IDX_CACHE_SEC); } catch (err) { /* 너무 크면 캐시 없이 */ }
  return people;
}

/* Index 신청서 응답 → 동의한 사람만, 같은 번호는 최신 응답으로. */
function indexPeopleFromSheet_() {
  var data = readSheet_(prop_('INDEX_SHEET_ID'));
  if (!data) return [];
  var h = data.head, C = {
    ts: colOf_(h, ['timestamp', '타임스탬프']), name: colOf_(h, ['이름(실명)', '이름']),
    nick: colOf_(h, ['카카오톡닉네임']), email: colOf_(h, ['이메일']), phone: colOf_(h, ['휴대폰']),
    org: colOf_(h, ['소속기업', '소속']), dept: colOf_(h, ['부서또는', '부서']), title: colOf_(h, ['직책']),
    ind: colOf_(h, ['산업/업종', '산업']), job: colOf_(h, ['직무분야', '직무']), work: colOf_(h, ['근무지']),
    li: colOf_(h, ['linkedin']), ig: colOf_(h, ['인스타그램']), want: colOf_(h, ['네트워킹하고']),
    give: colOf_(h, ['도움이나가치', '도움']), c1: colOf_(h, ['정보공개동의']), c2: colOf_(h, ['개인정보수집'])
  };
  var g = function (r, k) { return C[k] >= 0 ? r[C[k]] : ''; };
  var byKey = {};
  data.rows.forEach(function (r) {
    var name = clean_(g(r, 'name'));
    if (!name || !consented_(g(r, 'c1')) || !consented_(g(r, 'c2'))) return;
    var phone = phoneDigits_(g(r, 'phone'));
    var email = String(g(r, 'email') || '').trim().toLowerCase();
    var key = phone || ('e:' + email);
    var prev = byKey[key];
    if (prev && new Date(prev._ts) > new Date(g(r, 'ts'))) return;
    byKey[key] = {
      _ts: g(r, 'ts'), _phone: phone, _email: email,
      id: personId_(phone, 'e:' + email),
      name: name, nick: nick_(g(r, 'nick'), name), org: clean_(g(r, 'org')),
      dept: priv_(g(r, 'dept')), title: priv_(g(r, 'title')),
      ind: clean_(g(r, 'ind')) || '기타', job: clean_(g(r, 'job')) || '기타', work: clean_(g(r, 'work')),
      li: linkedin_(g(r, 'li')), ig: igId_(g(r, 'ig')), want: multi_(g(r, 'want')), give: clean_(g(r, 'give'))
    };
  });
  return Object.keys(byKey).map(function (k) { return byKey[k]; });
}

function orgMap_() {
  var sh = tab_(IDX_ORGMAP_SHEET, IDX_ORGMAP_HEADERS);
  var last = sh.getLastRow(), map = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 3).getValues().forEach(function (r) {
      var raw = String(r[0] || '').trim();
      if (raw) map[raw] = { org: String(r[1] || '').trim(), ind: String(r[2] || '').trim() };
    });
  }
  return { sheet: sh, map: map };
}

/* 가입 신청서 응답 → Index 신청자를 뺀 나머지. 같은 번호는 최신 응답으로. */
function joinPeople_(indexPhones) {
  var data = readSheet_(prop_('MEMBERS_SHEET_ID'));
  if (!data) return [];
  var h = data.head;
  var cTs = colOf_(h, ['timestamp', '타임스탬프']), cName = colOf_(h, ['성함', '이름']);
  var cPhone = colOf_(h, ['휴대폰']), cOrg = colOf_(h, ['소속']);
  var latest = {};
  data.rows.forEach(function (r) {
    var name = String(r[cName] || '').trim();
    if (!name) return;
    var phone = phoneDigits_(r[cPhone]);
    var key = phone || ('n:' + name);
    var prev = latest[key];
    if (!prev || new Date(r[cTs]) >= new Date(prev[cTs])) latest[key] = r;
  });
  var om = orgMap_(), unknown = [];
  var out = Object.keys(latest).filter(function (k) { return !indexPhones[k]; }).map(function (k) {
    var r = latest[k], raw = cOrg >= 0 ? String(r[cOrg] || '').trim() : '';
    var m = raw ? om.map[raw] : null;
    if (raw && !m) { unknown.push(raw); om.map[raw] = m = { org: '', ind: '' }; }
    return { id: personId_(phoneDigits_(r[cPhone]), k), mask: mask_(r[cName]), org: m ? m.org : '', ind: m && m.org ? m.ind : '' };
  });
  if (unknown.length) {
    var now = new Date();
    om.sheet.getRange(om.sheet.getLastRow() + 1, 1, unknown.length, 4)
      .setValues(unknown.map(function (u) { return [u, '', '', now]; }));
  }
  return out;
}

function hostPhones_() {
  var sh = tab_(IDX_HOSTS_SHEET, IDX_HOSTS_HEADERS);
  var last = sh.getLastRow(), set = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    var d = phoneDigits_(r[1]); if (d) set[d] = 1;
  });
  return set;
}

/* 화면에 내려줄 명부. 캐시해 두고, 호스트를 바꾸거나 트리거가 돌 때 새로 만든다. */
function indexPayload_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('indexPayload');
  if (hit) return JSON.parse(hit);

  var people = indexPeople_(), phones = {};
  people.forEach(function (p) { if (p._phone) phones[p._phone] = 1; });
  var hosts = hostPhones_(), hostIds = [];
  people.forEach(function (p) { if (p._phone && hosts[p._phone]) hostIds.push(p.id); });
  Object.keys(hosts).forEach(function (d) { if (!phones[d]) hostIds.push(personId_(d)); });
  var members = people.map(function (p) {
    var o = {}; Object.keys(p).forEach(function (k) { if (k.charAt(0) !== '_') o[k] = p[k]; }); return o;
  });
  var payload = { members: members, guests: joinPeople_(phones), hosts: hostIds };
  try { cache.put('indexPayload', JSON.stringify(payload), IDX_CACHE_SEC); } catch (err) { /* 100KB 넘으면 캐시 없이 */ }
  return payload;
}

/* ---------------- 로그인 토큰 ---------------- */

function makeToken_(id) {
  var body = id + '.' + (Date.now() + IDX_TOKEN_DAYS * 864e5);
  return body + '.' + hmac_('t:' + body);
}

/* 서명과 만료를 보고, 지금도 동의한 Index 신청자인지 확인한다. */
function checkToken_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 3) return null;
  var body = parts[0] + '.' + parts[1];
  if (hmac_('t:' + body) !== parts[2] || Number(parts[1]) < Date.now()) return null;
  var ok = indexPeople_().some(function (p) { return p.id === parts[0]; });
  return ok ? parts[0] : null;
}

/* ---------------- 무차별 대입 막기 ---------------- */

function tooMany_(key, limit) {
  return Number(CacheService.getScriptCache().get(key) || 0) >= limit;
}
function bump_(key, seconds) {
  var c = CacheService.getScriptCache();
  c.put(key, String(Number(c.get(key) || 0) + 1), seconds);
}

function adminOk_(key) {
  var k = prop_('INDEX_ADMIN_KEY');
  if (!k || tooMany_('idxAdminFail', 5)) return false;
  if (String(key || '') === k) return true;
  bump_('idxAdminFail', 900);
  return false;
}

/* ------------------------------------------------------------------ *
 * POST { action: 'index', ... }
 *   op 'login'  : name, tail, email      → token + 명부
 *   op 'data'   : token  또는 adminKey   → 명부
 *   op 'admin'  : adminKey               → 관리자 확인
 *   op 'host'   : adminKey, id, on       → 호스트 체크/해제
 * ------------------------------------------------------------------ */

function indexPost_(body) {
  var op = body.op;

  if (op === 'login') {
    var name = normName_(body.name), tail = last4_(body.tail);
    var email = String(body.email || '').trim().toLowerCase();
    if (!name || tail.length !== 4 || email.indexOf('@') < 1) {
      return json_({ ok: false, error: '이름, 휴대폰 뒷자리 4자리, 이메일을 모두 입력해 주세요.' });
    }
    if (tooMany_('idxFail:' + name, 5) || tooMany_('idxFailAll', 60)) {
      return json_({ ok: false, error: '시도가 너무 많습니다. 15분 뒤 다시 시도해 주세요.' });
    }
    var me = null;
    indexPeople_().forEach(function (p) {
      if (normName_(p.name) === name && last4_(p._phone) === tail && p._email === email) me = p;
    });
    if (!me) {
      bump_('idxFail:' + name, 900); bump_('idxFailAll', 600);
      return json_({ ok: false, error: 'Index 신청서의 정보와 일치하지 않습니다. 신청서에 적은 이름, 뒷자리, 이메일을 확인해 주세요.' });
    }
    return json_({ ok: true, token: makeToken_(me.id), me: me.id, data: indexPayload_() });
  }

  if (op === 'data') {
    if (body.adminKey ? adminOk_(body.adminKey) : checkToken_(body.token)) {
      return json_({ ok: true, data: indexPayload_() });
    }
    return json_({ ok: false, expired: true, error: '다시 본인 확인을 해 주세요.' });
  }

  /* 화면이 열리자마자 부른다. 사람이 입력하는 동안 서버를 깨우고 캐시를 채워 둔다. 데이터는 안 나간다. */
  if (op === 'warm') {
    indexPayload_();
    return json_({ ok: true });
  }

  if (op === 'admin') {
    return json_(adminOk_(body.adminKey) ? { ok: true } : { ok: false, error: '비밀번호가 맞지 않거나 시도가 너무 많습니다.' });
  }

  if (op === 'host') {
    if (!adminOk_(body.adminKey)) return json_({ ok: false, error: '관리자 확인에 실패했습니다.' });
    var id = String(body.id || ''), on = !!body.on;
    /* id → 휴대폰 번호. 명부 전체에서 같은 id를 찾는다. */
    var found = null;
    indexPeople_().forEach(function (p) { if (p.id === id) found = { name: p.name, phone: p._phone }; });
    if (!found) {
      var data = readSheet_(prop_('MEMBERS_SHEET_ID'));
      if (data) {
        var cName = colOf_(data.head, ['성함', '이름']), cPhone = colOf_(data.head, ['휴대폰']);
        data.rows.forEach(function (r) {
          var d = phoneDigits_(r[cPhone]);
          if (d && personId_(d) === id) found = { name: String(r[cName] || ''), phone: d };
        });
      }
    }
    if (!found || !found.phone) return json_({ ok: false, error: '대상을 찾지 못했습니다.' });

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var sh = tab_(IDX_HOSTS_SHEET, IDX_HOSTS_HEADERS);
      var last = sh.getLastRow();
      var rows = last >= 2 ? sh.getRange(2, 2, last - 1, 1).getValues() : [];
      for (var i = rows.length - 1; i >= 0; i--) {
        if (phoneDigits_(rows[i][0]) === found.phone) sh.deleteRow(i + 2);
      }
      if (on) sh.appendRow([found.name, "'" + found.phone, new Date()]);
    } finally {
      lock.releaseLock();
    }
    CacheService.getScriptCache().remove('indexPayload');
    indexPayload_();   // 바로 다시 채워 둔다
    return json_({ ok: true });
  }

  return json_({ ok: false, error: 'unknown op' });
}

/* ------------------------------------------------------------------ *
 * 속도: 명부를 미리 만들어 두는 트리거
 *
 * Apps Script 편집기 위쪽 함수 선택 칸에서 installIndexRefresh 를 고르고
 * [실행] 을 한 번 누른다(권한 승인). 그 뒤로는 5분마다 refreshIndexCache 가
 * 돌아 시트를 읽고 캐시를 채운다. 새 신청은 5분 안에 명부에 나타난다.
 * ------------------------------------------------------------------ */

function refreshIndexCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['indexPeople', 'indexPayload']);
  indexPeople_();
  indexPayload_();
}

function installIndexRefresh() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'refreshIndexCache') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('refreshIndexCache').timeBased().everyMinutes(5).create();
  refreshIndexCache();
}
