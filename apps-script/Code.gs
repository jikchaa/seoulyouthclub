/**
 * YOUTH CLUB SEOUL — 참석 신청 접수
 *
 * 이 스크립트가 하는 일은 둘이다.
 *   1. 방문자가 남긴 이름·휴대폰 번호·카카오톡 닉네임을 시트에 받아두고,
 *      바깥에는 이름과 순번만 돌려준다.
 *   2. 신청자가 YCLB Index(회원명부)에 있는지 이름+번호로 맞춰 보고,
 *      맞으면 공개 명단에 verified 표시만 붙인다.
 * 번호와 닉네임은 어떤 공개 응답에도 실리지 않는다 — 시트와 관리자만 본다.
 *
 * 정원 판정도 여기서 한다. 브라우저에서 세면 얼마든지 우회할 수 있지만,
 * 접수 시점에 시트가 순번을 매기면 실제 선착순이 보장된다.
 *
 * 배포: 확장 프로그램 → Apps Script → 이 코드 붙여넣기 →
 *      배포 → 새 배포 → 유형 '웹 앱' →
 *      실행: 나 / 액세스: 모든 사용자 → 배포 → /exec URL 복사
 * 코드를 바꾼 뒤에는: 배포 → 배포 관리 → 연필 → 버전 '새 버전' → 배포
 * (그래야 주소가 그대로 유지된다)
 */

/* 관리자 화면에서 번호를 불러올 때 쓰는 열쇠. 아무에게도 주지 말 것. */
var ADMIN_KEY = 'XPuTp2Fr52XqvgvHw-V6QOtu71yBsTAG';

/* 정원을 여기서 따로 관리하지 않는다. 공개된 이벤트 데이터가 유일한 출처다. */
var EVENTS_URL = 'https://jikchaa.github.io/seoulyouthclub/data/events.json';

var SHEET_NAME = 'signups';
/* 열 순서를 바꾸지 말 것 — 이미 쌓인 행이 이 순서로 들어 있다. 새 열은 끝에만 붙인다. */
var HEADERS = ['접수시각', '이벤트ID', '이벤트명', '이름', '휴대폰', '취소', '카카오톡 닉네임'];
var COL = { at: 0, eventId: 1, title: 2, name: 3, phone: 4, cancel: 5, kakao: 6 };

/* 회원명부 탭. 이 이름 중 하나로 된 탭을 찾고, 없으면 첫 번째 이름으로 만든다. */
var INDEX_SHEETS = ['YCLB Index', 'index', 'Index'];
var INDEX_HEADERS = ['이름', '휴대폰', '카카오톡 닉네임', '이메일', '인스타그램', '소속', '직책', '리멤버/링크드인'];

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

/* 하이픈·공백·괄호 상관없이 숫자만 남긴다.
   +82 로 시작하면 0 으로 바꾸고, 시트가 앞자리 0 을 먹어 10자리가 된 번호도 되살린다. */
function normPhone_(v) {
  var d = String(v == null ? '' : v).replace(/[^0-9]/g, '');
  if (d.indexOf('82') === 0 && d.length >= 11) d = '0' + d.slice(2);
  if (d.length === 10 && d.charAt(0) === '1') d = '0' + d;
  return d;
}

/* 띄어쓰기와 대소문자 차이는 같은 이름으로 본다. (김 민수 = 김민수, Amy = amy) */
function normName_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '').toLowerCase();
}

function key_(name, phone) { return normName_(name) + '|' + normPhone_(phone); }

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

/* 명부의 이름+번호 목록. 매 요청마다 시트를 읽으면 느리니 1분 캐시한다. */
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
      return { at: r[COL.at], name: String(r[COL.name]), phone: normPhone_(r[COL.phone]) };
    });
}

/* ------------------------------------------------------------------ *
 * GET — 공개. 이름·순번·인증 여부만 나간다.
 *   ?action=list                     모든 이벤트의 명단
 *   ?action=phones&key=ADMIN_KEY     관리자 전용, 번호·닉네임 포함
 * ------------------------------------------------------------------ */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'list';
  var index = indexKeys_();

  if (action === 'phones') {
    if (params.key !== ADMIN_KEY) {
      return json_({ ok: false, error: 'not authorized' });
    }
    var full = {};
    rows_().forEach(function (r) {
      var id = String(r[COL.eventId]);
      if (!full[id]) full[id] = [];
      full[id].push({
        name: String(r[COL.name]),
        phone: normPhone_(r[COL.phone]),
        kakao: String(r[COL.kakao] || ''),
        verified: !!index[key_(r[COL.name], r[COL.phone])],
        at: r[COL.at] ? new Date(r[COL.at]).toISOString() : ''
      });
    });
    return json_({ ok: true, signups: full });
  }

  if (action !== 'list') return json_({ ok: false, error: 'unknown action' });

  var out = {};
  rows_().forEach(function (r) {
    var id = String(r[COL.eventId]);
    if (!out[id]) out[id] = [];
    /* 번호·닉네임은 여기에 절대 넣지 않는다. 인증 여부는 결과만 싣는다. */
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

  if (body.action !== 'signup') return json_({ ok: false, error: 'unknown action' });

  var eventId = String(body.eventId || '').trim();
  var name = String(body.name || '').trim().slice(0, 20);
  var phone = normPhone_(body.phone);
  var kakao = String(body.kakao || '').trim().slice(0, 30);

  if (!eventId) return json_({ ok: false, error: '어느 모임인지 알 수 없습니다.' });
  if (name.length < 2) return json_({ ok: false, error: '이름을 두 글자 이상 입력해 주세요.' });
  if (phone.length < 10 || phone.length > 11) {
    return json_({ ok: false, error: '휴대폰 번호를 정확히 입력해 주세요.' });
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

    /* 같은 번호면 같은 사람이다 — 하이픈 유무와 상관없이 비교한다. */
    for (var i = 0; i < roster.length; i++) {
      if (roster[i].phone === phone) {
        var at = ((meta && meta.seedCount) || 0) + i + 1;
        return json_({ ok: false, error: '이미 신청하셨습니다 — ' + at + '번째입니다.' });
      }
    }

    sheet_().appendRow([
      new Date(),
      eventId,
      (meta && meta.title) || '',
      name,
      "'" + phone,          // 앞자리 0이 날아가지 않게 문자열로 넣는다
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
