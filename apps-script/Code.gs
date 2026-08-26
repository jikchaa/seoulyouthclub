/**
 * YOUTH CLUB SEOUL — 참석 신청 접수
 *
 * 이 스크립트가 하는 일은 하나다. 방문자가 남긴 이름·휴대폰 번호를
 * 스프레드시트에 받아두고, 바깥에는 이름과 순번만 돌려준다.
 * 번호는 어떤 공개 응답에도 실리지 않는다 — 시트와 관리자만 본다.
 *
 * 정원 판정도 여기서 한다. 브라우저에서 세면 얼마든지 우회할 수 있지만,
 * 접수 시점에 시트가 순번을 매기면 실제 선착순이 보장된다.
 *
 * 배포: 확장 프로그램 → Apps Script → 이 코드 붙여넣기 →
 *      배포 → 새 배포 → 유형 '웹 앱' →
 *      실행: 나 / 액세스: 모든 사용자 → 배포 → /exec URL 복사
 */

/* 관리자 화면에서 번호를 불러올 때 쓰는 열쇠. 아무에게도 주지 말 것. */
var ADMIN_KEY = 'XPuTp2Fr52XqvgvHw-V6QOtu71yBsTAG';

/* 정원을 여기서 따로 관리하지 않는다. 공개된 이벤트 데이터가 유일한 출처다. */
var EVENTS_URL = 'https://jikchaa.github.io/seoulyouthclub/data/events.json';

var SHEET_NAME = 'signups';
var HEADERS = ['접수시각', '이벤트ID', '이벤트명', '이름', '휴대폰', '취소'];

/* ------------------------------------------------------------------ */

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(function (r) { return r[1] && !r[5]; });   // 이벤트ID 있고 취소 아님
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* 이벤트 정원표. 매 요청마다 받아오면 느리니 잠깐 캐시한다. */
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
    .filter(function (r) { return String(r[1]) === String(eventId); })
    .map(function (r) { return { at: r[0], name: String(r[3]), phone: String(r[4]) }; });
}

/* ------------------------------------------------------------------ *
 * GET — 공개. 이름과 순번만 나간다.
 *   ?action=list                     모든 이벤트의 명단
 *   ?action=phones&key=ADMIN_KEY     관리자 전용, 번호 포함
 * ------------------------------------------------------------------ */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'list';

  if (action === 'phones') {
    if (params.key !== ADMIN_KEY) {
      return json_({ ok: false, error: 'not authorized' });
    }
    var full = {};
    rows_().forEach(function (r) {
      var id = String(r[1]);
      if (!full[id]) full[id] = [];
      full[id].push({
        name: String(r[3]),
        phone: String(r[4]),
        at: r[0] ? new Date(r[0]).toISOString() : ''
      });
    });
    return json_({ ok: true, signups: full });
  }

  if (action !== 'list') return json_({ ok: false, error: 'unknown action' });

  var out = {};
  rows_().forEach(function (r) {
    var id = String(r[1]);
    if (!out[id]) out[id] = [];
    out[id].push({ name: String(r[3]) });   // 번호는 여기에 절대 넣지 않는다
  });
  return json_({ ok: true, signups: out });
}

/* ------------------------------------------------------------------ *
 * POST — 신청 접수.  { action:'signup', eventId, name, phone }
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
  var phone = String(body.phone || '').replace(/[^0-9]/g, '');

  if (!eventId) return json_({ ok: false, error: '어느 모임인지 알 수 없습니다.' });
  if (name.length < 2) return json_({ ok: false, error: '이름을 두 글자 이상 입력해 주세요.' });
  if (phone.length < 10 || phone.length > 11) {
    return json_({ ok: false, error: '휴대폰 번호를 정확히 입력해 주세요.' });
  }

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
      ''
    ]);

    var seedCount = (meta && meta.seedCount) || 0;
    var position = seedCount + roster.length + 1;   // 화면에 찍히는 번호와 같은 기준
    var capacity = (meta && meta.capacity) || 0;
    var waitlisted = capacity > 0 && position > capacity;

    return json_({
      ok: true,
      position: waitlisted ? position - capacity : position,
      waitlisted: waitlisted
    });
  } finally {
    lock.releaseLock();
  }
}
