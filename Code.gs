// 전례봉사 가능 날짜 설문 — 구글 시트 Apps Script
// 시트 탭 이름: 응답 (없으면 자동으로 만들어요)
// 코드를 고친 뒤에는 꼭 [배포 → 배포 관리 → 편집 → 새 버전]으로 다시 배포해야 반영돼요.

var SHEET_ID = '11JQAo0B6XtxT5dw1RUPhcJajFMq_m-4KuBte758O-to'; // 전례봉사 가능날짜 응답 시트
var SHEET = '응답';
var HEAD = ['제출시각', '달', '이름', '세례명', '학년', '가능날짜', '가능횟수', '메모'];

function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // '2026-12'이 날짜로 바뀌지 않게
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 제출: 같은 달에 같은 아이가 다시 내면 그 줄을 덮어써요.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sh = sheet_();
    var month = String(d.month);
    sh.getRange('B:B').setNumberFormat('@');
    var rows = sh.getDataRange().getDisplayValues();
    (d.kids || []).forEach(function (k) {
      var line = [new Date(), month, k.name, k.baptism || '', k.grade || '',
                  (k.dates || []).join(', '), (k.dates || []).length, d.memo || ''];
      var found = -1;
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][1] === month && rows[i][2] === k.name) { found = i; break; }
      }
      if (found > 0) {
        sh.getRange(found + 1, 1, 1, line.length).setValues([line]);
      } else {
        sh.appendRow(line);
        rows.push(line.map(String));
      }
    });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 조회: ?month=2026-12 → 그 달 응답 목록 (제출 확인, 일정표에서 사용)
function doGet(e) {
  var month = (e && e.parameter && e.parameter.month) || '';
  var rows = sheet_().getDataRange().getDisplayValues().slice(1);
  var out = rows.filter(function (r) { return !month || r[1] === month; })
    .map(function (r) {
      return { time: r[0], month: r[1], name: r[2], baptism: r[3], grade: r[4],
               dates: r[5] ? r[5].split(', ') : [], memo: r[7] };
    });
  return json_({ ok: true, month: month, responses: out });
}

// ===== 배정표 만들기 =====
// Apps Script 편집기에서 함수 목록에 makeAssignSheet 를 고르고 [실행]을 누르면
// 다음 달(15일 전이면 이번 달) 응답으로 'N월 배정' 탭을 새로 만들어요.
// 이미 그 탭이 있으면 지우지 않고 옆에 새 탭으로 만들어요.

var ROSTER = [ // 이름, 세례명, 학년, 포지션
  ['김민희','요안나','중1','독서'], ['정해든','도로티아','중1','독서'], ['김재인','에밀리아','중2','해설'],
  ['김수찬','대건안드레아','초3','기도'], ['정이든','라파엘','초3','기도'], ['한수민','글라라','초3','기도'],
  ['김세민','에디타','초4','해설·독서'], ['윤찬영','제노','초4','독서·기도'], ['장연수','글라라','초4','기도'],
  ['김다민','레베카','초5','독서'], ['김수민','카리타스','초5','독서·기도'], ['손예담','안나','초5','기도'],
  ['김예빈','아나스타시아','초6','기도'], ['김지인','글라라','초6','해설'], ['문리원','다니엘라','초6','기도'],
  ['박제이','사라','초6','독서'], ['오지후','아나스타시아','초6','독서·기도'], ['이세찬','프란치스코','초6','해설'],
  ['이지안','에스테르','초6','해설'], ['임세아','마리스텔라','초6','독서'], ['전예진','라파엘라','초6','독서']
];
var POS_ORDER = ['해설', '독서', '독서·기도', '기도'];

function makeAssignSheet() {
  var d = new Date();
  if (d.getDate() >= 15) d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  makeAssign_(Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM'));
}

function makeAssign_(month) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var y = +month.slice(0, 4), m = +month.slice(5, 7);
  var rows = sheet_().getDataRange().getDisplayValues().slice(1)
    .filter(function (r) { return r[1] === month; });
  var info = {}; ROSTER.forEach(function (k) { info[k[0]] = k; });
  var byName = {}, dates = {};
  rows.forEach(function (r) {
    var ds = r[5] ? r[5].split(', ') : [];
    byName[r[2]] = ds;
    ds.forEach(function (s) { dates[s] = 1; });
  });
  var day = function (s) { return parseInt(s.split('월 ')[1], 10); };
  var dl = Object.keys(dates).sort(function (a, b) { return day(a) - day(b); });
  var rank = function (n) { var p = (info[n] || [])[3] || ''; var i = POS_ORDER.indexOf(p); if (i < 0 && p.indexOf('해설') === 0) i = 0; return i < 0 ? 9 : i; };
  var DOW = '일월화수목금토';

  var cols = dl.map(function (ds) {
    var names = Object.keys(byName).filter(function (n) { return byName[n].indexOf(ds) >= 0; });
    names.sort(function (a, b) { return rank(a) - rank(b) || a.localeCompare(b); });
    return names.map(function (n) {
      var k = info[n] || [n, '', '', ''];
      return n + ' (' + k[2] + '·' + k[3] + ')' + (byName[n].length === 1 ? ' ★' : '');
    });
  });

  var title = m + '월 배정';
  if (ss.getSheetByName(title)) title += ' ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd HHmm');
  var sh = ss.insertSheet(title, 0);

  var w = Math.max(dl.length, 1);
  var maxLen = cols.reduce(function (a, c) { return Math.max(a, c.length); }, 0);
  sh.getRange(1, 1).setValue(m + '월 전례봉사 배정 — 그날 봉사할 아이만 남기고 지우세요').setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1).setValue('필요 인원: 해설 1 · 독서 2 · 보편지향기도 2   |   ★ 이번 달 가능한 날이 하루뿐인 아이').setFontColor('#6b6378');
  if (dl.length) {
    sh.getRange(4, 1, 1, w).setValues([dl.map(function (s) {
      return s + ' (' + DOW[new Date(y, m - 1, day(s)).getDay()] + ')';
    })]).setFontWeight('bold').setBackground('#efe7f7');
    sh.getRange(5, 1, 1, w).setValues([cols.map(function (c) { return '가능 ' + c.length + '명'; })]).setFontColor('#6b6378');
    if (maxLen) {
      var grid = [];
      for (var i = 0; i < maxLen; i++) grid.push(cols.map(function (c) { return c[i] || ''; }));
      sh.getRange(6, 1, maxLen, w).setValues(grid);
    }
  }
  var after = 6 + maxLen + 1;
  var missing = ROSTER.map(function (k) { return k[0]; }).filter(function (n) { return !(n in byName); });
  var none = Object.keys(byName).filter(function (n) { return !byName[n].length; });
  sh.getRange(after, 1).setValue('미제출: ' + (missing.join(', ') || '없음'));
  sh.getRange(after + 1, 1).setValue('이번 달 어려움: ' + (none.join(', ') || '없음'));
  sh.setColumnWidths(1, w, 210);
  sh.setFrozenRows(5);
  ss.setActiveSheet(sh);
}
