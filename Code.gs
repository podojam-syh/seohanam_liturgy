// 전례봉사 가능 날짜 설문 — 구글 시트 Apps Script
// 시트 탭 이름: 응답 (없으면 자동으로 만들어요)
// 코드를 고친 뒤에는 꼭 [배포 → 배포 관리 → 편집 → 새 버전]으로 다시 배포해야 반영돼요.

var SHEET = '응답';
var HEAD = ['제출시각', '달', '이름', '세례명', '학년', '가능날짜', '가능횟수', '메모'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
