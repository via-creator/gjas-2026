// ── GJAS 2026 사전답사검증표 GAS ──
// Sheets 기록 + Drive HTML 결과 저장

const SHEET_ID   = '1kOMtLCtRVHsGOhXKRmZRX0YbKOWlkx1GUzfhW2WMdgw';   // 배포 후 교체
const FOLDER_IDS = {
  mona:  '1LFvbQOidslPWpgz-Vj31qCZtadvW69AX',
  arefi: '1RKJ_DOiVbG4kT78Zdomhq70HKXi1T4xy',
  citta: '14t-rO20SqzXla3h6Q668pnfph6tDR-Yy'
};

function doPost(e) {
  try {
    var raw  = e.postData.contents || e.postData.getDataAsString();
    var data = JSON.parse(raw);
    var site = data.site;
    var ts   = new Date().toLocaleString('ko-KR', {timeZone:'Asia/Seoul'});

    // ── 1. Google Sheets 기록 ──
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(site) || ss.insertSheet(site);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['제출시각','답사자','답사일시','이동수단','사전예약','협의결과',
                       '필수통과','권장통과','종합판정','종합의견','전체데이터']);
    }
    sheet.appendRow([
      ts,
      data.name    || '',
      data.date    || '',
      data.travel  || '',
      data.prebook || '',
      data.consult || '',
      data.mustPass   !== undefined ? data.mustPass   : '',
      data.recPass    !== undefined ? data.recPass    : '',
      data.verdict || '',
      data.verdictMemo || '',
      JSON.stringify(data)
    ]);

    // ── 2. Drive HTML 결과 저장 ──
    var html    = buildResultHtml(data, ts);
    var fname   = site.toUpperCase() + '_검증결과_' + (data.name||'익명') + '_' + Utilities.formatDate(new Date(),'Asia/Seoul','HHmm') + '.html';
    var blob    = Utilities.newBlob(html, 'text/html; charset=utf-8', fname);
    var folder  = DriveApp.getFolderById(FOLDER_IDS[site]);
    var file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();

    return ContentService
      .createTextOutput(JSON.stringify({status:'ok', url: fileUrl}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function buildResultHtml(d, ts) {
  var siteNames = {mona:'모나밸리 MONAVALLEY', arefi:'아레피 ALEFFEE', citta:'치타누오바 Cittá Nuova'};
  var siteColors = {mona:'#2E5894', arefi:'#6A5D7B', citta:'#004D40'};
  var concepts   = {mona:'품음', arefi:'오름', citta:'열림'};
  var name  = siteNames[d.site]  || d.site;
  var color = siteColors[d.site] || '#2F3E46';
  var kw    = concepts[d.site]   || '';

  var checksHtml = '';
  if (d.checks && Array.isArray(d.checks)) {
    d.checks.forEach(function(c) {
      var icon  = c.state === 'ok' ? '✅' : c.state === 'warn' ? '⚠️' : '○';
      var pTag  = c.priority === 'must' ? '<span style="color:#9B2335;font-size:10px;font-weight:700">[필수]</span>' :
                  c.priority === 'rec'  ? '<span style="color:#2E5894;font-size:10px;font-weight:700">[권장]</span>' :
                                          '<span style="color:#5C6066;font-size:10px">[부가]</span>';
      checksHtml += '<tr><td style="padding:7px 10px;border-bottom:1px solid #F0EDE8;font-size:12px">'
        + icon + ' ' + pTag + ' ' + (c.label||'') + '</td>'
        + '<td style="padding:7px 10px;border-bottom:1px solid #F0EDE8;font-size:11px;color:#5A6068">'
        + (c.hint||'') + '</td></tr>\n';
    });
  }

  return '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>' + name + ' 검증 결과</title>'
    + '<style>'
    + 'body{font-family:"Noto Sans KR",sans-serif;color:#1A1A1A;background:#F2F4F7;margin:0;padding:20px}'
    + '.card{background:white;border-radius:10px;padding:20px;margin-bottom:16px;box-shadow:0 2px 8px rgba(0,0,0,.06)}'
    + '.header-band{background:' + color + ';color:white;border-radius:10px;padding:16px 20px;margin-bottom:16px}'
    + 'table{width:100%;border-collapse:collapse}'
    + 'th{background:#2F3E46;color:white;padding:8px 10px;text-align:left;font-size:12px}'
    + '.verdict-box{border:2px solid ' + color + ';border-radius:8px;padding:12px 16px;text-align:center}'
    + '.verdict-val{font-size:20px;font-weight:700;color:' + color + '}'
    + '</style>'
    + '<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;600;700&display=swap" rel="stylesheet">'
    + '</head><body>'
    + '<div class="header-band">'
    + '<div style="font-size:10px;letter-spacing:2px;opacity:.7;margin-bottom:4px">GJAS 2026 · 사전답사검증표</div>'
    + '<div style="font-size:22px;font-weight:700">' + kw + ' · ' + name + '</div>'
    + '<div style="font-size:11px;opacity:.8;margin-top:4px">제출: ' + ts + '</div>'
    + '</div>'
    + '<div class="card">'
    + '<table><tr><th colspan="2">현장 확인 정보</th></tr>'
    + '<tr><td style="padding:7px 10px;font-size:12px;font-weight:600;border-bottom:1px solid #F0EDE8;width:100px">답사자</td>'
    + '<td style="padding:7px 10px;font-size:13px;border-bottom:1px solid #F0EDE8">' + (d.name||'-') + '</td></tr>'
    + '<tr><td style="padding:7px 10px;font-size:12px;font-weight:600;border-bottom:1px solid #F0EDE8">답사 일시</td>'
    + '<td style="padding:7px 10px;font-size:13px;border-bottom:1px solid #F0EDE8">' + (d.date||'-') + '</td></tr>'
    + '<tr><td style="padding:7px 10px;font-size:12px;font-weight:600">협의 결과</td>'
    + '<td style="padding:7px 10px;font-size:13px">' + (d.consult||'-') + '</td></tr>'
    + '</table></div>'
    + '<div class="card">'
    + '<table><tr><th>체크 항목</th><th style="width:35%">비고</th></tr>'
    + checksHtml + '</table></div>'
    + '<div class="card">'
    + '<div class="verdict-box">'
    + '<div style="font-size:10px;letter-spacing:1px;color:#5A6068;margin-bottom:6px">종합 판정</div>'
    + '<div class="verdict-val">' + (d.verdict === 'ok' ? '✅ 적합' : d.verdict === 'hold' ? '△ 보류' : d.verdict === 'no' ? '✕ 부적합' : '미판정') + '</div>'
    + '</div>'
    + (d.verdictMemo ? '<p style="font-size:13px;margin-top:12px;color:#2F3E46">' + d.verdictMemo + '</p>' : '')
    + '</div>'
    + '<div style="text-align:center;font-size:10px;color:#A0A5AB;margin-top:20px">'
    + 'ARCHIVIA Architects · via@archivia.kr · GJAS_사전답사검증표_v2.0</div>'
    + '</body></html>';
}

function doGet(e) {
  return ContentService.createTextOutput('GJAS GAS OK').setMimeType(ContentService.MimeType.TEXT);
}
