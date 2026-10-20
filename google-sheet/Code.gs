/**
 * 업장 방문일지 → 구글 시트 자동 동기화
 *
 * 1) 아래 KEY 의 '여기에-비밀번호' 를 원하는 비밀번호로 바꾸세요. (앱에도 같은 비밀번호를 넣어요)
 * 2) 배포 → 새 배포 → 유형: 웹 앱 / 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포
 * 3) 나온 "웹 앱 URL" 을 방문일지 앱의 "구글 시트 연결" 칸에 붙여 넣으세요.
 * 명함 글자 읽기: 왼쪽 "서비스 +" 에서 "Drive API" 를 추가해야 해요.
 */
const KEY = '여기에-비밀번호';

const SHEET_NAME = '방문기록';
const FOLDER_NAME = '방문일지 사진';
const HEADERS = ['방문일시', '업장명', '주소', '메모', '위도', '경도', 'GPS 오차(m)', '지도', '사진1', '사진2', '사진3', '사진4', '기록 ID', '받은 시각', '대표자', '이메일', '전화번호', '명함 사진'];
const ID_COL = 13; // '기록 ID' 열 번호

function doGet() {
  return json_({ ok: true, app: 'visit-log', message: '방문일지 시트가 연결을 기다리고 있어요.' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (KEY === '여기에-비밀번호') return json_({ ok: false, error: 'key_not_set', message: '스크립트의 KEY 비밀번호를 먼저 바꿔 주세요.' });
    if (body.key !== KEY) return json_({ ok: false, error: 'bad_key', message: '비밀번호가 맞지 않아요.' });
    if (body.ping) return json_({ ok: true, message: '연결됐어요.' });
    if (body.ocr) return ocr_(body.ocr);

    const v = body.visit || {};
    if (!v.id || !v.name) return json_({ ok: false, error: 'bad_request', message: '기록 내용이 비어 있어요.' });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sheet = sheet_();
      // 같은 기록이 이미 있으면 다시 넣지 않음 (다시 보내도 안전)
      const last = sheet.getLastRow();
      if (last > 1) {
        const ids = sheet.getRange(2, ID_COL, last - 1, 1).getValues().flat().map(String);
        if (ids.indexOf(String(v.id)) !== -1) return json_({ ok: true, duplicate: true });
      }

      const folder = folder_();
      const links = (body.photos || []).slice(0, 4).map(function (dataUrl, i) {
        const m = /^data:([^;]+);base64,(.*)$/.exec(dataUrl);
        if (!m) return '';
        const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], safeName_(v) + '_' + (i + 1) + '.jpg');
        return folder.createFile(blob).getUrl();
      });
      while (links.length < 4) links.push('');

      let cardLink = '';
      const cm = body.card && /^data:([^;]+);base64,(.*)$/.exec(body.card);
      if (cm) cardLink = folder.createFile(Utilities.newBlob(Utilities.base64Decode(cm[2]), cm[1], safeName_(v) + '_명함.jpg')).getUrl();

      const hasLoc = v.lat != null && v.lng != null;
      const row = [
        v.visitedAt ? new Date(v.visitedAt) : '',
        v.name, v.address || '', v.memo || '',
        hasLoc ? v.lat : '', hasLoc ? v.lng : '', v.locAcc || '',
        hasLoc ? 'https://map.kakao.com/link/map/' + encodeURIComponent(v.name) + ',' + v.lat + ',' + v.lng : '',
      ].concat(links, [String(v.id), new Date(), v.person || '', v.email || '', v.phone || '', cardLink]);
      sheet.appendRow(row);
      sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm');
      return json_({ ok: true });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err) });
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  // 예전 버전으로 만든 시트에는 새 칸 제목을 덧붙임
  if (sh.getLastColumn() < HEADERS.length) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  return sh;
}

/** 명함 사진 글자 읽기: 구글 문서로 변환(OCR)한 뒤 글자만 꺼내고 임시 문서는 지움 */
function ocr_(dataUrl) {
  if (typeof Drive === 'undefined') return json_({ ok: false, error: 'ocr_setup', message: 'Apps Script 왼쪽 "서비스 +"에서 Drive API를 추가해 주세요.' });
  const m = /^data:([^;]+);base64,(.*)$/.exec(dataUrl);
  if (!m) return json_({ ok: false, error: 'bad_request', message: '명함 사진을 읽을 수 없어요.' });
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'card.jpg');
  const file = Drive.Files.create({ name: '명함 글자 읽기(임시)', mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'ko' });
  try {
    return json_({ ok: true, text: DocumentApp.openById(file.id).getBody().getText() });
  } finally {
    DriveApp.getFileById(file.id).setTrashed(true);
  }
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function safeName_(v) {
  const d = v.visitedAt ? Utilities.formatDate(new Date(v.visitedAt), 'Asia/Seoul', 'yyyyMMdd_HHmm') : 'visit';
  return d + '_' + String(v.name).replace(/[\\/:*?"<>|]/g, '').slice(0, 30);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 처음 한 번 실행해서 권한을 허락할 때 쓰는 함수 (위쪽 실행 버튼으로 실행) */
function setup() {
  sheet_();
  folder_();
  DocumentApp.getActiveDocument; // 명함 읽기에 필요한 문서 권한을 함께 받기 위한 줄
  if (typeof Drive !== 'undefined') Drive.Files.list({ pageSize: 1 });
}
