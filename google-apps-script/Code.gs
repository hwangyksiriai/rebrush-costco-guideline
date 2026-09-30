/**
 * 리브러쉬 × 코스트코 캠페인 신청서 → 구글시트 저장
 *
 * A/B/C 타입(고료) 신청이 모두 '시트1' 한 곳에 쌓이고,
 * '타입'·'고료' 열로 구분됩니다. (필터로 고료별 확인 가능)
 *
 * 배포 방법
 * 1. 구글시트 > 확장 프로그램 > Apps Script
 * 2. 이 코드로 전체 교체 후 저장
 * 3. 배포 > 배포 관리 > 연필(수정) > 버전: 새 버전 > 배포
 *    (새 버전으로 배포하면 웹 앱 URL은 그대로 유지됩니다)
 */

const SHEET_NAME = '시트1';

const HEADERS = [
  '제출일시', '타입', '고료', '이름', '인스타그램', '휴대폰', '이메일',
  '코스트코 회원권', '방문 지점', '방문 예정일', '요청사항', '캠페인'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      d.tier ? d.tier + ' Type' : '',
      d.price || '',
      d.name || '',
      normalizeInstagram(d.instagram),
      String(d.phone || ''),
      d.email || '',
      d.membership || '',
      d.store || '',
      d.visitDate || '',
      d.note || '',
      d.campaign || ''
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// '시트1'이 없으면 첫 번째 시트 사용, 비어 있으면 헤더 자동 생성
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.getRange('F:F').setNumberFormat('@'); // 휴대폰 앞자리 0 보존
  }
  return sheet;
}

// 페이지에서 변환되지 않은 값이 들어와도 시트에는 전체 주소로 저장
function normalizeInstagram(value) {
  let v = String(value || '').trim().replace(/\s+/g, '');
  if (!v) return '';
  const m = v.match(/instagram\.com\/([^/?#]+)/i);
  let handle = m ? m[1] : v;
  handle = handle.replace(/^@+/, '').replace(/[/?#].*$/, '');
  if (!/^[A-Za-z0-9._]{1,30}$/.test(handle)) return v;
  return 'https://www.instagram.com/' + handle;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
