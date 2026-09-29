/**
 * 리브러쉬 × 코스트코 캠페인 신청서 → 구글시트 저장
 *
 * 배포 방법
 * 1. 새 구글시트 생성 → 확장 프로그램 > Apps Script
 * 2. 이 코드를 붙여넣고 저장
 * 3. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4. 발급된 웹 앱 URL을 index.html의 GAS_WEB_APP_URL에 붙여넣기
 *
 * 티어(A~G)별로 탭이 자동 생성되어 저장됩니다. (예: "A_10만원")
 */

const HEADERS = [
  '제출일시', '캠페인', '티어', '고료', '이름', '인스타그램', '휴대폰', '이메일',
  '코스트코 회원권', '방문 지점', '요청사항'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tabName = [d.tier || '미지정', d.price].filter(Boolean).join('_');
    let sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange('G:G').setNumberFormat('@'); // 휴대폰 앞자리 0 보존
    }
    sheet.appendRow([
      new Date(), d.campaign, d.tier, d.price, d.name, normalizeInstagram(d.instagram),
      String(d.phone || ''), d.email, d.membership, d.store, d.note
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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
