/** @OnlyCurrentDoc */

/**
 * plantmosphere. 공동구매 신청 접수
 *
 * 사용법: 신청을 받을 "비공개" 구글 시트에서 확장 프로그램 > Apps Script를 열고
 * 이 코드를 붙여 넣은 뒤, 배포 > 새 배포 > 웹 앱으로 배포하세요.
 * (실행 사용자: 나 / 액세스 권한: 모든 사용자)
 */

const ORDER_SHEET = '신청';
const HEADER = ['접수시각', '접수번호', '이름(입금자명)', '전화번호', '이메일', '주소', '규정동의', 'No.', '학명', '개체', '수량', '상태'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);

    // 스팸 방지: 숨은 입력칸이 채워져 있으면 조용히 무시
    if (d.hp) return json({ ok: true, id: '-' });

    const name = clip(d.name, 40), tel = clip(d.tel, 30), email = clip(d.email, 80), addr = clip(d.addr, 200);
    const items = Array.isArray(d.items) ? d.items.slice(0, 60) : [];
    if (!name || !tel || !email || !addr || d.agree !== true || !items.length) return json({ ok: false, error: 'missing_fields' });

    const sheet = getSheet();
    const id = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
    const now = new Date();
    const rows = items.map(it => [
      now, id, name, tel, email, addr, '동의',
      Number(it.no) || '', clip(it.name, 80), clip(it.ind, 10), Math.max(1, Math.min(20, Number(it.qty) || 1)), '접수'
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADER.length).setValues(rows);
    return json({ ok: true, id: id });
  } catch (err) {
    return json({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}

// 브라우저로 웹 앱 주소를 열었을 때 동작 확인용
function doGet() {
  return json({ ok: true, service: 'plantmosphere-groupbuy' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ORDER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ORDER_SHEET);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  }
  return sh;
}

function clip(v, n) {
  return String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, n).trim();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
