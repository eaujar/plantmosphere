/** @OnlyCurrentDoc */

/**
 * plantmosphere. 공동구매 신청 접수
 *
 * 사용법: 신청을 받을 "비공개" 구글 시트에서 확장 프로그램 > Apps Script를 열고
 * 이 코드를 붙여 넣은 뒤, 배포 > 새 배포 > 웹 앱으로 배포하세요.
 * (실행 사용자: 나 / 액세스 권한: 모든 사용자)
 */

const ORDER_SHEET = '신청';
const IMG_BASE = 'https://kbombpark.github.io/plantmosphere/';
const HEADER = ['접수시각', '접수번호', '이름(입금자명)', '전화번호', '이메일', '주소', '규정동의', 'No.', '사진', '학명', '개체', '수량', '공구가', '금액', '상태'];

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
    const rows = items.map(it => {
      const no = Number(it.no) || '';
      const qty = Math.max(1, Math.min(20, Number(it.qty) || 1));
      const price = Math.max(0, Number(it.price) || 0);
      // 사진은 사이트의 img/번호.jpg만 허용 (다른 주소는 넣지 않음)
      const img = /^img\/\d{2}\.jpg$/.test(String(it.img || '')) ? '=IMAGE("' + IMG_BASE + it.img + '")' : '';
      return [now, id, name, tel, email, addr, '동의', no, img, clip(it.name, 80), clip(it.ind, 10), qty,
        price || '미정', price ? price * qty : '미정', '접수'];
    });
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, rows.length, HEADER.length).setValues(rows);
    sheet.setRowHeights(start, rows.length, 90);
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
    sh.setColumnWidth(HEADER.indexOf('사진') + 1, 80);
  }
  return sh;
}

function clip(v, n) {
  return String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, n).trim();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
