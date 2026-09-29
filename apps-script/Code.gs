/** @OnlyCurrentDoc */

/**
 * plantmosphere. 공동구매 신청 접수
 *
 * 사용법: 신청을 받을 "비공개" 구글 시트에서 확장 프로그램 > Apps Script를 열고
 * 이 코드를 붙여 넣은 뒤, 배포 > 새 배포 > 웹 앱으로 배포하세요.
 * (실행 사용자: 나 / 액세스 권한: 모든 사용자)
 *
 * 신청 수정: 페이지의 '신청 내용 수정'으로 다시 신청하면 이름·전화번호가 같은
 * 이전 신청(상태가 '접수' 또는 '접수(수정 …)')을 '수정됨'으로 바꾸고 새 내용을 기록합니다.
 * 운영자가 상태를 '확정' 등 다른 값으로 바꾼 신청은 페이지에서 수정할 수 없습니다.
 */

const VERSION = 2;
const ORDER_SHEET = '신청';
const IMG_BASE = 'https://kbombpark.github.io/plantmosphere/';
const HEADER = ['접수시각', '접수번호', '이름(입금자명)', '전화번호', '이메일', '주소', '규정동의', 'No.', '사진', '학명', '개체', '수량', '공구가', '금액', '상태'];
const COL = name => HEADER.indexOf(name); // 0부터 시작하는 열 번호

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
    let status = '접수';
    let replaced = [];

    if (d.mode === 'edit') {
      const found = findActive(sheet, name, tel);
      if (!found.rows.length) return json({ ok: false, error: found.locked ? 'locked' : 'not_found' });
      found.rows.forEach(r => sheet.getRange(r, COL('상태') + 1).setValue('수정됨 → ' + id));
      replaced = found.ids;
      status = '접수(수정: ' + replaced.join(', ') + ')';
    }

    const now = new Date();
    const rows = items.map(it => {
      const no = Number(it.no) || '';
      const qty = Math.max(1, Math.min(20, Number(it.qty) || 1));
      const price = Math.max(0, Number(it.price) || 0);
      // 사진은 사이트의 img/번호.jpg만 허용 (다른 주소는 넣지 않음)
      const img = /^img\/\d{2}\.jpg$/.test(String(it.img || '')) ? '=IMAGE("' + IMG_BASE + it.img + '")' : '';
      return [now, id, name, tel, email, addr, '동의', no, img, clip(it.name, 80), clip(it.ind, 10), qty,
        price || '미정', price ? price * qty : '미정', status];
    });
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, COL('전화번호') + 1, rows.length, 1).setNumberFormat('@'); // 010의 0이 사라지지 않게
    sheet.getRange(start, 1, rows.length, HEADER.length).setValues(rows);
    sheet.setRowHeights(start, rows.length, 90);
    return json({ ok: true, id: id, mode: d.mode === 'edit' ? 'edit' : 'new', replaced: replaced });
  } catch (err) {
    return json({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}

// 이름·전화번호가 같은 신청 중 아직 수정할 수 있는 줄을 찾습니다.
function findActive(sheet, name, tel) {
  const last = sheet.getLastRow();
  const res = { rows: [], ids: [], locked: false };
  if (last < 2) return res;
  const values = sheet.getRange(2, 1, last - 1, HEADER.length).getValues();
  const wantName = normName(name), wantTel = normTel(tel);
  values.forEach((v, i) => {
    if (normName(v[COL('이름(입금자명)')]) !== wantName || normTel(v[COL('전화번호')]) !== wantTel) return;
    const st = String(v[COL('상태')] || '');
    if (st === '접수' || st.indexOf('접수(수정') === 0) {
      res.rows.push(i + 2);
      const oid = String(v[COL('접수번호')]);
      if (res.ids.indexOf(oid) < 0) res.ids.push(oid);
    } else if (st.indexOf('수정됨') !== 0) {
      res.locked = true; // 확정 등 운영자가 상태를 바꾼 신청
    }
  });
  return res;
}

function normName(v) { return String(v || '').replace(/^'/, '').replace(/\s+/g, '').toLowerCase(); }
function normTel(v) { return String(v || '').replace(/\D/g, '').replace(/^0+/, ''); }

// 브라우저로 웹 앱 주소를 열었을 때 동작 확인용 (페이지가 수정 기능 지원 여부를 확인할 때도 사용)
function doGet() {
  return json({ ok: true, service: 'plantmosphere-groupbuy', version: VERSION });
}

function getSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ORDER_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ORDER_SHEET);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
    sh.setColumnWidth(COL('사진') + 1, 80);
  }
  return sh;
}

function clip(v, n) {
  return String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, n).trim();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
