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

const VERSION = 8;
const NOTIFY_TO = 'eaujar.kr@gmail.com'; // 신청 알림을 받을 주소
const ORDER_SHEET = '신청';
const SUMMARY_SHEET = '신청자별';
const CLOSE_AT = new Date('2026-10-15T23:59:59+09:00'); // 모집 마감 (사이트와 같게)
const IMG_BASE = 'https://eaujar.github.io/plantmosphere/';
// 공구가(원): 사이트의 PRICE_KRW와 같아야 합니다. 브라우저가 보낸 금액 대신 이 표로 계산합니다.
const PRICE_KRW = {1:82000,2:82000,3:77000,4:82000,6:82000,7:77000,9:55000,10:82000,11:55000,12:55000,13:55000,14:55000,15:77000,16:82000,17:71000,19:82000,21:82000,22:77000,23:77000,24:82000,25:77000,26:77000,27:55000,28:82000,29:55000,31:66000,32:82000,33:71000,34:55000,36:82000,37:71000,38:71000,39:77000,40:66000,41:82000,42:77000,43:82000,44:82000,45:82000,46:66000,47:55000,48:55000,49:77000,50:71000,51:55000,52:55000,53:55000,54:55000,55:55000,56:55000,57:55000,58:136000};
// 개체 정보: 번호 → [학명, 개체 구분, 사진]. 사이트의 PLANTS와 같아야 합니다. 기록은 브라우저가 보낸 값 대신 이 표를 씁니다.
const PLANTS = {1:["Amorphophallus pendulus","","img/01.jpg"],2:["Ardisia sp","","img/02.jpg"],3:["Ardisia sp blue","","img/03.jpg"],4:["Begonia new species from langkat sumatra","","img/04.jpg"],6:["Codonoboea","","img/06.jpg"],7:["Codonoboea sp","","img/07.jpg"],9:["Davalia sp","","img/09.jpg"],10:["Dendroconche annabelle/microsorum","","img/10.jpg"],11:["Discorea sp","","img/11.jpg"],12:["Diplazium sp","개체 A","img/12.jpg"],13:["Diplazium sp","개체 B","img/13.jpg"],14:["Diplazium sp","개체 C","img/14.jpg"],15:["Elatostemma new sp papua","","img/15.jpg"],16:["Ficus sp","","img/16.jpg"],17:["Picus sp sumatra","","img/17.jpg"],19:["Homalomena aaf hasei","","img/19.jpg"],21:["Homalomena hasei","","img/21.jpg"],22:["Homalomena mini dolpin silver","","img/22.jpg"],23:["Homalomena rugosum","","img/23.jpg"],24:["Homalomena sp hairy","","img/24.jpg"],25:["Homalomena sp sumatra","개체 A","img/25.jpg"],26:["Homalomena sp sumatra","개체 B","img/26.jpg"],27:["Hymenophyllum","","img/27.jpg"],28:["Lindsaea sp blue","","img/28.jpg"],29:["Lindsaea sp sumatra","","img/29.jpg"],31:["Mapania sp","","img/31.jpg"],32:["Mapania sp papua","","img/32.jpg"],33:["Microsorum sp","","img/33.jpg"],34:["Microsorum sp blue","","img/34.jpg"],36:["Pinanga disticha","","img/36.jpg"],37:["Piper sp","","img/37.jpg"],38:["Piper sp borneo","","img/38.jpg"],39:["Piper sp papua","","img/39.jpg"],40:["Piper sp Sulawesi","","img/40.jpg"],41:["Piper sp sumatra","","img/41.jpg"],42:["Piper sp timika","","img/42.jpg"],43:["Sarcopyramis green sp","","img/43.jpg"],44:["Sarcopyramis sp aceh","","img/44.jpg"],45:["Sarcopyramis sp blue","","img/45.jpg"],46:["Sarcopyramis sp kalimantan","","img/46.jpg"],47:["Selaginella sp","","img/47.jpg"],48:["Selaginella sp jambi","","img/48.jpg"],49:["Selliguea sp sumatra","","img/49.jpg"],50:["Sonerila sp borneo","","img/50.jpg"],51:["Sonerila sp dragon","","img/51.jpg"],52:["Sonerila sp java","","img/52.jpg"],53:["Teratophylum sp","","img/53.jpg"],54:["Teratophylum sp borneo","","img/54.jpg"],55:["Fern","","img/55.jpg"],56:["Fern sp","","img/56.jpg"],57:["Fern sp sumatra","","img/57.jpg"],58:["Plant no id","","img/58.jpg"]};
const HEADER = ['접수시각', '접수번호', '이름(입금자명)', '전화번호', '이메일', '주소', '규정동의', 'No.', '사진', '학명', '개체', '수량', '공구가', '금액', '상태'];
const COL = name => HEADER.indexOf(name); // 0부터 시작하는 열 번호

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);

    // 스팸 방지: 숨은 입력칸이 채워져 있으면 조용히 무시
    if (d.hp) return json({ ok: true, id: '-' });

    if (Date.now() > CLOSE_AT.getTime()) return json({ ok: false, error: 'closed' });

    const name = clip(d.name, 40), email = clip(d.email, 80), addr = clip(d.addr, 200);
    const tel = String(d.tel == null ? '' : d.tel).replace(/[^0-9+\-\s()]/g, '').trim().slice(0, 20); // 숫자·기호만
    const items = (Array.isArray(d.items) ? d.items.slice(0, 60) : []).filter(it => PRICE_KRW[Number(it.no)] && PLANTS[Number(it.no)]);
    if (!name || normTel(tel).length < 8 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !addr || d.agree !== true || !items.length) return json({ ok: false, error: 'missing_fields' });

    // 같은 신청이 다시 오면(응답 유실 후 재전송 등) 새로 기록하지 않고 처음 접수번호를 돌려줍니다.
    const key = /^[A-Za-z0-9-]{8,64}$/.test(String(d.key || '')) ? 'req:' + d.key : '';
    const cache = CacheService.getScriptCache();
    if (key) {
      const prev = cache.get(key);
      if (prev) return json(Object.assign(JSON.parse(prev), { dup: true }));
    }

    // 대량 신청 방지: 같은 전화번호 10분에 5건, 전체 10분에 60건까지
    if (!allow('rate:tel:' + normTel(tel), 5) || !allow('rate:all', 60)) return json({ ok: false, error: 'busy' });

    const sheet = getSheet();
    ensureSummary();
    const id = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
    let status = '접수';
    let replaced = [];
    let changed = [];

    if (d.mode === 'edit') {
      const found = findActive(sheet, name, tel);
      if (!found.rows.length) return json({ ok: false, error: found.locked ? 'locked' : 'not_found' });
      found.rows.forEach(r => sheet.getRange(r, COL('상태') + 1).setValue('수정됨 → ' + id));
      replaced = found.ids;
      // 이메일·주소가 바뀌었으면 알림 메일에 표시 (남의 신청을 바꿔치기하는 경우를 알아채기 위해)
      changed = [];
      if (found.emails.length && found.emails.indexOf(email) < 0) changed.push('이메일: ' + found.emails.join(', ') + ' → ' + email);
      if (found.addrs.length && found.addrs.indexOf(addr) < 0) changed.push('주소: ' + found.addrs.join(' / ') + ' → ' + addr);
      status = '접수(수정: ' + replaced.join(', ') + ')';
    }

    const now = new Date();
    const rows = items.map(it => {
      const no = Number(it.no) || '';
      const qty = Math.max(1, Math.min(20, Number(it.qty) || 1));
      const price = PRICE_KRW[no] || 0;
      const p = PLANTS[no]; // 학명·개체·사진은 서버 표 기준
      const img = p[2] ? '=IMAGE("' + IMG_BASE + p[2] + '")' : '';
      return [now, id, name, tel, email, addr, '동의', no, img, p[0], p[1], qty,
        price || '미정', price ? price * qty : '미정', status];
    });
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, COL('전화번호') + 1, rows.length, 1).setNumberFormat('@'); // 010의 0이 사라지지 않게
    sheet.getRange(start, 1, rows.length, HEADER.length).setValues(rows);
    sheet.setRowHeights(start, rows.length, 90);
    const result = { ok: true, id: id, mode: d.mode === 'edit' ? 'edit' : 'new', replaced: replaced };
    if (key) cache.put(key, JSON.stringify(result), 21600); // 6시간 보관
    lock.releaseLock();

    // 새 신청·수정 알림 메일 (메일이 실패해도 신청은 접수된 상태로 둡니다)
    try { notify({ id, name, tel, email, addr, rows, replaced, changed, edit: d.mode === 'edit' }); } catch (mailErr) { console.error('알림 메일 실패: ' + mailErr); }

    return json(result);
  } catch (err) {
    return json({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}

// 이름·전화번호가 같은 신청 중 아직 수정할 수 있는 줄을 찾습니다.
function findActive(sheet, name, tel) {
  const last = sheet.getLastRow();
  const res = { rows: [], ids: [], locked: false, emails: [], addrs: [] };
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
      const pe = String(v[COL('이메일')]), pa = String(v[COL('주소')]);
      if (res.emails.indexOf(pe) < 0) res.emails.push(pe);
      if (res.addrs.indexOf(pa) < 0) res.addrs.push(pa);
    } else if (st.indexOf('수정됨') !== 0) {
      res.locked = true; // 확정 등 운영자가 상태를 바꾼 신청
    }
  });
  return res;
}

// 운영자(스크립트 소유자)에게 신청 요약 메일을 보냅니다.
function notify(o) {
  const to = NOTIFY_TO;
  if (MailApp.getRemainingDailyQuota() < 3) { console.error('메일 한도 부족으로 알림 생략: ' + o.id); return; }
  const won = n => Number(n).toLocaleString('ko-KR') + '원';
  const total = o.rows.reduce((a, r) => a + (Number(r[COL('금액')]) || 0), 0);
  const count = o.rows.reduce((a, r) => a + Number(r[COL('수량')]), 0);
  const lines = o.rows.map(r => {
    const no = String(r[COL('No.')]).padStart(2, '0');
    const ind = r[COL('개체')] ? ' (' + r[COL('개체')] + ')' : '';
    return 'No.' + no + ' ' + r[COL('학명')] + ind + ' × ' + r[COL('수량')] + ' = ' + won(r[COL('금액')]);
  });
  let tag = o.edit ? '[수정] ' : '';
  if (o.changed && o.changed.length) tag += '[연락처 변경] ';
  const subject = '[plantmosphere] ' + tag + (o.edit ? '수정 신청 ' : '새 신청 ') + o.id + ' · ' + o.name + ' · ' + won(total);
  const body = [
    (o.edit ? '신청 내용이 수정되었습니다.' : '새 공동구매 신청이 들어왔습니다.'),
    o.edit && o.replaced.length ? '대체된 이전 신청: ' + o.replaced.join(', ') : '',
    o.changed && o.changed.length ? '⚠ 연락처 정보가 바뀌었습니다. 본인이 수정한 것인지 확인해 주세요.\n' + o.changed.join('\n') : '',
    '',
    '접수번호: ' + o.id,
    '이름(입금자명): ' + o.name,
    '전화번호: ' + o.tel,
    '이메일: ' + o.email,
    '주소: ' + o.addr,
    '',
    '[신청 개체] ' + o.rows.length + '종 ' + count + '개체',
  ].concat(lines, [
    '',
    '최종 입금액: ' + won(total),
    '',
    '신청 시트: ' + SpreadsheetApp.getActive().getUrl(),
  ]).filter((l, i, a) => !(l === '' && a[i - 1] === '')).join('\n');
  MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'plantmosphere. 공동구매' });
}

// 편집기에서 이 함수를 선택해 ▶ 실행하면 알림 메일이 오는지 바로 확인할 수 있습니다.
function testMail() {
  getSheet(); ensureSummary(); // '신청'·'신청자별' 탭도 미리 만들어 둡니다
  MailApp.sendEmail({ to: NOTIFY_TO, subject: '[plantmosphere] 알림 메일 테스트',
    body: '이 메일이 보이면 신청 알림이 정상적으로 발송됩니다.\n남은 하루 발송 가능 수: ' + MailApp.getRemainingDailyQuota(),
    name: 'plantmosphere. 공동구매' });
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

// '신청자별' 탭: '신청' 탭을 신청 건별로 묶어 종 수·개체 수·합계 금액을 보여줍니다. ('수정됨' 줄 제외)
function ensureSummary() {
  const ss = SpreadsheetApp.getActive();
  const formula = "=QUERY('" + ORDER_SHEET + "'!A:O, \"select min(A), B, C, D, E, F, count(H), sum(L), sum(N), max(O) " +
    "where H is not null and not O starts with '수정됨' group by B, C, D, E, F order by min(A) " +
    "label min(A) '접수시각', B '접수번호', C '이름(입금자명)', D '전화번호', E '이메일', F '주소', " +
    "count(H) '종 수', sum(L) '개체 수', sum(N) '합계 금액', max(O) '상태'\", 1)";
  let sh = ss.getSheetByName(SUMMARY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SUMMARY_SHEET);
    sh.setFrozenRows(1);
    sh.getRange('A1:J1').setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.getRange('I:I').setNumberFormat('#,##0"원"');
    sh.setColumnWidth(6, 260);
  }
  // '신청' 탭을 지웠다가 다시 만들면 수식이 깨지므로(#REF) 항상 원래 수식으로 되돌립니다.
  if (sh.getRange('A1').getFormula() !== formula) sh.getRange('A1').setFormula(formula);
}

// 10분 단위로 횟수를 세어 limit을 넘으면 false
function allow(name, limit) {
  const c = CacheService.getScriptCache();
  const k = name + ':' + Math.floor(Date.now() / 600000);
  const n = Number(c.get(k) || 0) + 1;
  c.put(k, String(n), 700);
  return n <= limit;
}

// 시트 수식 주입 방지: 앞뒤 공백·제어문자를 먼저 정리한 뒤, =+-@ 로 시작하면 글자로 저장되게 ' 를 붙입니다.
function clip(v, n) {
  const s = String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, n);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
