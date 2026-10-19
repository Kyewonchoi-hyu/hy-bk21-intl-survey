/**
 * 한양대학교 · 5단계 BK21사업 대비 국제공동연구 설문조사 백엔드 (Google Apps Script)
 * (대학원전략팀 · 국제교류팀 · 연구전략기획팀)
 *
 * - 응답은 이 스크립트가 연결된 구글 시트의 '응답' 탭에 한 줄씩 쌓입니다.
 * - 같은 이메일로 다시 제출하면 기존 줄을 덮어씁니다(수정 횟수 +1, 접수번호는 그대로).
 * - 제출하면 접수번호(BK21-001 형식)가 자동으로 부여되고, 교수님 이메일로 안내 메일이 발송됩니다.
 *   메일은 이 스크립트를 배포한 구글 계정의 Gmail에서 나갑니다.
 * - 이 설문은 파일을 받지 않습니다(드라이브 권한 불필요).
 * - 관리자 비밀번호는 코드에 적지 않고 '스크립트 속성'(ADMIN_PASSWORD)에만 둡니다.
 *   → 아래 setPassword() 함수의 비밀번호를 바꾼 뒤 한 번 실행하고, 다시 지워 두세요.
 * - 기존 '독일 BK 사전설문'과 응답이 섞이지 않도록 새 스프레드시트에 따로 설치해 주세요.
 */

const SHEET_NAME  = '응답';
const TZ          = 'Asia/Seoul';
const RECEIPT_PREFIX = 'BK21-';
const MAIL_FROM_NAME = '한양대학교 대학원전략팀';
const SURVEY_TITLE   = '5단계 BK21사업 대비 국제공동연구 설문조사';

const HEADERS = [
  '접수번호', '최초 제출', '최종 수정', '수정 횟수', '안내메일',
  '성명', '직위', '단과대학', '학과', '이메일', '연락처',
  '1.주요 연구 분야', '1.관심 키워드', '1.협력 희망 분야',
  '2.협력 희망 기관', '2.협력 방안', '2.기타', '2.구체적 방안',
  '3.수용 여부', '3.인원', '3.시기', '3.대상', '3.조건·비고',
  '4.파견 여부', '4.파견 주체', '4.인원', '4.시기', '4.대상', '4.조건·비고'
];

/* ---------- 최초 1회 실행 ---------- */
function setPassword() {
  const PASSWORD     = '여기에-관리자-비밀번호';   // ← 바꾸고 실행 → 실행 후 이 줄의 값은 지워 두세요
  const REPLY_TO     = '';   // (선택) 교수님이 안내 메일에 '답장'하면 받을 주소. 예) 'team@hanyang.ac.kr'
  const NOTIFY_EMAIL = '';   // (선택) 제출될 때마다 숨은참조로 사본을 받을 담당자 주소
  if (PASSWORD.indexOf('여기에') === 0) throw new Error('PASSWORD 값을 먼저 바꿔 주세요.');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_PASSWORD', PASSWORD);
  props.setProperty('REPLY_TO', REPLY_TO);
  props.setProperty('NOTIFY_EMAIL', NOTIFY_EMAIL);
  props.deleteProperty('RECEIPT_SEQ');   // 예전 방식(따로 세던 번호)의 흔적 정리. 접수번호는 이제 시트에서 계산합니다.
  sheet_();
  Logger.log('관리자 비밀번호 저장, 시트 준비 완료. 오늘 보낼 수 있는 메일: ' + MailApp.getRemainingDailyQuota() + '통');
}

/** (선택) 안내 메일이 어떻게 보이는지 내 메일로 미리 받아 보기 */
function testMail() {
  const me = Session.getActiveUser().getEmail();
  const sample = { name: '홍길동', position: '교수', college: '공과대학', dept: '기계공학부', email: me,
    q1_main: '예시 연구 분야', q1_keywords: '예시, 키워드', q1_wish: '예시 협력 분야',
    q2_insts: ['아헨공대'], q2_types: ['공동연구 프로젝트'], q3_avail: '가능', q4_avail: '조건부 가능' };
  sendReceiptMail_(sample, RECEIPT_PREFIX + '000', false, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
  Logger.log(me + ' 로 예시 메일을 보냈습니다.');
}

/* ---------- 웹 앱 진입점 ---------- */
function doGet() {
  return json_({ ok: true, config: getConfig_() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  try {
    switch (body.action) {
      case 'submit':
        return json_(submit_(body.data || {}));
      case 'admin_list':
        checkAdmin_(body.password);
        return json_({ ok: true, headers: HEADERS, rows: listRows_(), config: getConfig_() });
      case 'config_set':
        checkAdmin_(body.password);
        setConfig_(body.config || {});
        return json_({ ok: true, config: getConfig_() });
      default:
        return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: err.code || 'server_error', message: String(err.message || err) });
  }
}

/* ---------- 제출 ---------- */
function submit_(d) {
  const email = clean_(d.email).toLowerCase();
  if (!clean_(d.name) || !clean_(d.position) || !clean_(d.college) || !clean_(d.dept) || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ||
      !clean_(d.q1_main) || !clean_(d.q1_keywords) || !clean_(d.q1_wish) || !clean_(d.q3_avail) || !clean_(d.q4_avail)) {
    throw fail_('missing');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = sheet_();
    const col = name => HEADERS.indexOf(name);
    const values = sh.getDataRange().getValues();
    let rowIndex = -1, prev = null;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][col('이메일')]).toLowerCase() === email) { rowIndex = i + 1; prev = values[i]; break; }
    }

    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    const j = v => Array.isArray(v) ? v.map(clean_).join('; ') : clean_(v);
    const receipt = (prev && prev[col('접수번호')]) ? String(prev[col('접수번호')]) : nextReceipt_(values);
    const row = [
      receipt,
      prev ? prev[col('최초 제출')] : now,
      now,
      prev ? Number(prev[col('수정 횟수')] || 0) + 1 : 0,
      '발송 중',
      j(d.name), j(d.position), j(d.college), j(d.dept), email, j(d.phone),
      j(d.q1_main), j(d.q1_keywords), j(d.q1_wish),
      j(d.q2_insts), j(d.q2_types), j(d.q2_other), j(d.q2_detail),
      j(d.q3_avail), j(d.q3_count), j(d.q3_terms), j(d.q3_targets), j(d.q3_note),
      j(d.q4_avail), j(d.q4_units), j(d.q4_count), j(d.q4_terms), j(d.q4_targets), j(d.q4_note)
    ];
    let target = rowIndex;
    if (rowIndex > 0) sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    else { sh.appendRow(row); target = sh.getLastRow(); }
    SpreadsheetApp.flush();   // 잠금을 풀기 전에 시트에 확실히 기록 → 다음 제출이 이 번호를 보고 +1

    // 안내 메일 (실패해도 접수는 유지)
    const updated = rowIndex > 0;
    let mailSent = false, mailNote;
    try {
      if (MailApp.getRemainingDailyQuota() < 1) throw new Error('일일 발송 한도 초과');
      sendReceiptMail_(d, receipt, updated, now);
      mailSent = true; mailNote = '발송 ' + now;
    } catch (err) {
      mailNote = '실패: ' + String(err.message || err).slice(0, 80);
    }
    sh.getRange(target, col('안내메일') + 1).setValue(mailNote);
    return { ok: true, updated: updated, receipt: receipt, mailSent: mailSent };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 접수번호 · 안내 메일 ---------- */
/**
 * 다음 접수번호 = 시트에 남아 있는 접수번호 중 가장 큰 번호 + 1.
 * 번호를 따로 저장해 두지 않고 매번 시트에서 계산하므로,
 * 시트의 응답을 모두 지우면 다음 제출은 다시 001부터 시작합니다.
 * (호출부가 스크립트 잠금을 잡고 있어 동시에 제출해도 번호가 겹치지 않습니다)
 */
function nextReceipt_(values) {
  const c = HEADERS.indexOf('접수번호');
  let max = 0;
  for (let i = 1; i < values.length; i++) {
    const m = String(values[i][c] || '').match(/(\d+)\s*$/);
    if (m && String(values[i][c]).indexOf(RECEIPT_PREFIX) === 0) max = Math.max(max, Number(m[1]));
  }
  const n = max + 1;
  return RECEIPT_PREFIX + ('000' + n).slice(-Math.max(3, String(n).length));
}

function sendReceiptMail_(d, receipt, updated, when) {
  const props = PropertiesService.getScriptProperties();
  const cfg = getConfig_();
  const h = v => String(v === undefined || v === null ? '' : (Array.isArray(v) ? v.join(', ') : v))
    .replace(/^'/, '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const pv = v => String(v === undefined || v === null ? '' : (Array.isArray(v) ? v.join(', ') : v)).replace(/^'/, '');
  const rows = [
    ['접수 일시', when + (updated ? ' (수정 접수)' : '')],
    ['성명', d.name], ['직위', d.position], ['소속', [d.college, d.dept].filter(Boolean).join(' ')],
    ['주요 연구 분야', d.q1_main], ['관심 분야 키워드', d.q1_keywords], ['향후 협력 희망 분야', d.q1_wish],
    ['협력 희망 기관', d.q2_insts], ['협력 방안', d.q2_types],
    ['3. 학생·연구자 수용', d.q3_avail], ['4. 학생·연구자 파견', d.q4_avail]
  ].filter(r => h(r[1]) !== '');
  const contact = cfg.contact || '대학원전략처 대학원전략팀';
  const due = cfg.due || '2026.10.22.(목) 오후 5시';
  const subject = '[한양대학교] ' + SURVEY_TITLE + ' ' + (updated ? '수정 접수' : '접수 완료') + ' (' + receipt + ')';

  const html =
    '<div style="font-family:\'Malgun Gothic\',\'Apple SD Gothic Neo\',sans-serif;max-width:640px;margin:0 auto;color:#18212F;line-height:1.7">' +
      '<div style="background:#0E4A84;color:#fff;padding:22px 24px;border-radius:10px 10px 0 0">' +
        '<div style="font-size:12px;letter-spacing:.08em;opacity:.85">HANYANG UNIVERSITY · 대학원전략팀 · 국제교류팀 · 연구전략기획팀</div>' +
        '<div style="font-size:19px;font-weight:700;margin-top:4px">' + h(SURVEY_TITLE) + '</div>' +
      '</div>' +
      '<div style="border:1px solid #DCE2EA;border-top:0;border-radius:0 0 10px 10px;padding:24px">' +
        '<p style="margin:0 0 14px">' + h(d.name) + ' 교수님, 안녕하십니까.<br>설문이 정상적으로 ' + (updated ? '수정 접수' : '접수') + '되었습니다. 바쁘신 중에 응답해 주셔서 감사합니다.</p>' +
        '<div style="background:#E9F0F8;border-radius:8px;padding:14px 16px;margin:0 0 18px">' +
          '<div style="font-size:12px;color:#5B6678">접수번호</div>' +
          '<div style="font-size:24px;font-weight:800;color:#0E4A84;letter-spacing:.02em">' + h(receipt) + '</div>' +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
          rows.map(r => '<tr><th style="text-align:left;vertical-align:top;padding:8px 10px;border-bottom:1px solid #EEF1F5;color:#5B6678;font-weight:600;width:34%">' + h(r[0]) + '</th>' +
                        '<td style="padding:8px 10px;border-bottom:1px solid #EEF1F5;white-space:pre-wrap">' + h(r[1]) + '</td></tr>').join('') +
        '</table>' +
        '<p style="margin:18px 0 0;font-size:13.5px;color:#5B6678">내용을 고치시려면 ' + h(due) + '까지 설문 페이지에서 <b>같은 이메일</b>로 다시 제출해 주십시오. 접수번호는 그대로 유지됩니다.<br>문의: ' + h(contact) + '</p>' +
      '</div>' +
      '<p style="font-size:12px;color:#8A94A6;text-align:center;margin:14px 0 0">본 메일은 설문 제출 시 자동으로 발송되었습니다.</p>' +
    '</div>';

  const text = d.name + ' 교수님, 설문이 정상적으로 ' + (updated ? '수정 접수' : '접수') + '되었습니다.\n\n' +
    rows.map(r => r[0] + ': ' + pv(r[1])).join('\n') +
    '\n\n수정은 ' + due + '까지 같은 이메일로 다시 제출해 주십시오.\n문의: ' + contact;

  const opt = { to: String(d.email).trim(), subject: subject, htmlBody: html, body: text, name: MAIL_FROM_NAME };
  const replyTo = props.getProperty('REPLY_TO'); if (replyTo) opt.replyTo = replyTo;
  const bcc = props.getProperty('NOTIFY_EMAIL'); if (bcc) opt.bcc = bcc;
  MailApp.sendEmail(opt);
}

/* ---------- 관리자 ---------- */
function checkAdmin_(pw) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) throw fail_('no_password_set');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('admin_fails') || 0);
  if (fails >= 10) throw fail_('locked');
  if (String(pw || '') !== real) {
    cache.put('admin_fails', String(fails + 1), 600);   // 10분간 10회 틀리면 잠금
    Utilities.sleep(700);
    throw fail_('wrong_password');
  }
}

function listRows_() {
  const values = sheet_().getDataRange().getDisplayValues();
  const head = values[0];
  return values.slice(1).filter(r => r.join('') !== '').map(r => {
    const o = {}; head.forEach((h, i) => { o[h] = r[i]; }); return o;
  }).reverse();   // 최근 제출이 위로
}

function getConfig_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('CONFIG') || '{}'); }
  catch (e) { return {}; }
}
function setConfig_(c) {
  const v = { institution: clean_(c.institution).slice(0, 200), due: clean_(c.due).slice(0, 100), contact: clean_(c.contact).slice(0, 200) };
  PropertiesService.getScriptProperties().setProperty('CONFIG', JSON.stringify(v));
}

/* ---------- 도우미 ---------- */
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
  } else if (String(sh.getRange(1, 1).getValues()[0][0]).trim() !== HEADERS[0]) {
    // 제목 행만 지워진 경우: 맨 위에 제목 행을 다시 넣어 응답 줄이 제목으로 읽히지 않게 함
    sh.insertRowBefore(1);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    return sh;
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#E9F0F8');
  return sh;
}
function clean_(v) {
  let s = (v === null || v === undefined) ? '' : String(v).trim().slice(0, 3000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 시트 수식 실행 방지
  return s;
}
function fail_(code) { const e = new Error(code); e.code = code; return e; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
