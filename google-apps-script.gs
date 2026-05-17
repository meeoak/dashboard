/**
 * 영업팀 통합 대시보드 — Google Apps Script Webhook
 *
 * 역할:
 *  - 대시보드(static HTML)에서 POST 요청을 받아
 *  - 에이전트별 시트(탭)에 일일 보고서를 누적 기록
 *
 * 시트 구조 (자동 생성):
 *  - 한 스프레드시트 안에 에이전트 이름별로 시트(탭) 자동 생성
 *  - 각 시트의 한 행 = 한 회원의 긴급 관리 항목
 *  - "확인" 컬럼은 체크박스로 자동 변환 (에이전트가 시트에서 직접 체크)
 *  - 같은 회원이 또 올라오면 새 행으로 누적 (시계열 추적 가능)
 *
 * ─────────────────────────────────────────────────────────────
 *  설치 순서 (1회만):
 *   1) script.google.com 접속 → 새 프로젝트
 *   2) 이 파일 전체 내용을 Code.gs 에 붙여넣기
 *   3) 코드 상단의 SPREADSHEET_ID 를 자신의 스프레드시트 ID로 교체
 *      (스프레드시트 URL의 /d/와 /edit 사이 문자열)
 *   4) "배포 > 새 배포" → 유형: 웹앱
 *      - 실행 계정: 본인
 *      - 액세스 권한: 모든 사용자 (대시보드가 정적이라 익명 POST 필요)
 *   5) 배포 후 발급된 URL 복사
 *   6) 대시보드의 에이전트 인사이트 탭 → ⚙️ 버튼 클릭 → URL 붙여넣기
 *
 *  보안 주의:
 *   - URL을 아는 사람만 데이터를 보낼 수 있음 (URL = 비밀 키)
 *   - 외부에 노출되지 않게 관리
 *   - 시트는 별도로 "링크가 있는 모든 사용자 보기 권한" 등으로 공개 설정 가능
 * ─────────────────────────────────────────────────────────────
 */

// 영업팀 통합 대시보드 시트
const SPREADSHEET_ID = '1NykBzZVm5YF1Y6ifluP209jSxVsHA741OYTFCp7DEU4';

// 시트 컬럼 헤더 (순서대로)
const HEADERS = [
  '전송일시',     // A
  '대상일',       // B (대시보드의 D.today)
  '우선순위',     // C (높음/중간/낮음)
  '아이콘',       // D
  '회원명',       // E
  '후원자',       // F
  '배정일',       // G
  '방문일',       // H
  '상태',         // I (s/d/w 등 원본)
  '경과일',       // J
  '잔여(D-day)',  // K (이관 D-day)
  '사유',         // L
  '비고/메모',    // M
  '확인',         // N ← 체크박스 (에이전트가 직접 체크)
  // KPI 요약 (각 보고서 첫 행에만 채워짐)
  '기간',         // O
  'DB',           // P
  '예약',         // Q
  '방문',         // R
  '클로',         // S
  '예약율',       // T
  '방문율',       // U
  '클로율',       // V
  'DB→클로'       // W
];

const SEV_LBL = {3:'🔴 높음', 2:'🟡 중간', 1:'🟢 낮음'};
const ST_LBL  = {s:'방문예약', d:'방문완료', ls:'랑숭', vc:'비디오콜', w:'일정대기', x:'취소', cl:'클로징', none:'미진행'};

function doPost(e){
  try {
    const body = e.postData ? e.postData.contents : '{}';
    const payload = JSON.parse(body);
    const result = writeAgentReport(payload);
    return ContentService
      .createTextOutput(JSON.stringify({ok:true, ...result}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err){
    return ContentService
      .createTextOutput(JSON.stringify({ok:false, error: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// health check + 배포 진단 + 테스트 메일 발송 endpoint
// 브라우저로 직접 열어서 SPREADSHEET_ID / 시트 접근 가능 여부 확인 가능
//
// 테스트 메일 발송:
//   {WEBHOOK_URL}?action=testmail                              → 기본 (devi → meeoak0512@gmail.com)
//   {WEBHOOK_URL}?action=testmail&agent=devi&to=you@gmail.com  → 커스텀
function doGet(e){
  const params = (e && e.parameter) || {};

  // 테스트 메일 발송
  if(params.action === 'testmail'){
    const to = params.to || 'meeoak0512@gmail.com';
    const q  = params.agent || 'devi';
    try {
      const result = sendTestReport_(q, to);
      return ContentService
        .createTextOutput(JSON.stringify({ok:true, action:'testmail', ...result}, null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err){
      return ContentService
        .createTextOutput(JSON.stringify({ok:false, action:'testmail', error: String(err)}, null, 2))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  const info = {
    ok: true,
    msg: 'webhook alive',
    t: new Date().toISOString(),
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetOk: false,
    spreadsheetName: null,
    tabs: null,
    err: null
  };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    info.spreadsheetOk = true;
    info.spreadsheetName = ss.getName();
    info.tabs = ss.getSheets().map(function(s){return s.getName();});
  } catch(err){
    info.err = String(err);
  }
  return ContentService
    .createTextOutput(JSON.stringify(info))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeAgentReport(payload){
  if(!payload || !payload.agent) throw new Error('payload.agent 누락');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = String(payload.agent.name || payload.agent.key || 'UNKNOWN').slice(0, 100);
  let sheet = ss.getSheetByName(sheetName);
  if(!sheet){
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setFrozenRows(1);
    // 확인 컬럼 체크박스 (전체 컬럼에 적용)
    const checkCol = HEADERS.indexOf('확인') + 1;
    if(checkCol > 0){
      const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      sheet.getRange(2, checkCol, sheet.getMaxRows()-1, 1).setDataValidation(rule);
    }
  }
  const now = new Date();
  const tz = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  const nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
  const periodStr = `${(payload.period||{}).from||''} ~ ${(payload.period||{}).to||''}`;
  const items = (payload.urgent && payload.urgent.items) || [];
  const kpi = payload.kpi || {};

  // 헤더 행 (단일행 KPI 요약 + 0건 안내)
  if(items.length === 0){
    const row = [
      nowStr, payload.today||'', '-', 'ℹ️',
      '(긴급 관리 항목 없음)', '', '', '', '', '', '', '', '', false,
      periodStr, kpi.db||0, kpi.sched||0, kpi.visit||0, kpi.cl||0,
      (kpi.r_ds||0)+'%', (kpi.r_sv||0)+'%', (kpi.r_vc||0)+'%', (kpi.r_dc||0)+'%'
    ];
    sheet.appendRow(row);
    return {sheet: sheetName, appended: 1, urgent: 0};
  }
  // 긴급 항목들을 행으로 변환 (첫 행에만 KPI 채움)
  const rows = items.map((it, idx)=> [
    nowStr,
    payload.today||'',
    SEV_LBL[it.severity] || '-',
    it.icon || '',
    it.nm || '',
    it.pa || '',
    it.ad || '',
    it.vd || '',
    ST_LBL[it.st] || it.st || '',
    (it.daysSince ?? '').toString(),
    (it.daysLeft ?? '').toString(),
    it.reason || '',
    it.memo || '',
    false, // 확인 체크박스
    idx===0 ? periodStr : '',
    idx===0 ? (kpi.db||0) : '',
    idx===0 ? (kpi.sched||0) : '',
    idx===0 ? (kpi.visit||0) : '',
    idx===0 ? (kpi.cl||0) : '',
    idx===0 ? ((kpi.r_ds||0)+'%') : '',
    idx===0 ? ((kpi.r_sv||0)+'%') : '',
    idx===0 ? ((kpi.r_vc||0)+'%') : '',
    idx===0 ? ((kpi.r_dc||0)+'%') : ''
  ]);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, HEADERS.length).setValues(rows);
  // 새로 쓴 확인 컬럼에 체크박스 적용
  const checkCol = HEADERS.indexOf('확인') + 1;
  if(checkCol > 0){
    sheet.getRange(startRow, checkCol, rows.length, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }
  // 우선순위 색상 (선택)
  for(let i=0; i<rows.length; i++){
    const sev = items[i].severity;
    const color = sev>=3 ? '#fee2e2' : sev===2 ? '#fef3c7' : sev===1 ? '#dcfce7' : '#f1f5f9';
    sheet.getRange(startRow + i, 3, 1, 2).setBackground(color);
  }
  return {sheet: sheetName, appended: rows.length, urgent: items.length};
}

// ════════════════════════════════════════════════════════════════
//  📧 매일 오전 7시 (WIB) 에이전트별 일일 보고서 자동 이메일
// ════════════════════════════════════════════════════════════════
//
// 설치 순서 (1회만):
//  1) 이 코드 전체를 GAS 프로젝트에 붙여넣기 (저장)
//  2) GAS 콘솔에서 함수 'setupDailyReportTrigger' 선택 → ▶️ 실행
//     · 첫 실행 시 권한 동의 화면 → 모두 허용 (Gmail/Spreadsheet 권한)
//  3) 트리거 메뉴에서 'sendDailyReports' 가 매일 7시로 등록됐는지 확인
//  4) 테스트: 'sendDailyReportsNow' 함수 ▶️ 실행 (오늘 즉시 발송)
//
// 변경 가능한 설정:
//  - REPORT_CFG.TIMEZONE: 'Asia/Jakarta' (WIB) / 'Asia/Seoul' (KST)
//  - REPORT_CFG.HOUR: 발송 시간 (24시간 표기, e.g. 7 = 오전 7시)
//  - REPORT_CFG.EMAIL_SHEET_ID / EMAIL_GID: 에이전트 이메일 매핑 시트
//  - REPORT_CFG.DATA_SHEET_ID: 대시보드 원본 데이터 시트
// ────────────────────────────────────────────────────────────────

const REPORT_CFG = {
  EMAIL_SHEET_ID: '1eTjA_f2nf5xmVLT3G_9la_GTQnJkv87Dp1jKeXdWjeg',
  EMAIL_GID: 1063851809,
  EMAIL_SHEET_NAME: '이메일정보',  // GID 못 찾으면 이름으로 fallback
  DATA_SHEET_ID: '1TMtrsOIN9UylvYolwkyl5EZzUXd1DcDNg6raPPpBep0',
  DATA_GID: 0,
  TIMEZONE: 'Asia/Jakarta',  // WIB · 사용자 기본 인도네시아 에이전트들
  HOUR: 7,                   // 오전 7시 발송
  FROM_NAME: '영업팀 대시보드 · Sales Team Dashboard',
  CC: '',                    // CC 받을 매니저 이메일 (e.g., 'manager@kaia.co.kr'). 비우면 없음.
};

// 트리거 등록 (1회만 실행)
function setupDailyReportTrigger() {
  // 기존 트리거 제거 (중복 방지)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyReports')
    .forEach(t => ScriptApp.deleteTrigger(t));
  // 새 트리거: 매일 REPORT_CFG.HOUR 시
  ScriptApp.newTrigger('sendDailyReports')
    .timeBased()
    .atHour(REPORT_CFG.HOUR)
    .everyDays(1)
    .inTimezone(REPORT_CFG.TIMEZONE)
    .create();
  return '✅ Daily report trigger registered: every day at ' + REPORT_CFG.HOUR + ':00 ' + REPORT_CFG.TIMEZONE;
}

// 트리거 제거 (필요 시)
function removeDailyReportTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyReports')
    .forEach(t => ScriptApp.deleteTrigger(t));
  return '🗑️ Daily report trigger removed';
}

// 즉시 테스트 발송 (트리거 기다리지 않고 지금 실행)
function sendDailyReportsNow() {
  return sendDailyReports();
}

// 메인 — 매일 트리거에 의해 실행됨
function sendDailyReports() {
  const today = Utilities.formatDate(new Date(), REPORT_CFG.TIMEZONE, 'yyyy-MM-dd');
  const monthStart = today.slice(0,7) + '-01';
  Logger.log('📧 Daily report run · today=' + today);

  // 1) 에이전트 이메일 매핑 가져오기
  const agents = getAgentEmails_();
  Logger.log('  · 에이전트 목록: ' + agents.length + '명');
  if(agents.length === 0){
    Logger.log('  ⚠️ 이메일 매핑 시트가 비어있음 — 시트 헤더에 "이름" + "이메일" 컬럼이 있는지 확인');
    return {ok: false, error: 'no agent emails found'};
  }

  // 2) 원본 데이터 시트에서 records 한 번만 로드 (캐시)
  const records = loadDataRecords_();
  Logger.log('  · 데이터 레코드: ' + records.length + '건 로드');

  // 3) 각 에이전트에게 이메일 발송
  let sent = 0, failed = 0;
  agents.forEach(({name, email}) => {
    try {
      const stats = calcAgentStats_(records, name, monthStart, today);
      if(stats.db === 0 && stats.visit === 0 && stats.cl === 0){
        Logger.log('  · ' + name + ' 활동 없음 — 메일 발송 skip');
        return;
      }
      const html = buildReportHTML_({name, stats, today, monthStart});
      const subject = '[일일 보고 · Laporan Harian] ' + name + ' · ' + today;
      const opts = { htmlBody: html, name: REPORT_CFG.FROM_NAME };
      if(REPORT_CFG.CC) opts.cc = REPORT_CFG.CC;
      GmailApp.sendEmail(email, subject, '', opts);
      sent++;
      Logger.log('  ✅ ' + name + ' → ' + email);
    } catch(err){
      failed++;
      Logger.log('  ❌ ' + name + ' 실패: ' + err);
    }
  });
  return {ok: true, today, sent, failed, total: agents.length};
}

// 헬퍼: 이메일 매핑 시트 읽기 (헤더 자동 감지)
function getAgentEmails_() {
  const ss = SpreadsheetApp.openById(REPORT_CFG.EMAIL_SHEET_ID);
  const sheets = ss.getSheets();
  let sheet = sheets.find(s => s.getSheetId() === REPORT_CFG.EMAIL_GID);
  if(!sheet && REPORT_CFG.EMAIL_SHEET_NAME){
    sheet = sheets.find(s => s.getName() === REPORT_CFG.EMAIL_SHEET_NAME);
  }
  if(!sheet) sheet = sheets[0];  // fallback: 첫 번째 탭
  Logger.log('  · 이메일 시트: ' + sheet.getName() + ' (gid=' + sheet.getSheetId() + ')');
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];

  // 헤더에서 이름 / 이메일 컬럼 자동 탐지 (한·영·인니 키워드)
  const headers = data[0].map(h => String(h||'').toLowerCase().trim());
  const nameIdx  = headers.findIndex(h => /이름|name|nama|agen|에이전트|agent/.test(h));
  const emailIdx = headers.findIndex(h => /email|이메일|mail|메일/.test(h));
  if(nameIdx < 0 || emailIdx < 0){
    Logger.log('⚠️ 헤더에서 이름/이메일 컬럼 찾기 실패: [' + headers.join(' | ') + ']');
    return [];
  }
  return data.slice(1)
    .map(row => ({
      name: String(row[nameIdx]||'').trim(),
      email: String(row[emailIdx]||'').trim()
    }))
    .filter(r => r.name && r.email && r.email.indexOf('@') > 0);
}

// 헬퍼: 데이터 시트 → 정규화된 records 배열
function loadDataRecords_() {
  const ss = SpreadsheetApp.openById(REPORT_CFG.DATA_SHEET_ID);
  const sheets = ss.getSheets();
  let sheet = sheets.find(s => s.getSheetId() === REPORT_CFG.DATA_GID) || sheets[0];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  // 헤더는 row 0, 데이터는 row 1+
  // 컬럼 인덱스는 대시보드의 rowToRecord 와 동일하게 가정 (시트 컬럼 순서 안 바뀐다는 전제)
  return data.slice(1).map(row => ({
    ad: parseDate_(row[1]),
    channel: String(row[2]||'').trim(),
    childName: String(row[3]||'').trim(),
    parentName: String(row[4]||'').trim(),
    phone: String(row[5]||'').trim(),
    province: String(row[10]||'').trim(),
    agent: String(row[12]||'').trim(),
    vd: parseDate_(row[14]||''),
    statusA: String(row[17]||'').trim(),
    statusB: String(row[18]||'').trim(),
    clType: String(row[20]||'').trim()
  })).filter(r => r.ad && r.agent);
}

function parseDate_(s) {
  if(!s) return '';
  if(s instanceof Date){
    return Utilities.formatDate(s, REPORT_CFG.TIMEZONE, 'yyyy-MM-dd');
  }
  s = String(s).trim();
  let m = s.match(/^(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if(m) return m[1]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[3]).padStart(2,'0');
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if(m) return m[3]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[1]).padStart(2,'0');
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(m) return m[1]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[3]).padStart(2,'0');
  return '';
}

// 헬퍼: 에이전트별 stats 계산 (이번 달 + 어제 / 오늘)
function calcAgentStats_(records, agentName, monthStart, today) {
  let db = 0, visit = 0, cl = 0;
  let dbToday = 0, visitToday = 0, clToday = 0;
  const yesterday = Utilities.formatDate(
    new Date(new Date(today).getTime() - 86400000),
    REPORT_CFG.TIMEZONE, 'yyyy-MM-dd'
  );
  records.forEach(r => {
    if(r.agent !== agentName) return;
    const inMonth = r.ad >= monthStart && r.ad <= today;
    if(inMonth) db++;
    // 어제 배정 (오전 7시 발송 시점에서 가장 임팩트)
    if(r.ad === yesterday) dbToday++;

    // visit / cl 판단 (status 코드)
    const sB = r.statusB;
    const visited = sB.indexOf('Visit Selesai') >= 0 || sB.indexOf('Langsung') >= 0 || sB.indexOf('Video Call') >= 0;
    const isCl = r.clType && !/^tidak\b/i.test(r.clType);

    if(visited && inMonth){
      visit++;
      if(r.vd === yesterday) visitToday++;
    }
    if(isCl && inMonth){
      cl++;
      if(r.vd === yesterday) clToday++;
    }
  });
  const r_dc = db ? Math.round(cl/db*1000)/10 : 0;
  // 활동지수: 클로징×5 + 방문완료/랑숭/비디오콜 각 1점 (visit 에 모두 포함)
  const score = cl * 5 + visit;
  // 오늘 기준 페이스 (월 일수 대비)
  const td = new Date(today);
  const lastDay = new Date(td.getFullYear(), td.getMonth()+1, 0).getDate();
  const pacePct = Math.round(td.getDate() / lastDay * 100);
  const scorePct = Math.min(100, score);  // KPI 100점 기준
  const paceDiff = scorePct - pacePct;
  return { db, visit, cl, r_dc, dbToday, visitToday, clToday, yesterday, score, scorePct, pacePct, paceDiff, lastDay };
}

// 헬퍼: 보고서 HTML 빌더 (풀버전 · 인니어 · 모닝 보고)
// 이모지는 HTML entity 로 처리하여 4byte surrogate-pair 손상 방지
function buildReportHTML_({name, stats, today, monthStart}) {
  const monthLabel = today.slice(0,7);
  const ff = `font-family:Segoe UI,-apple-system,BlinkMacSystemFont,Roboto,Helvetica,Arial,sans-serif;color:#0f172a`;
  const initial = String(name||'?').charAt(0).toUpperCase();
  const paceColor = stats.paceDiff >= 0 ? '#15803d' : '#dc2626';
  const paceSign  = stats.paceDiff >= 0 ? '+' : '';

  // 활동율 (방문완료/DB)
  const visitRate = stats.db ? Math.round(stats.visit/stats.db*1000)/10 : 0;
  const clRate    = stats.visit ? Math.round(stats.cl/stats.visit*1000)/10 : 0;

  return ''
    + '<div style="' + ff + ';max-width:640px;margin:0 auto;padding:18px;background:#e2e8f0">'

    // ─── 모닝 배너 (응원 메시지) ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#fbbf24 0%,#f59e0b 50%,#ea580c 100%);color:#fff;border-radius:14px;margin-bottom:14px">'
    + '<tr><td style="padding:22px 24px">'
    + '  <div style="font-size:11px;font-weight:800;letter-spacing:3px;opacity:.95">&#x1F305; ' + today + ' &middot; LAPORAN PAGI &middot; MORNING REPORT</div>'
    + '  <div style="font-size:24px;font-weight:900;letter-spacing:-.5px;margin-top:4px">Selamat Pagi, ' + name + '! &#x2600;&#xFE0F;</div>'
    + '  <div style="font-size:13.5px;font-weight:600;opacity:.95;margin-top:6px;line-height:1.55">'
    + '    Setiap pagi adalah kesempatan baru untuk meraih closing!<br>'
    + '    Hari ini adalah hari yang penuh peluang &mdash; <b>semangat dan tetap fokus!</b> &#x1F4AA;'
    + '  </div>'
    + '</td></tr></table>'

    // ─── 헤더 ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#1e3a8a,#1e40af);color:#fff;border-radius:14px;margin-bottom:14px">'
    + '<tr><td style="padding:22px 24px">'
    + '  <div style="font-size:11px;font-weight:700;color:#fbbf24;letter-spacing:2px">SALES TEAM &middot; KAIA</div>'
    + '  <div style="font-size:22px;font-weight:900;letter-spacing:-.5px;margin-top:4px">&#x1F4CB; Laporan Manajemen Member &middot; ' + name + '</div>'
    + '  <div style="font-size:11.5px;font-weight:600;opacity:.85;margin-top:6px">&#x1F4C5; ' + monthStart + ' ~ ' + monthLabel + '-' + stats.lastDay + ' &middot; Performa Bulanan</div>'
    + '</td></tr></table>'

    // ─── 이번달 성과 ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:14px;border:1px solid #e2e8f0;margin-bottom:14px">'
    + '<tr><td style="padding:18px">'
    + '  <div style="font-size:12.5px;color:#64748b;margin:0 0 14px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">&#x1F4CA; ' + monthLabel + ' Performa Bulanan &middot; 이번달 성과</div>'
    + '  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">'
    + statRow_('Alokasi DB &middot; 배정 DB', stats.db, '#1e3a8a')
    + statRow_('Visit Selesai &middot; 방문 완료', stats.visit + ' <span style="font-size:13px;color:#64748b;font-weight:700">(' + visitRate + '%)</span>', '#16a34a')
    + statRow_('Closing &middot; 클로징', stats.cl + ' <span style="font-size:13px;color:#64748b;font-weight:700">(' + clRate + '%)</span>', '#dc2626')
    + statRow_('DB &rarr; Closing &middot; DB&rarr;클로', stats.r_dc + '%', '#7c3aed', true)
    + '  </table>'
    + '</td></tr></table>'

    // ─── 활동지수 (원본 디자인) ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(180deg,#fefce8 0%,#fff 100%);border:1px solid #fde68a;border-radius:14px;margin-bottom:14px">'
    + '<tr><td style="padding:16px 18px">'
    + '  <div style="display:block;margin-bottom:4px"><span style="font-size:15px;font-weight:900;color:#0f172a">&#x1F3C6; Skor Aktivitas &middot; 활동지수</span></div>'
    + '  <div style="font-size:11.5px;color:#64748b;font-weight:700;margin-bottom:14px">Target KPI 100 pts &middot; &#x1F4CD; Pace hari ini ' + stats.pacePct + '%</div>'

    + '  <table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(90deg,#fefce8 0%,#fff 60%);border:1px solid #fde68a;border-radius:10px"><tr>'
    + '    <td style="padding:13px 11px;width:46px;font-size:28px;text-align:center;vertical-align:middle">&#x1F947;</td>'
    + '    <td style="padding:13px 4px;width:46px;vertical-align:middle">'
    + '      <div style="width:38px;height:38px;border-radius:50%;background:#16a34a;color:#fff;text-align:center;line-height:38px;font-size:14px;font-weight:900">' + initial + '</div>'
    + '    </td>'
    + '    <td style="padding:13px 8px;vertical-align:middle">'
    + '      <div style="font-size:14.5px;font-weight:900;color:#0f172a;letter-spacing:-.3px">' + name + '</div>'
    + '      <div style="font-size:10.5px;color:#64748b;font-weight:700;margin-bottom:6px">TIM &middot; Indonesia</div>'
    + '      <div style="position:relative;height:14px;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:7px">'
    + '        <div style="width:' + stats.scorePct + '%;height:100%;background:linear-gradient(90deg,#3b82f6,#1e40af);border-radius:7px"></div>'
    + '        <div style="position:absolute;top:-3px;bottom:-3px;left:' + stats.pacePct + '%;width:3px;background:#0f172a;border-radius:2px"></div>'
    + '      </div>'
    + '      <div style="display:block;font-size:10px;color:#94a3b8;margin-top:3px;font-weight:700">'
    + '        <span>0 pts</span><span style="float:right">100 pts (KPI)</span>'
    + '      </div>'
    + '    </td>'
    + '    <td style="padding:13px 8px;vertical-align:middle;text-align:center;width:140px">'
    + '      <div style="background:linear-gradient(135deg,#bfdbfe,#dbeafe);color:#1e3a8a;font-size:22px;font-weight:900;border-radius:9px;padding:8px 12px;border:1.5px solid #2563eb;letter-spacing:-.5px;line-height:1">'
    + stats.score + '<span style="font-size:11px;font-weight:800;color:#3b82f6;margin-left:2px">pts</span>'
    + '      </div>'
    + '      <div style="font-size:11px;font-weight:800;color:#64748b;background:#f1f5f9;padding:3px 8px;border-radius:6px;border:1px solid #e2e8f0;margin-top:4px">'
    + '        Pace <b style="color:' + paceColor + ';font-weight:900;font-size:12px">' + paceSign + stats.paceDiff + '</b>'
    + '      </div>'
    + '    </td>'
    + '  </tr></table>'

    + '  <div style="margin-top:10px;padding:10px 12px;background:#fffbeb;border-left:3px solid #fbbf24;border-radius:6px;font-size:12px;color:#78350f;line-height:1.55">'
    + '    <b>&#x1F4CD; Diagnosis:</b> Saat ini ' + stats.score + ' poin (' + stats.scorePct + '%). Pace standar hari ini ' + stats.pacePct + '%.'
    + (stats.paceDiff < 0
        ? ' Perlu <b>+' + Math.abs(stats.paceDiff) + ' poin</b> agar tepat waktu &mdash; <b style="color:#dc2626">1 closing = +5 poin</b>!'
        : ' <b style="color:#15803d">Anda sudah di atas pace! Lanjutkan momentum ini!</b>')
    + '  </div>'

    + '</td></tr></table>'

    // ─── 어제 활동 ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:14px;border:1px solid #e2e8f0;margin-bottom:14px">'
    + '<tr><td style="padding:18px">'
    + '  <div style="font-size:12.5px;color:#64748b;margin:0 0 14px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">&#x1F4C5; Aktivitas Kemarin (' + stats.yesterday + ') &middot; 어제 활동</div>'
    + '  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:8px 0">'
    + '  <tr>'
    + '    <td width="33%" style="background:#eff6ff;border-radius:10px;padding:14px;text-align:center">'
    + '      <div style="font-size:30px;font-weight:900;color:#1e3a8a;letter-spacing:-1px;line-height:1">' + stats.dbToday + '</div>'
    + '      <div style="font-size:11px;color:#475569;font-weight:700;margin-top:4px">Alokasi Baru</div>'
    + '      <div style="font-size:10px;color:#94a3b8;font-weight:700">신규 배정</div>'
    + '    </td>'
    + '    <td width="33%" style="background:#f0fdf4;border-radius:10px;padding:14px;text-align:center">'
    + '      <div style="font-size:30px;font-weight:900;color:#16a34a;letter-spacing:-1px;line-height:1">' + stats.visitToday + '</div>'
    + '      <div style="font-size:11px;color:#475569;font-weight:700;margin-top:4px">Visit Selesai</div>'
    + '      <div style="font-size:10px;color:#94a3b8;font-weight:700">방문 완료</div>'
    + '    </td>'
    + '    <td width="33%" style="background:#fef2f2;border-radius:10px;padding:14px;text-align:center">'
    + '      <div style="font-size:30px;font-weight:900;color:#dc2626;letter-spacing:-1px;line-height:1">' + stats.clToday + '</div>'
    + '      <div style="font-size:11px;color:#475569;font-weight:700;margin-top:4px">Closing</div>'
    + '      <div style="font-size:10px;color:#94a3b8;font-weight:700">클로징</div>'
    + '    </td>'
    + '  </tr></table>'
    + '</td></tr></table>'

    // ─── 모닝 동기부여 ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:linear-gradient(135deg,#16a34a,#15803d);color:#fff;border-radius:14px;margin-bottom:14px">'
    + '<tr><td style="padding:22px 24px">'
    + '  <div style="font-size:11px;font-weight:800;letter-spacing:3px;color:#bbf7d0">&#x1F4AA; PESAN MOTIVASI PAGI</div>'
    + '  <div style="font-size:17px;font-weight:900;margin-top:8px;line-height:1.5">' + dailyMessage_(stats) + '</div>'
    + '</td></tr></table>'

    // ─── 푸터 ───
    + '<div style="text-align:center;padding:10px;color:#94a3b8;font-size:10.5px;letter-spacing:.5px">'
    + '  KAIA SALES TEAM DASHBOARD &middot; Selamat berkarya, ' + name + '! &#x1F31F;<br>'
    + '  Laporan otomatis dikirim setiap pagi pukul ' + REPORT_CFG.HOUR + ':00 ' + REPORT_CFG.TIMEZONE
    + '</div>'

    + '</div>';
}

function statRow_(label, value, color, isLast) {
  const border = isLast ? '' : 'border-bottom:1px solid #f1f5f9';
  return ''
    + '<tr><td style="padding:11px 0;' + border + ';font-size:13.5px;font-weight:700;color:#475569">' + label + '</td>'
    + '<td style="padding:11px 0;' + border + ';text-align:right;font-size:24px;font-weight:900;color:' + color + ';letter-spacing:-.6px">' + value + '</td></tr>';
}

function dailyMessage_(stats) {
  if(stats.clToday >= 2) return '&#x1F389; Kemarin closing ' + stats.clToday + ' kali &mdash; performa luar biasa! Pertahankan momentum hari ini!';
  if(stats.clToday === 1) return '&#x1F525; Closing 1 kali kemarin! Tetap fokus &mdash; hari ini bisa lebih baik!';
  if(stats.visitToday >= 3) return '&#x1F680; Visit ' + stats.visitToday + ' kali kemarin &mdash; closing berikutnya sudah dekat! Closing menanti!';
  if(stats.dbToday >= 3) return '&#x1F4CB; ' + stats.dbToday + ' DB baru kemarin &mdash; kontak cepat = 80% closing! Hubungi sekarang!';
  return '&#x2600;&#xFE0F; Hari baru, peluang baru. Hari ini, hubungi minimal 1 member lebih banyak!';
}

// ════════════════════════════════════════════════════════════════
//  🧪 테스트 발송 — 특정 에이전트 보고서를 내 이메일로 한 번 받아보기
// ════════════════════════════════════════════════════════════════
// GAS 콘솔에서 'sendTestToMe_DeviRahayu' 함수 ▶️ 실행
function sendTestToMe_DeviRahayu() {
  return sendTestReport_('devi', 'meeoak0512@gmail.com');
}

// 범용 테스트 — agentNameQuery 는 부분일치 (대소문자 무관)
function sendTestReport_(agentNameQuery, toEmail) {
  const today = Utilities.formatDate(new Date(), REPORT_CFG.TIMEZONE, 'yyyy-MM-dd');
  const monthStart = today.slice(0,7) + '-01';
  Logger.log('🧪 Test report · query="' + agentNameQuery + '" → ' + toEmail);

  const records = loadDataRecords_();
  Logger.log('  · 데이터 레코드: ' + records.length + '건 로드');

  // 부분일치로 에이전트 이름 찾기
  const q = String(agentNameQuery||'').toLowerCase().trim();
  const allAgents = Array.from(new Set(records.map(r => r.agent).filter(Boolean)));
  const matched = allAgents.filter(a => String(a).toLowerCase().indexOf(q) >= 0);

  if(matched.length === 0){
    Logger.log('❌ 에이전트 매칭 실패 — 전체 에이전트 목록:');
    allAgents.sort().forEach(a => Logger.log('   · ' + a));
    return {ok: false, error: 'agent not found', allAgents};
  }
  const agentName = matched[0];
  if(matched.length > 1){
    Logger.log('⚠️ 후보 ' + matched.length + '명 — 첫 번째 사용: ' + agentName);
    Logger.log('   다른 후보: ' + matched.slice(1).join(', '));
  } else {
    Logger.log('✓ 매칭된 에이전트: ' + agentName);
  }

  const stats = calcAgentStats_(records, agentName, monthStart, today);
  Logger.log('  · stats: ' + JSON.stringify(stats));
  const html = buildReportHTML_({name: agentName, stats, today, monthStart});
  const subject = '[TEST · 일일 보고 · Laporan Harian] ' + agentName + ' · ' + today;
  GmailApp.sendEmail(toEmail, subject, '', {
    htmlBody: html,
    name: REPORT_CFG.FROM_NAME
  });
  Logger.log('✅ 테스트 메일 발송 완료 → ' + toEmail);
  return {ok: true, agent: agentName, to: toEmail, today, stats};
}
