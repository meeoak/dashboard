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

// health check + 배포 진단
// 브라우저로 직접 열어서 SPREADSHEET_ID / 시트 접근 가능 여부 확인 가능
function doGet(){
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
  return { db, visit, cl, r_dc, dbToday, visitToday, clToday, yesterday };
}

// 헬퍼: 보고서 HTML 빌더
function buildReportHTML_({name, stats, today, monthStart}) {
  const monthLabel = today.slice(0,7);
  const css = `font-family:'Segoe UI',-apple-system,sans-serif;color:#0f172a`;
  return ''
    + '<div style="' + css + ';max-width:600px;margin:0 auto;padding:24px;background:#f8fafc">'
    + '  <div style="background:linear-gradient(135deg,#1e3a8a,#1e40af);color:#fff;padding:20px 24px;border-radius:14px 14px 0 0">'
    + '    <div style="font-size:11px;font-weight:700;color:#fbbf24;letter-spacing:3px;margin-bottom:4px">' + today + ' · MORNING REPORT</div>'
    + '    <div style="font-size:22px;font-weight:900;letter-spacing:-.5px">🌅 ' + name + ' 일일 보고</div>'
    + '    <div style="font-size:12px;font-weight:600;opacity:.85;margin-top:4px">Selamat pagi · 좋은 아침 — 오늘도 화이팅!</div>'
    + '  </div>'
    + '  <div style="background:#fff;padding:22px 24px;border-radius:0 0 14px 14px;border:1px solid #e2e8f0;border-top:none">'
    + '    <h2 style="font-size:12.5px;color:#64748b;margin:0 0 14px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">📊 ' + monthLabel + ' 이번달 성과 · Performa Bulanan</h2>'
    + '    <table style="width:100%;border-collapse:collapse">'
    + statRow_('배정 DB · Alokasi DB', stats.db, '#1e3a8a')
    + statRow_('방문 완료 · Visit Selesai', stats.visit, '#16a34a')
    + statRow_('클로징 · Closing', stats.cl, '#dc2626')
    + statRow_('DB → 클로징 · DB→Close', stats.r_dc + '%', '#7c3aed', true)
    + '    </table>'
    + '    <h2 style="font-size:12.5px;color:#64748b;margin:20px 0 12px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">📅 어제 (' + stats.yesterday + ') 활동 · Aktivitas Kemarin</h2>'
    + '    <div style="display:flex;gap:10px;flex-wrap:wrap">'
    + '      <div style="flex:1;min-width:120px;padding:14px;background:#eff6ff;border-radius:9px;text-align:center">'
    + '        <div style="font-size:28px;font-weight:900;color:#1e3a8a;letter-spacing:-1px">' + stats.dbToday + '</div>'
    + '        <div style="font-size:11px;color:#475569;font-weight:700">신규 배정</div>'
    + '      </div>'
    + '      <div style="flex:1;min-width:120px;padding:14px;background:#f0fdf4;border-radius:9px;text-align:center">'
    + '        <div style="font-size:28px;font-weight:900;color:#16a34a;letter-spacing:-1px">' + stats.visitToday + '</div>'
    + '        <div style="font-size:11px;color:#475569;font-weight:700">방문 완료</div>'
    + '      </div>'
    + '      <div style="flex:1;min-width:120px;padding:14px;background:#fef2f2;border-radius:9px;text-align:center">'
    + '        <div style="font-size:28px;font-weight:900;color:#dc2626;letter-spacing:-1px">' + stats.clToday + '</div>'
    + '        <div style="font-size:11px;color:#475569;font-weight:700">클로징</div>'
    + '      </div>'
    + '    </div>'
    + '    <div style="margin-top:22px;padding:14px;background:linear-gradient(135deg,#fef3c7,#fefce8);border-radius:9px;border-left:5px solid #f59e0b">'
    + '      <div style="font-size:12px;font-weight:900;color:#78350f;margin-bottom:4px">💪 오늘의 한 줄</div>'
    + '      <div style="font-size:13px;color:#92400e;font-weight:600;line-height:1.5">' + dailyMessage_(stats) + '</div>'
    + '    </div>'
    + '    <p style="margin:18px 0 0;color:#94a3b8;font-size:10.5px;text-align:center;letter-spacing:.5px">SALES TEAM DASHBOARD · Auto-sent at ' + REPORT_CFG.HOUR + ':00 ' + REPORT_CFG.TIMEZONE + '</p>'
    + '  </div>'
    + '</div>';
}

function statRow_(label, value, color, isLast) {
  const border = isLast ? '' : 'border-bottom:1px solid #f1f5f9';
  return ''
    + '<tr><td style="padding:11px 0;' + border + ';font-size:13.5px;font-weight:700;color:#475569">' + label + '</td>'
    + '<td style="padding:11px 0;' + border + ';text-align:right;font-size:24px;font-weight:900;color:' + color + ';letter-spacing:-.6px">' + value + '</td></tr>';
}

function dailyMessage_(stats) {
  if(stats.clToday >= 2) return '🎉 어제 ' + stats.clToday + '건 클로징 — 멋진 성과입니다! Pertahankan momentum hari ini!';
  if(stats.clToday === 1) return '🔥 어제 1건 클로징 성공! Tetap fokus — hari ini bisa lebih baik!';
  if(stats.visitToday >= 3) return '🚀 어제 방문 ' + stats.visitToday + '건 — 다음 클로징이 멀지 않았어요! Closing menanti!';
  if(stats.dbToday >= 3) return '📋 신규 DB ' + stats.dbToday + '건 배정 — 빠른 첫 컨택이 클로징의 80%! Hubungi segera!';
  return '☀️ 새로운 하루입니다. 오늘은 1건이라도 더 컨택해보세요! Hari ini, hubungi member lebih banyak!';
}
