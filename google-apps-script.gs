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
      const urgent = getUrgentMembers_(records, name, today);
      const cal    = getCalendarSummary_(records, name, monthStart, today);
      const html = buildReportHTML_({name, stats, urgent, cal, records, today, monthStart});
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
  const r_dv = db ? Math.round(visit/db*1000)/10 : 0;
  const r_vc = visit ? Math.round(cl/visit*1000)/10 : 0;
  // 활동지수: 클로징×5 + 방문완료/랑숭/비디오콜 각 1점 (visit 에 모두 포함)
  const score = cl * 5 + visit;
  // 오늘 기준 페이스 (월 일수 대비)
  const td = new Date(today);
  const lastDay = new Date(td.getFullYear(), td.getMonth()+1, 0).getDate();
  const pacePct = Math.round(td.getDate() / lastDay * 100);
  const scorePct = Math.min(100, score);
  const paceDiff = scorePct - pacePct;

  // DB 활동 분포 (DB 배정 분석용)
  const myInMonth = records.filter(r => r.agent === agentName && r.ad >= monthStart && r.ad <= today);
  let progressing = 0, waiting = 0, idle = 0, etc = 0;
  myInMonth.forEach(r => {
    const sB = (r.statusB||'').toLowerCase();
    const sA = (r.statusA||'').toLowerCase();
    if(sB.indexOf('visit selesai') >= 0 || sB.indexOf('langsung') >= 0 || sB.indexOf('video call') >= 0 || sB.indexOf('booking') >= 0 || sA.indexOf('daftar') >= 0) progressing++;
    else if(sB.indexOf('wait') >= 0 || sB.indexOf('pending') >= 0 || sB.indexOf('tunggu') >= 0) waiting++;
    else if(!sB || sB === 'none' || sB.indexOf('belum') >= 0 || sB.indexOf('tidak') >= 0) idle++;
    else etc++;
  });
  const actRate = db ? Math.round(progressing / db * 1000)/10 : 0;

  return {
    db, visit, cl, r_dc, r_dv, r_vc,
    dbToday, visitToday, clToday, yesterday,
    score, scorePct, pacePct, paceDiff, lastDay,
    progressing, waiting, idle, etc, actRate
  };
}

// 긴급 회원 분류 (4 카테고리)
function getUrgentMembers_(records, agentName, today) {
  const td = new Date(today);
  const my = records.filter(r => r.agent === agentName);
  const simulasi = [], belum = [], dorman = [], menunggu = [];

  my.forEach(r => {
    const sB = (r.statusB||'').toLowerCase();
    const sA = (r.statusA||'').toLowerCase();
    const visited = sB.indexOf('visit selesai') >= 0 || sB.indexOf('langsung') >= 0 || sB.indexOf('video call') >= 0;
    const isCl    = r.clType && !/^tidak\b/i.test(r.clType);
    const isDaftar = sA.indexOf('daftar') >= 0;
    const daysAd = r.ad ? Math.floor((td - new Date(r.ad))/86400000) : 0;
    const daysVd = r.vd ? Math.floor((td - new Date(r.vd))/86400000) : 0;
    const memberName = r.childName || r.parentName || '(no name)';

    if(isCl && !isDaftar && daysVd >= 3 && daysVd <= 30){
      menunggu.push({name: memberName, days: daysVd, reason: 'Closing setuju &middot; menunggu pendaftaran'});
    } else if(visited && !isCl && daysVd >= 7){
      simulasi.push({name: memberName, days: daysVd, reason: 'Visit selesai &middot; perlu follow-up closing'});
    } else if(!visited && daysAd >= 30){
      dorman.push({name: memberName, days: daysAd, reason: 'Tidak ada aktivitas 30+ hari'});
    } else if(!visited && daysAd >= 3 && daysAd <= 14){
      belum.push({name: memberName, days: daysAd, reason: 'DB baru &middot; kontak pertama menanti'});
    }
  });

  // 각 카테고리 D-day 큰 순으로 정렬
  [simulasi, belum, dorman, menunggu].forEach(arr => arr.sort((a,b)=>b.days-a.days));
  return { simulasi, belum, dorman, menunggu };
}

// 방문 캘린더 카운트 (월간)
function getCalendarSummary_(records, agentName, monthStart, today) {
  const my = records.filter(r => r.agent === agentName && r.vd && r.vd >= monthStart.slice(0,7)+'-01' && r.vd <= monthStart.slice(0,7)+'-31');
  let visit = 0, meeting = 0, visual = 0, jadwal = 0;
  my.forEach(r => {
    const sB = (r.statusB||'').toLowerCase();
    if(sB.indexOf('visit selesai') >= 0) visit++;
    if(sB.indexOf('video call') >= 0 || sB.indexOf('meeting') >= 0) meeting++;
    if(sB.indexOf('langsung') >= 0 || sB.indexOf('visual') >= 0) visual++;
    if(r.vd) jadwal++;
  });
  return { visit, meeting, visual, jadwal };
}

// 헬퍼: 보고서 HTML 빌더 (풀버전 · 인니어 · 모닝 보고)
// 이모지는 HTML entity 로 처리하여 4byte surrogate-pair 손상 방지
function buildReportHTML_({name, stats, urgent, cal, records, today, monthStart}) {
  urgent  = urgent  || {simulasi:[], belum:[], dorman:[], menunggu:[]};
  cal     = cal     || {visit:0, meeting:0, visual:0, jadwal:0};
  records = records || [];
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

    // ─── DB 배정 분석 (3 카드) ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:14px;border:1px solid #e2e8f0;margin-bottom:14px">'
    + '<tr><td style="padding:18px">'
    + '  <div style="font-size:12.5px;color:#64748b;margin:0 0 14px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">&#x1F4CA; Analisis Alokasi DB &middot; DB 배정 분석</div>'
    + '  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:8px 0"><tr>'
    + '    <td width="35%" style="background:linear-gradient(135deg,#1e293b,#0f172a);color:#fff;padding:16px;border-radius:11px;text-align:center;vertical-align:middle">'
    + '      <div style="font-size:10px;color:#fbbf24;font-weight:800;letter-spacing:1.5px">TOTAL ALOKASI DB</div>'
    + '      <div style="font-size:44px;font-weight:900;letter-spacing:-2px;margin:4px 0">' + stats.db + '</div>'
    + '      <div style="font-size:10px;color:#cbd5e1;font-weight:600">' + monthStart + ' ~ ' + monthLabel + '-' + stats.lastDay + '</div>'
    + '    </td>'
    + '    <td width="65%" style="background:#f8fafc;padding:14px;border-radius:11px;vertical-align:top">'
    + '      <table width="100%" cellpadding="0" cellspacing="0"><tr>'
    + '        <td style="font-size:10.5px;color:#64748b;font-weight:800;letter-spacing:1px">TINGKAT AKTIVITAS</td>'
    + '        <td style="text-align:right"><span style="font-size:22px;font-weight:900;color:#1e3a8a">' + stats.actRate + '%</span> <span style="font-size:10.5px;color:#64748b;font-weight:700">' + stats.progressing + ' / ' + stats.db + '</span></td>'
    + '      </tr></table>'
    + '      <div style="height:9px;background:#e2e8f0;border-radius:5px;margin:8px 0 10px">'
    + '        <div style="width:' + stats.actRate + '%;height:100%;background:linear-gradient(90deg,#3b82f6,#1e40af);border-radius:5px"></div>'
    + '      </div>'
    + '      <table width="100%" cellpadding="0" cellspacing="0" style="font-size:11.5px">'
    + '        <tr><td style="padding:4px 0;color:#475569;font-weight:700">&#x1F7E2; Sedang Diproses &middot; 진행중</td><td style="padding:4px 0;text-align:right;color:#16a34a;font-weight:900">' + stats.progressing + '</td></tr>'
    + '        <tr><td style="padding:4px 0;color:#475569;font-weight:700">&#x1F7E1; Tertunda &middot; 보류대기</td><td style="padding:4px 0;text-align:right;color:#f59e0b;font-weight:900">' + stats.waiting + '</td></tr>'
    + '        <tr><td style="padding:4px 0;color:#475569;font-weight:700">&#x1F534; Belum Diproses &middot; 미진행</td><td style="padding:4px 0;text-align:right;color:#dc2626;font-weight:900">' + stats.idle + '</td></tr>'
    + '        <tr><td style="padding:4px 0;color:#475569;font-weight:700">&#x26AA; Lainnya &middot; 기타</td><td style="padding:4px 0;text-align:right;color:#64748b;font-weight:900">' + stats.etc + '</td></tr>'
    + '      </table>'
    + '    </td>'
    + '  </tr></table>'
    + '</td></tr></table>'

    // ─── 이번달 성과 ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:14px;border:1px solid #e2e8f0;margin-bottom:14px">'
    + '<tr><td style="padding:18px">'
    + '  <div style="font-size:12.5px;color:#64748b;margin:0 0 14px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">&#x1F4C8; ' + monthLabel + ' Performa Bulanan &middot; 이번달 성과</div>'
    + '  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">'
    + statRow_('Alokasi DB &middot; 배정 DB', stats.db, '#1e3a8a')
    + statRow_('Visit Selesai &middot; 방문 완료', stats.visit + ' <span style="font-size:13px;color:#64748b;font-weight:700">(' + visitRate + '%)</span>', '#16a34a')
    + statRow_('Closing &middot; 클로징', stats.cl + ' <span style="font-size:13px;color:#64748b;font-weight:700">(' + clRate + '%)</span>', '#dc2626')
    + statRow_('DB &rarr; Closing &middot; DB&rarr;클로', stats.r_dc + '%', '#7c3aed', true)
    + '  </table>'
    + '</td></tr></table>'

    // ─── 에이전트 현황 (1행 표) ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:14px;border:1px solid #e2e8f0;margin-bottom:14px">'
    + '<tr><td style="padding:18px">'
    + '  <div style="font-size:12.5px;color:#64748b;margin:0 0 12px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">&#x1F4C8; Statistik Agen &middot; 에이전트 현황</div>'
    + '  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:11.5px">'
    + '    <thead><tr style="background:#f1f5f9;color:#475569">'
    + '      <th style="padding:8px 6px;text-align:left;font-weight:800;letter-spacing:.3px">Agen</th>'
    + '      <th style="padding:8px 6px;font-weight:800">DB</th>'
    + '      <th style="padding:8px 6px;font-weight:800">Visit</th>'
    + '      <th style="padding:8px 6px;font-weight:800">Closing</th>'
    + '      <th style="padding:8px 6px;font-weight:800">T.Visit</th>'
    + '      <th style="padding:8px 6px;font-weight:800">T.Closing</th>'
    + '      <th style="padding:8px 6px;font-weight:800">DB&rarr;CL</th>'
    + '    </tr></thead>'
    + '    <tbody><tr style="background:#eff6ff">'
    + '      <td style="padding:10px 6px;text-align:left;font-weight:900;color:#1e3a8a">' + name + '</td>'
    + '      <td style="padding:10px 6px;text-align:center;font-weight:900">' + stats.db + '</td>'
    + '      <td style="padding:10px 6px;text-align:center;font-weight:900;color:#16a34a">' + stats.visit + '</td>'
    + '      <td style="padding:10px 6px;text-align:center;font-weight:900;color:#dc2626">' + stats.cl + '</td>'
    + '      <td style="padding:10px 6px;text-align:center">' + stats.r_dv + '%</td>'
    + '      <td style="padding:10px 6px;text-align:center;color:#dc2626;font-weight:900">' + stats.r_vc + '%</td>'
    + '      <td style="padding:10px 6px;text-align:center;color:#7c3aed;font-weight:900">' + stats.r_dc + '%</td>'
    + '    </tr></tbody>'
    + '  </table>'
    + '</td></tr></table>'

    // ─── 긴급 회원 관리 ───
    + buildUrgentSection_(urgent)

    // ─── 방문 캘린더 (4 카드) ───
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:14px;border:1px solid #e2e8f0;margin-bottom:14px">'
    + '<tr><td style="padding:18px">'
    + '  <div style="font-size:12.5px;color:#64748b;margin:0 0 14px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">&#x1F4C5; Kalender Visit &middot; 방문 캘린더 (' + monthLabel + ')</div>'
    + '  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:8px 0"><tr>'
    + '    <td width="25%" style="background:#eff6ff;border:2px solid #dbeafe;border-radius:10px;padding:14px;text-align:center">'
    + '      <div style="font-size:10px;color:#1e3a8a;font-weight:800;letter-spacing:1px">VISIT</div>'
    + '      <div style="font-size:32px;font-weight:900;color:#1e3a8a;letter-spacing:-1.5px;line-height:1">' + cal.visit + '</div>'
    + '      <div style="font-size:10px;color:#64748b;font-weight:700;margin-top:4px">방문</div>'
    + '    </td>'
    + '    <td width="25%" style="background:#fffbeb;border:2px solid #fef3c7;border-radius:10px;padding:14px;text-align:center">'
    + '      <div style="font-size:10px;color:#b45309;font-weight:800;letter-spacing:1px">MEETING</div>'
    + '      <div style="font-size:32px;font-weight:900;color:#b45309;letter-spacing:-1.5px;line-height:1">' + cal.meeting + '</div>'
    + '      <div style="font-size:10px;color:#64748b;font-weight:700;margin-top:4px">미팅</div>'
    + '    </td>'
    + '    <td width="25%" style="background:#f0fdf4;border:2px solid #dcfce7;border-radius:10px;padding:14px;text-align:center">'
    + '      <div style="font-size:10px;color:#166534;font-weight:800;letter-spacing:1px">VISUAL</div>'
    + '      <div style="font-size:32px;font-weight:900;color:#16a34a;letter-spacing:-1.5px;line-height:1">' + cal.visual + '</div>'
    + '      <div style="font-size:10px;color:#64748b;font-weight:700;margin-top:4px">비주얼</div>'
    + '    </td>'
    + '    <td width="25%" style="background:#fef2f2;border:2px solid #fee2e2;border-radius:10px;padding:14px;text-align:center">'
    + '      <div style="font-size:10px;color:#991b1b;font-weight:800;letter-spacing:1px">JADWAL</div>'
    + '      <div style="font-size:32px;font-weight:900;color:#dc2626;letter-spacing:-1.5px;line-height:1">' + cal.jadwal + '</div>'
    + '      <div style="font-size:10px;color:#64748b;font-weight:700;margin-top:4px">일정</div>'
    + '    </td>'
    + '  </tr></table>'
    + buildCalendarTable_(records, name, monthLabel, today, stats.lastDay)
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

function buildCalendarTable_(records, agentName, monthLabel, today, lastDay) {
  const [year, month] = monthLabel.split('-').map(Number);
  const firstDow = new Date(year, month-1, 1).getDay();
  const todayDay = parseInt(today.split('-')[2]);

  const daily = {};
  records.filter(r => r.agent === agentName && r.vd && r.vd.indexOf(monthLabel) === 0)
    .forEach(r => {
      const day = parseInt(r.vd.split('-')[2]);
      const sB = (r.statusB||'').toLowerCase();
      if(!daily[day]) daily[day] = {visit:0, visual:0, meeting:0};
      if(sB.indexOf('visit selesai') >= 0) daily[day].visit++;
      else if(sB.indexOf('langsung') >= 0) daily[day].visual++;
      else if(sB.indexOf('video call') >= 0) daily[day].meeting++;
      else daily[day].meeting = (daily[day].meeting||0); // jadwal 카운트는 표시 X
    });

  const days = ['Min','Sen','Sel','Rab','Kam','Jum','Sab'];
  let html = '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:10px;background:#fff;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;margin-top:10px">';
  html += '<tr style="background:#f8fafc">';
  days.forEach(d => { html += '<th style="padding:6px;font-weight:800;color:#475569;letter-spacing:.5px;border-right:1px solid #e2e8f0;font-size:10px">' + d + '</th>'; });
  html += '</tr>';

  let row = '<tr>';
  for(let i=0; i<firstDow; i++){
    row += '<td style="height:48px;width:14.28%;border-right:1px solid #f1f5f9;border-top:1px solid #f1f5f9"></td>';
  }
  let dow = firstDow;
  for(let d=1; d<=lastDay; d++){
    const info = daily[d];
    const isToday = d === todayDay;
    let content = '<div style="font-weight:' + (isToday?'900':'700') + ';color:' + (isToday?'#92400e':'#64748b') + ';font-size:10.5px">' + d + (isToday?' &middot; Hari ini':'') + '</div>';
    if(info){
      if(info.visit) content += '<div style="font-size:9px;color:#dc2626;font-weight:800;line-height:1.3">&#x25CF; V ' + info.visit + '</div>';
      if(info.visual) content += '<div style="font-size:9px;color:#16a34a;font-weight:800;line-height:1.3">&#x25CF; L ' + info.visual + '</div>';
      if(info.meeting) content += '<div style="font-size:9px;color:#7c3aed;font-weight:800;line-height:1.3">&#x25CF; M ' + info.meeting + '</div>';
    }
    const bg = isToday ? 'background:#fef3c7;' : '';
    row += '<td style="' + bg + 'height:48px;padding:4px 5px;vertical-align:top;border-right:1px solid #f1f5f9;border-top:1px solid #f1f5f9;width:14.28%">' + content + '</td>';
    dow = (dow+1) % 7;
    if(dow === 0 && d < lastDay){ html += row + '</tr>'; row = '<tr>'; }
  }
  while(dow !== 0 && dow < 7){
    row += '<td style="height:48px;border-right:1px solid #f1f5f9;border-top:1px solid #f1f5f9"></td>';
    dow++;
  }
  html += row + '</tr></table>';
  return html;
}

function buildUrgentSection_(urgent) {
  const cats = [
    {key:'simulasi', list:urgent.simulasi, label:'Simulasi Pendapatan &middot; 수익 시뮬', icon:'&#x1F4B0;', bg:'#fef2f2', border:'#fee2e2', color:'#991b1b', pillBg:'#fee2e2'},
    {key:'belum',    list:urgent.belum,    label:'Belum Dihubungi &middot; 영업 미진행', icon:'&#x1F4DE;', bg:'#fffbeb', border:'#fef3c7', color:'#b45309', pillBg:'#fef3c7'},
    {key:'dorman',   list:urgent.dorman,   label:'Pemulihan Dorman &middot; 휴면 회복', icon:'&#x1F634;', bg:'#eef2ff', border:'#e0e7ff', color:'#4338ca', pillBg:'#e0e7ff'},
    {key:'menunggu', list:urgent.menunggu, label:'Menunggu Daftar &middot; 등록 대기',  icon:'&#x1F4DD;', bg:'#f0fdf4', border:'#dcfce7', color:'#15803d', pillBg:'#dcfce7'}
  ];
  let total = 0; cats.forEach(c => total += c.list.length);

  let html = '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:14px;border:1px solid #e2e8f0;margin-bottom:14px">'
    + '<tr><td style="padding:18px">'
    + '  <div style="font-size:12.5px;color:#64748b;margin:0 0 12px;text-transform:uppercase;letter-spacing:1.5px;font-weight:800">&#x1F6A8; Member Prioritas Tinggi &middot; 긴급 회원 관리 (' + total + ')</div>'
    + '  <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:5px 0;margin-bottom:10px"><tr>';
  cats.forEach(c => {
    html += '<td width="25%" style="background:' + c.pillBg + ';color:' + c.color + ';padding:8px;border-radius:8px;text-align:center;font-size:10.5px;font-weight:900">' + c.icon + ' ' + c.list.length + '</td>';
  });
  html += '</tr></table>';

  if(total === 0){
    html += '  <div style="padding:14px;background:#f0fdf4;border:1px dashed #86efac;border-radius:8px;color:#15803d;font-size:12.5px;font-weight:700;text-align:center">&#x2705; Tidak ada member mendesak hari ini. Pertahankan!</div>';
  } else {
    cats.forEach(c => {
      if(c.list.length === 0) return;
      const showList = c.list.slice(0, 10);
      const more = c.list.length - showList.length;
      html += '  <div style="border:1px solid ' + c.border + ';border-radius:9px;margin-bottom:8px;overflow:hidden">';
      html += '    <div style="background:' + c.bg + ';padding:7px 12px;font-size:11px;font-weight:900;color:' + c.color + ';letter-spacing:.2px">' + c.icon + ' ' + c.label + ' (' + c.list.length + ')</div>';
      showList.forEach(m => {
        html += '<table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid ' + c.border + '"><tr>'
          + '<td style="padding:8px 12px;font-size:12px"><b style="color:#0f172a">' + m.name + '</b><br><span style="font-size:10.5px;color:#64748b">' + m.reason + '</span></td>'
          + '<td style="padding:8px 12px;text-align:right;font-size:11px;color:' + c.color + ';font-weight:900;width:60px;white-space:nowrap">D+' + m.days + '</td>'
          + '</tr></table>';
      });
      if(more > 0){
        html += '<div style="padding:6px 12px;font-size:10.5px;color:#64748b;background:#fafafa;text-align:center;font-weight:700">&hellip; dan ' + more + ' member lainnya</div>';
      }
      html += '  </div>';
    });
  }
  html += '</td></tr></table>';
  return html;
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
  const urgent = getUrgentMembers_(records, agentName, today);
  const cal = getCalendarSummary_(records, agentName, monthStart, today);
  Logger.log('  · stats: ' + JSON.stringify(stats));
  Logger.log('  · urgent: simulasi=' + urgent.simulasi.length + ' belum=' + urgent.belum.length + ' dorman=' + urgent.dorman.length + ' menunggu=' + urgent.menunggu.length);
  const html = buildReportHTML_({name: agentName, stats, urgent, cal, records, today, monthStart});
  const subject = '[TEST · 일일 보고 · Laporan Harian] ' + agentName + ' · ' + today;
  GmailApp.sendEmail(toEmail, subject, '', {
    htmlBody: html,
    name: REPORT_CFG.FROM_NAME
  });
  Logger.log('✅ 테스트 메일 발송 완료 → ' + toEmail);
  return {ok: true, agent: agentName, to: toEmail, today, stats};
}
