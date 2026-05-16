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

// ⚠️ 자신의 스프레드시트 ID로 교체
const SPREADSHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID_HERE';

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

// 간단한 health check
function doGet(){
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, msg:'webhook alive', t:new Date().toISOString()}))
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
