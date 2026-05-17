# ============================================================
#  WINK / GAS 자동 배포 설정 (1회만 실행)
# ============================================================
#  사용법 (PowerShell):
#    cd C:\path\to\dashboard
#    .\scripts\setup-clasp.ps1
#
#  실행 후에는 GAS 코드를 수정할 때마다 아래 2줄만 치면 끝:
#    clasp push
#    clasp deploy -i <deploymentId> -d "변경 설명"
# ============================================================

$ErrorActionPreference = 'Stop'

function Say($msg, $color = 'White') {
    Write-Host $msg -ForegroundColor $color
}
function Ok($msg)   { Say "  [OK] $msg" 'Green' }
function Info($msg) { Say "  $msg" 'Cyan' }
function Warn($msg) { Say "  [!] $msg" 'Yellow' }
function Die($msg)  { Say "  [X] $msg" 'Red'; exit 1 }

Say ""
Say "========================================================" 'Yellow'
Say "  WINK / Apps Script  자동 배포 설정" 'Yellow'
Say "========================================================" 'Yellow'
Say ""

# 1) 작업 위치 확인 (repo 루트에서 실행되어야 함)
if (-not (Test-Path "gas/Code.gs")) {
    Die "gas/Code.gs 가 안 보여요. dashboard 폴더 루트에서 실행하세요."
}
Ok "리포지토리 루트 확인"

# 2) Node.js / npm 확인
$nodeCmd = Get-Command node -ErrorAction SilentlyContinue
$npmCmd  = Get-Command npm  -ErrorAction SilentlyContinue
if (-not $nodeCmd -or -not $npmCmd) {
    Warn "Node.js 가 설치되어 있지 않아요."
    Say  "  먼저 설치하세요: https://nodejs.org/  (LTS 버전 추천)"
    Die  "설치 후 이 스크립트를 다시 실행하세요."
}
Ok "Node.js / npm 확인"

# 3) clasp 설치 (없으면)
$claspCmd = Get-Command clasp -ErrorAction SilentlyContinue
if (-not $claspCmd) {
    Info "clasp 가 없어요. 지금 설치합니다... (npm install -g @google/clasp)"
    npm install -g "@google/clasp"
    if ($LASTEXITCODE -ne 0) {
        Die "clasp 설치 실패. 관리자 권한 PowerShell 로 다시 시도해보세요."
    }
    Ok "clasp 설치 완료"
} else {
    Ok "clasp 이미 설치되어 있음"
}

# 4) Google 로그인 (이미 되어 있으면 스킵)
$rcPath = "$HOME/.clasprc.json"
if (-not (Test-Path $rcPath)) {
    Say ""
    Info "Google 로그인 — 브라우저 창이 열립니다. 본인 계정으로 승인하세요."
    clasp login
    if (-not (Test-Path $rcPath)) {
        Die "로그인이 안 됐어요. 다시 시도하세요."
    }
}
Ok "Google 로그인 됨"

# 5) Script ID 입력
Say ""
Say "----- Script ID 가져오기 -----" 'Yellow'
Say "  1) https://script.google.com 접속"
Say "  2) 기존 대시보드 프로젝트 열기 (일일 보고서 보내는 그것)"
Say "  3) URL 예: https://script.google.com/.../d/AKfycb.../edit"
Say "             ----------------- /d/ 와 /edit 사이가 Script ID"
Say ""
$scriptId = Read-Host "Script ID 를 붙여넣으세요"
if ([string]::IsNullOrWhiteSpace($scriptId)) {
    Die "Script ID 가 비었어요."
}

# 6) .clasp.json 작성
$claspJson = @{ scriptId = $scriptId.Trim(); rootDir = "gas" } | ConvertTo-Json
Set-Content -Path ".clasp.json" -Value $claspJson -Encoding utf8
Ok ".clasp.json 작성 완료 (gitignore 됨)"

# 7) 원격 코드 → 로컬 백업 (덮어쓰기 전에 안전장치)
Say ""
Info "기존 원격 코드를 백업 폴더로 가져옵니다..."
$backupDir = "gas-backup-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
$null = New-Item -ItemType Directory -Path $backupDir
$origClasp = Get-Content ".clasp.json" -Raw
$bkClasp = @{ scriptId = $scriptId.Trim(); rootDir = $backupDir } | ConvertTo-Json
Set-Content -Path ".clasp.json" -Value $bkClasp -Encoding utf8
clasp pull 2>&1 | Out-Null
Set-Content -Path ".clasp.json" -Value $origClasp -Encoding utf8
if ((Get-ChildItem $backupDir -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0) {
    Ok "백업 완료 → $backupDir"
} else {
    Warn "백업이 비어있어요 (원격 코드가 없거나 접근 권한 문제). 그래도 진행합니다."
}

# 8) 새 코드 푸시
Say ""
Info "새 코드를 GAS 에 업로드 중..."
clasp push --force
if ($LASTEXITCODE -ne 0) {
    Die "푸시 실패. clasp push 출력을 확인하세요."
}
Ok "코드 업로드 완료"

# 9) wink_setup 함수 실행 안내
Say ""
Say "----- 시트 / Drive 폴더 초기화 -----" 'Yellow'
Say "  GAS 에디터에서 한 번만 실행해야 하는 게 있어요 (권한 동의용)."
Say "  1) https://script.google.com 에서 프로젝트 열기"
Say "  2) 함수 드롭다운 → 'wink_setup' 선택"
Say "  3) 실행 (재생 버튼) → 권한 모두 허용"
Say "  4) 로그(Logger.log)에 'WINK 초기 설치 완료' 보이면 OK"
Say ""
Read-Host "위 작업이 끝났으면 Enter (스킵하려면 그냥 Enter)" | Out-Null

# 10) 배포
Say ""
Say "----- 기존 배포 업데이트 -----" 'Yellow'
Say "  현재 배포 목록을 가져옵니다..."
Say ""
clasp deployments
Say ""
Say "  위에서 'AKfycby...' 로 시작하는 deploymentId 중,"
Say "  웹앱(에이전트 페이지가 쓰는 그것)을 골라 ID 를 입력하세요."
Say "  (잘 모르겠으면 그냥 Enter — 새 배포를 만들지만 URL 이 바뀝니다)"
Say ""
$depId = Read-Host "deploymentId"
if ([string]::IsNullOrWhiteSpace($depId)) {
    Warn "기존 배포 ID 없이 새 배포 생성합니다. URL 이 바뀔 수 있어요."
    clasp deploy -d "WINK 에이전트 앱 (신규 배포)"
} else {
    clasp deploy -i $depId.Trim() -d "WINK 에이전트 앱 추가 ($(Get-Date -Format 'yyyy-MM-dd'))"
}
if ($LASTEXITCODE -ne 0) {
    Die "배포 실패. clasp deploy 출력을 확인하세요."
}
Ok "배포 완료"

# 끝
Say ""
Say "========================================================" 'Green'
Say "  완료! 이제 폰에서 아래 주소를 열어보세요:" 'Green'
Say "    https://meeoak.github.io/dashboard/agent.html" 'Cyan'
Say "  로그인: devi / 1234" 'Cyan'
Say "========================================================" 'Green'
Say ""
Say "[다음부터 GAS 코드 수정 후 반영하는 법]"
Say "  1) Claude 가 gas/Code.gs 수정 → git pull"
Say "  2) clasp push"
Say "  3) clasp deploy -i $depId -d ""변경 설명"""
Say ""
