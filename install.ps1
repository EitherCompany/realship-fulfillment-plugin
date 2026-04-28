# realship-fulfillment v2 자동 설치 스크립트
# ------------------------------------------------
# 사용법: 이 파일을 우클릭 → "PowerShell로 실행"
#        또는 관리자 PowerShell에서 .\install.ps1 실행
# ------------------------------------------------

$ErrorActionPreference = "Stop"

# Claude Desktop은 UWP 앱이라 AppData 경로가 LocalCache 아래로 가상화됨
$ClaudeUwpRoot = "$env:LOCALAPPDATA\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude"
$SkillsRoot = Join-Path $ClaudeUwpRoot "local-agent-mode-sessions\skills-plugin\3b804d49-80d9-4628-afd6-03f669b050b2\3456faf5-3072-462a-881d-882dc982e345\skills"
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceDir  = Join-Path $ScriptDir "skills\realship-fulfillment"

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " realship-fulfillment v2 통합 설치 스크립트" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# 경로 확인
if (-not (Test-Path $SkillsRoot)) {
    Write-Host "[오류] Cowork 스킬 폴더를 찾을 수 없습니다:" -ForegroundColor Red
    Write-Host "  $SkillsRoot"
    Write-Host ""
    Write-Host "Cowork가 이 경로에 설치되어 있는지 확인하세요." -ForegroundColor Yellow
    Read-Host "종료하려면 Enter"
    exit 1
}

if (-not (Test-Path $SourceDir)) {
    Write-Host "[오류] 이 스크립트와 같은 폴더에 'skills\realship-fulfillment' 폴더가 있어야 합니다." -ForegroundColor Red
    Read-Host "종료하려면 Enter"
    exit 1
}

# 1) 기존 real-shipping 백업 후 삭제
$RealShipping = Join-Path $SkillsRoot "real-shipping"
if (Test-Path $RealShipping) {
    $Backup = Join-Path $env:TEMP ("real-shipping_backup_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    Write-Host "[1/3] 기존 real-shipping 폴더 백업 → $Backup" -ForegroundColor Yellow
    Copy-Item -Recurse -Force $RealShipping $Backup
    Write-Host "       삭제 중..."
    Remove-Item -Recurse -Force $RealShipping
    Write-Host "       완료" -ForegroundColor Green
} else {
    Write-Host "[1/3] real-shipping 폴더 없음 (스킵)" -ForegroundColor DarkGray
}

# 2) 기존 realship-fulfillment 백업 후 덮어쓰기
$TargetDir = Join-Path $SkillsRoot "realship-fulfillment"
if (Test-Path $TargetDir) {
    $Backup = Join-Path $env:TEMP ("realship-fulfillment_backup_{0}" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    Write-Host "[2/3] 기존 realship-fulfillment 폴더 백업 → $Backup" -ForegroundColor Yellow
    Copy-Item -Recurse -Force $TargetDir $Backup
    Remove-Item -Recurse -Force $TargetDir
    Write-Host "       기존 폴더 삭제 완료"
} else {
    Write-Host "[2/3] realship-fulfillment 폴더 없음 — 새로 생성합니다" -ForegroundColor DarkGray
}

# 3) 새 버전 복사
Write-Host "[3/3] v2 새 버전 설치 중..." -ForegroundColor Yellow
Copy-Item -Recurse -Force $SourceDir $SkillsRoot
Write-Host "       완료" -ForegroundColor Green

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host " 설치 완료. Cowork를 재시작하세요." -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "확인 방법: Cowork 재시작 후 '실배송 처리해줘'라고 입력해서"
Write-Host "           이 스킬 하나만 트리거되는지 확인하세요."
Write-Host ""
Read-Host "종료하려면 Enter"
