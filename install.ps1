# ==============================================================================
# [학습일지 평가] 스킬 자동 설치 및 배포 스크립트 (install.ps1)
# 대상 환경: Antigravity (Gemini) & Claude Code
# 원광대학교 한의과대학 경혈학2 학습일지 평가 자동화 시스템 v3.2
# ==============================================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " [학습일지 평가] 스킬 자동 배포 및 환경 설정 도구 (v3.2) " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$SourceDir = $PSScriptRoot
if (-not $SourceDir) {
    $SourceDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

Write-Host "[1/4] 원본 스킬 경로 확인: $SourceDir" -ForegroundColor Yellow

# 필수 파일 존재 검증
$RequiredFiles = @("SKILL.md", "scripts\eval_engine.py", "scripts\rename_logs.py", "references\rubric.md")
foreach ($rf in $RequiredFiles) {
    $checkPath = Join-Path $SourceDir $rf
    if (-not (Test-Path $checkPath)) {
        Write-Error "필수 파일이 누락되었습니다: $checkPath"
        exit 1
    }
}
Write-Host "  -> 필수 파일 검증 완료 (정상)" -ForegroundColor Green

# 2. 설치 타깃 경로 정의
$UserHome = [System.Environment]::GetFolderPath('UserProfile')
$AntigravitySkillDir = Join-Path $UserHome ".gemini\config\skills\study-log-eval"
$ClaudeSkillDir = Join-Path $UserHome ".claude\skills\study-log-eval"

Write-Host "`n[2/4] 스킬 디렉토리 배포 중..." -ForegroundColor Yellow

function Deploy-SkillDirectory {
    param (
        [string]$TargetDir,
        [string]$PlatformName
    )
    
    if (-not (Test-Path $TargetDir)) {
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
        Write-Host "  + [$PlatformName] 디렉토리 생성: $TargetDir" -ForegroundColor Cyan
    } else {
        Write-Host "  * [$PlatformName] 기존 디렉토리 갱신: $TargetDir" -ForegroundColor Cyan
    }

    Copy-Item -Path (Join-Path $SourceDir "SKILL.md") -Destination $TargetDir -Force
    Copy-Item -Path (Join-Path $SourceDir "학습일지_평가지침.md") -Destination $TargetDir -Force -ErrorAction SilentlyContinue

    $scriptsDest = Join-Path $TargetDir "scripts"
    if (-not (Test-Path $scriptsDest)) { New-Item -ItemType Directory -Path $scriptsDest -Force | Out-Null }
    Copy-Item -Path (Join-Path $SourceDir "scripts\*") -Destination $scriptsDest -Recurse -Force

    $refDest = Join-Path $TargetDir "references"
    if (-not (Test-Path $refDest)) { New-Item -ItemType Directory -Path $refDest -Force | Out-Null }
    Copy-Item -Path (Join-Path $SourceDir "references\*") -Destination $refDest -Recurse -Force

    Write-Host "  -> [$PlatformName] 스킬 배포 완료!" -ForegroundColor Green
}

# Antigravity 배포
Deploy-SkillDirectory -TargetDir $AntigravitySkillDir -PlatformName "Antigravity (Gemini)"

# Claude Code 배포
Deploy-SkillDirectory -TargetDir $ClaudeSkillDir -PlatformName "Claude Code"

# 3. Python 의존성 검사
Write-Host "`n[3/4] Python 런타임 및 필수 라이브러리 검사..." -ForegroundColor Yellow
$PythonCmd = Get-Command python -ErrorAction SilentlyContinue
if ($null -eq $PythonCmd) {
    Write-Warning "Python 명령어를 찾을 수 없습니다. Python 3.9+ 및 pypdf, openpyxl을 수동 설치해주세요."
} else {
    Write-Host "  -> 감지된 Python: $($PythonCmd.Source)" -ForegroundColor Green
    
    $checkProc = Start-Process -FilePath "python" -ArgumentList "-c", "`"import pypdf, openpyxl`"" -Wait -PassThru -NoNewWindow
    if ($checkProc.ExitCode -eq 0) {
        Write-Host "  -> 필수 패키지 (pypdf, openpyxl)가 이미 정상 설치되어 있습니다." -ForegroundColor Green
    } else {
        Write-Host "  ! 필수 패키지 미설치 감지. 자동 설치를 진행합니다..." -ForegroundColor DarkYellow
        $installProc = Start-Process -FilePath "python" -ArgumentList "-m", "pip", "install", "--upgrade", "pypdf", "openpyxl" -Wait -PassThru -NoNewWindow
        if ($installProc.ExitCode -eq 0) {
            Write-Host "  -> 필수 패키지 (pypdf, openpyxl) 설치 성공!" -ForegroundColor Green
        } else {
            Write-Warning "패키지 자동 설치에 실패했습니다. pip install pypdf openpyxl을 직접 실행해주세요."
        }
    }
}

# 4. 배포 완료 및 사용 안내
Write-Host "`n[4/4] 설치 및 배포가 완료되었습니다!" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " [스킬 호출 방법]" -ForegroundColor Yellow
Write-Host "   1) 프롬프트에 [학습일지평가] 또는 [학습일지 평가] 입력"
Write-Host "   2) CLI 직접 실행 예시:"
Write-Host "      python `"$SourceDir\scripts\eval_engine.py`" --target `"폴더경로`""
Write-Host "==========================================================" -ForegroundColor Cyan
