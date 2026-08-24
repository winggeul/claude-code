<#
매일 한 번 도는 전체 작업.

  수집 → 집계 → 배포 → 원본 삭제

작업 스케줄러가 이 파일 하나만 실행하면 된다.

사용:
    powershell -ExecutionPolicy Bypass -File crm4_daily.ps1
    powershell -ExecutionPolicy Bypass -File crm4_daily.ps1 -Start 2026-08-20 -Backfill        지난 날짜 (기준이 달라짐)
    powershell -ExecutionPolicy Bypass -File crm4_daily.ps1 -SkipCollect                        이미 받은 CSV 만 집계
#>

param(
    # 수집 범위. 비워 두면 어제 하루만 받는다.
    # 빠진 날을 저절로 메우지 않는다. 지난 날짜를 지금 뽑으면 그날 뽑았을 때가 아니라
    # 오늘 기준으로 갱신된 값이 나와서, 다음날 받아 둔 다른 날들과 기준이 어긋난다.
    [datetime]$Start,
    [datetime]$End = (Get-Date).Date.AddDays(-1),

    # 지난 날짜를 받겠다고 분명히 밝히는 스위치.
    [switch]$Backfill,

    # 수집을 건너뛰고 이미 있는 CSV 만 집계한다. 과거분 파일을 직접 넣었을 때 쓴다.
    [switch]$SkipCollect,

    # 배포까지 하지 않고 결과 파일만 만든다.
    [switch]$SkipDeploy,

    # 집계 후에도 원본 CSV 를 남긴다. 숫자를 대조해야 할 때만 쓴다.
    [switch]$KeepRaw
)

# ─────────── 설정 (한 번만 고치면 된다) ───────────

# 작업 폴더. 이 스크립트와 aggregate.mjs, template.html 이 함께 있는 곳.
$Base = Split-Path -Parent $MyInvocation.MyCommand.Path

# 내려받은 CSV 가 잠시 머무는 곳. 집계가 끝나면 지워지므로 이 PC 안이면 된다.
$RawRoot = Join-Path $Base "raw"

# 누적 집계본. 원본을 지워도 지금까지 센 결과는 여기 남는다. 건수뿐이라 개인정보가 없다.
$Store = Join-Path $Base "store.json"

# 배포할 파일이 놓이는 폴더. 버셀은 이 폴더를 통째로 올린다.
$SiteDir = Join-Path $Base "site"
$OutFile = Join-Path $SiteDir "index.html"

$Template = Join-Path $Base "template.html"
$Collect  = Join-Path $Base "crm4_collect.ps1"
$Agg      = Join-Path $Base "aggregate.mjs"

# 버셀 토큰. 환경변수에 두는 편이 안전하지만, 없으면 아래에 직접 적어도 된다.
$VercelToken = $env:VERCEL_TOKEN

# ─────────────────────────────────────────────────

$ErrorActionPreference = "Stop"

# node 는 UTF-8 로 출력하는데 PowerShell 5.1 은 시스템 기본 인코딩으로 읽어 한글이 깨진다.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$log = New-Object System.Collections.Generic.List[string]

# 로그는 여기 한 곳에만 쌓는다. 수집기도 같은 파일에 이어 쓴다.
$LogDir = Join-Path $Base "logs"
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$LogFile = Join-Path $LogDir ("run_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")

function Say([string]$msg, [string]$color = "Gray") {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    Write-Host $line -ForegroundColor $color
    $log.Add($line)
}

function Fail([string]$msg) {
    Say $msg "Red"
    $log | Out-File $LogFile -Encoding utf8 -Append
    Write-Host "로그: $LogFile"
    exit 1
}

Say "=== 시작 ===" "Cyan"

foreach ($f in @($Template, $Agg)) {
    if (-not (Test-Path $f)) { Fail "파일이 없습니다: $f" }
}
New-Item -ItemType Directory -Path $RawRoot -Force | Out-Null
New-Item -ItemType Directory -Path $SiteDir -Force | Out-Null

# ─── 1. 수집 ───────────────────────────────────────

if ($SkipCollect) {
    Say "수집 건너뜀"
} else {
    if (-not (Test-Path $Collect)) { Fail "파일이 없습니다: $Collect" }

    $collectArgs = @("-ExecutionPolicy", "Bypass", "-File", $Collect, "-RawRoot", $RawRoot,
                     "-LogFile", $LogFile, "-End", $End.ToString("yyyy-MM-dd"))
    if ($PSBoundParameters.ContainsKey("Start")) { $collectArgs += @("-Start", $Start.ToString("yyyy-MM-dd")) }
    if ($Backfill) { $collectArgs += "-Backfill" }

    Say "수집 시작" "Cyan"
    $log | Out-File $LogFile -Encoding utf8 -Append    # 수집기가 이어 쓰기 전에 여기까지 먼저 남긴다
    $log.Clear()
    & powershell @collectArgs
    if ($LASTEXITCODE -eq 2) {
        Say "지난 날짜라 받지 않았습니다. 기준을 맞추려면 그날 다음날에 받아야 합니다." "Yellow"
    } elseif ($LASTEXITCODE -ne 0) {
        Say "수집이 정상 종료되지 않았습니다 (코드 $LASTEXITCODE). 받아둔 파일만으로 계속합니다." "Yellow"
    }
}

# ─── 2. 집계 ───────────────────────────────────────

$csv = @(Get-ChildItem $RawRoot -Filter *.csv -ErrorAction SilentlyContinue)
if ($csv.Count -eq 0 -and -not (Test-Path $Store)) { Fail "집계할 CSV 도 누적본도 없습니다." }

Say "집계 시작 (CSV $($csv.Count)개)" "Cyan"

$aggArgs = @($Agg, "--raw", $RawRoot, "--store", $Store, "--template", $Template, "--out", $OutFile)
if (-not $KeepRaw) { $aggArgs += "--delete" }

$aggOut = & node @aggArgs 2>&1
$aggOut | ForEach-Object { Say "  $_" }
if ($LASTEXITCODE -ne 0) { Fail "집계 실패. 원본 CSV 는 지우지 않았으므로 다음 실행에서 다시 시도됩니다." }
if (-not (Test-Path $OutFile)) { Fail "결과 파일이 만들어지지 않았습니다." }

# ─── 3. 배포 ───────────────────────────────────────

if ($SkipDeploy) {
    Say "배포 건너뜀 - 결과: $OutFile" "Green"
} elseif (-not $VercelToken) {
    Say "VERCEL_TOKEN 이 없어 배포를 건너뜁니다. 결과: $OutFile" "Yellow"
} else {
    Say "배포 시작" "Cyan"
    $depOut = & npx --yes vercel deploy $SiteDir --prod --yes --token $VercelToken 2>&1
    $depOut | ForEach-Object { Say "  $_" }
    if ($LASTEXITCODE -ne 0) { Fail "배포 실패. 집계 결과는 $OutFile 에 있습니다." }
    Say "배포 완료" "Green"
}

Say "=== 끝 ===" "Cyan"
$log | Out-File $LogFile -Encoding utf8 -Append
Write-Host "로그: $LogFile"
