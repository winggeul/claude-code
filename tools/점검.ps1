<#
설치 후 점검.

다른 프로그램을 깔고 나서 이걸 한 번 돌리면, CRM4 자동화가 그대로 도는지 확인합니다.
확인만 합니다 - 아무것도 지우거나 올리지 않습니다. 누적본(store.json)도 건드리지 않습니다.

사용:
    powershell -ExecutionPolicy Bypass -File .\점검.ps1
    powershell -ExecutionPolicy Bypass -File .\점검.ps1 -CRM4      CRM4 창까지 확인 (창을 열어 두세요)
    powershell -ExecutionPolicy Bypass -File .\점검.ps1 -Deploy    실제 배포까지 해 봄
#>

param(
    # CRM4 창과 컨트롤을 찾는지까지 본다. 클릭은 하지 않는다. CRM4 통합고객목록 창이 열려 있어야 한다.
    [switch]$CRM4,

    # 집계 결과를 실제로 올려 본다. 데이터는 어차피 그대로라 화면 내용은 바뀌지 않는다.
    [switch]$Deploy
)

$Base = Split-Path -Parent $MyInvocation.MyCommand.Path
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Continue"

$Site = "https://daily-db-count-2mzqhq.vercel.app"
$TemplateUrl = "https://raw.githubusercontent.com/woori-marketing/crm-db-dashboard/refs/heads/main/dashboard/index.html"

$fail = 0
$warn = 0
$no = 0

function Head([string]$t) { Write-Host ""; Write-Host $t -ForegroundColor Cyan }

function Ok([string]$what, [string]$detail = "") {
    Write-Host ("  [정상] " + $what) -ForegroundColor Green
    if ($detail) { Write-Host ("         " + $detail) -ForegroundColor DarkGray }
}
function Bad([string]$what, [string]$detail = "") {
    $script:fail++
    Write-Host ("  [문제] " + $what) -ForegroundColor Red
    if ($detail) { Write-Host ("         " + $detail) -ForegroundColor DarkGray }
}
function Warn([string]$what, [string]$detail = "") {
    $script:warn++
    Write-Host ("  [주의] " + $what) -ForegroundColor Yellow
    if ($detail) { Write-Host ("         " + $detail) -ForegroundColor DarkGray }
}
function Skip([string]$what, [string]$detail = "") {
    $script:no++
    Write-Host ("  [생략] " + $what) -ForegroundColor DarkGray
    if ($detail) { Write-Host ("         " + $detail) -ForegroundColor DarkGray }
}

# 바깥 프로그램이 stderr 에 한 줄만 써도 오류로 올라온다. 종료 코드로만 판단한다.
function Run([string]$exe, [string[]]$argv) {
    try {
        $out = & $exe @argv 2>&1 | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { "$_" }
        } | Where-Object { "$_".Trim() -ne "" }
        return [pscustomobject]@{ Out = ($out -join "`n"); Code = $LASTEXITCODE }
    } catch {
        return [pscustomobject]@{ Out = $_.Exception.Message; Code = -1 }
    }
}

Write-Host "=== 설치 후 점검 ===" -ForegroundColor Cyan
Write-Host ("폴더: " + $Base)
Write-Host ("시각: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))

# ─── 1. Node.js ───────────────────────────────────

Head "1. Node.js"

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    Bad "node 를 찾지 못했습니다" "새로 깐 프로그램이 PATH 를 건드렸을 수 있습니다. PowerShell 을 새로 열어 보고, 그래도 없으면 Node.js 를 다시 설치하세요."
} else {
    $v = Run "node" @("-v")
    if ($v.Code -eq 0) { Ok "node $($v.Out)" $node.Source } else { Bad "node 가 실행되지 않습니다" $v.Out }
}

$npx = Get-Command npx.cmd -ErrorAction SilentlyContinue
if (-not $npx) { $npx = Get-Command npx -ErrorAction SilentlyContinue }
if (-not $npx) {
    Bad "npx 를 찾지 못했습니다" "배포가 안 됩니다. Node.js 설치를 확인하세요."
} elseif ($npx.Source -notlike "*.cmd") {
    Warn "npx.cmd 가 아니라 $($npx.Source) 로 잡힙니다" "실행 정책에 막힐 수 있습니다. 배포 확인(6번)에서 결과를 보세요."
} else {
    Ok "npx" $npx.Source
}

# ─── 2. 파일 ──────────────────────────────────────

Head "2. 파일"

$need = @{
    "crm4_daily.ps1"   = "매일 도는 전체 작업"
    "crm4_collect.ps1" = "CRM4 에서 CSV 받기"
    "aggregate.mjs"    = "세어서 화면 만들기"
    "roster.mjs"       = "담당자 명단 (없으면 집계가 멈춥니다)"
    "template.html"    = "화면 틀"
    "store.json"       = "누적본"
    "package.json"     = "인원 저장 기능이 쓰는 목록"
    "api\headcount.js" = "인원 저장하는 자리"
}
foreach ($f in ($need.Keys | Sort-Object)) {
    $p = Join-Path $Base $f
    if (Test-Path $p) {
        $len = (Get-Item $p).Length
        if ($len -eq 0) { Bad "$f 가 비어 있습니다" $need[$f] }
        else { Ok $f ("{0:N0} 바이트 · {1}" -f $len, $need[$f]) }
    } else {
        Bad "$f 가 없습니다" $need[$f]
    }
}

# ─── 3. 누적본 ────────────────────────────────────

Head "3. 누적본"

$Store = Join-Path $Base "store.json"
if (-not (Test-Path $Store)) {
    Bad "store.json 이 없습니다" "지금까지 쌓인 숫자가 사라집니다. 백업에서 되돌리세요."
} else {
    $raw = Get-Content $Store -Raw -Encoding UTF8
    $hit = [regex]::Matches($raw, '"date"\s*:\s*"(\d{4}-\d{2}-\d{2})"')
    if ($hit.Count -eq 0) {
        Bad "store.json 을 읽지 못했습니다" "내용이 깨졌을 수 있습니다."
    } else {
        $dates = @($hit | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $last = $dates[-1]
        $yesterday = (Get-Date).Date.AddDays(-1)
        $gap = ($yesterday - [datetime]$last).Days
        $msg = "$($hit.Count)줄 · $($dates.Count)일 · $($dates[0]) ~ $last"
        if ($gap -le 1) { Ok "누적본 정상" $msg }
        elseif ($gap -le 3) { Warn "마지막 수집이 $gap 일 전입니다" $msg }
        else { Bad "마지막 수집이 $gap 일 전입니다" "$msg — 수집이 멈춰 있었을 수 있습니다. logs 폴더를 보세요." }

        if ($raw -match '"branch"') { Ok "지점(시흥·천안)이 붙어 있습니다" }
        else { Bad "지점이 붙어 있지 않습니다" "옛 store.json 입니다. 새 것으로 덮어쓰세요." }
    }
}

# ─── 4. 집계 ──────────────────────────────────────

Head "4. 집계 (진짜 파일은 건드리지 않습니다)"

$tmp = Join-Path $env:TEMP ("crm4_check_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path (Join-Path $tmp "raw") -Force | Out-Null

$agg = Join-Path $Base "aggregate.mjs"
$tpl = Join-Path $Base "template.html"
if ((Test-Path $agg) -and (Test-Path $tpl) -and (Test-Path $Store) -and $node) {
    Copy-Item $Store (Join-Path $tmp "store.json") -Force
    $r = Run "node" @($agg, "--raw", (Join-Path $tmp "raw"), "--store", (Join-Path $tmp "store.json"),
                      "--template", $tpl, "--out", (Join-Path $tmp "index.html"))
    $out = Join-Path $tmp "index.html"
    if ($r.Code -ne 0) {
        Bad "집계가 실패했습니다" $r.Out
    } elseif (-not (Test-Path $out)) {
        Bad "결과 파일이 만들어지지 않았습니다" $r.Out
    } else {
        $size = (Get-Item $out).Length
        $html = Get-Content $out -Raw -Encoding UTF8
        if ($size -lt 100000) { Bad "결과가 너무 작습니다 ($size 바이트)" }
        elseif ($html -notmatch "시흥") { Bad "결과에 지점이 들어가지 않았습니다" "roster.mjs 나 store.json 을 확인하세요." }
        else { Ok "집계 정상" ("{0:N0} 바이트 · 지점 열 있음" -f $size) }
    }
} else {
    Skip "집계" "앞 단계에서 빠진 것이 있습니다."
}

Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

# ─── 5. 인터넷 ────────────────────────────────────

Head "5. 인터넷"

try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 } catch {}

function Fetch([string]$url) {
    try {
        $res = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 20
        return [pscustomobject]@{ Ok = $true; Code = $res.StatusCode; Body = $res.Content }
    } catch {
        return [pscustomobject]@{ Ok = $false; Code = 0; Body = $_.Exception.Message }
    }
}

$t = Fetch $TemplateUrl
if ($t.Ok -and $t.Body.Length -gt 10000) { Ok "화면 틀을 받아옵니다 (깃허브)" ("{0:N0} 바이트" -f $t.Body.Length) }
else { Bad "화면 틀을 받지 못했습니다" "새로 깐 프로그램의 방화벽·백신이 막았을 수 있습니다. $($t.Body)" }

$s = Fetch $Site
if ($s.Ok -and $s.Body -match "시흥") { Ok "대시보드가 열립니다" $Site }
elseif ($s.Ok) { Warn "대시보드는 열리는데 지점이 안 보입니다" "아직 새 화면이 안 올라갔을 수 있습니다." }
else { Bad "대시보드에 닿지 못했습니다" $s.Body }

$h = Fetch ($Site + "/api/headcount")
if (-not $h.Ok) {
    Bad "인원 저장 기능에 닿지 못했습니다" $h.Body
} elseif ($h.Body -match '"fallback"') {
    Bad "인원 저장소를 읽지 못하고 있습니다" "버셀에서 BLOB_READ_WRITE_TOKEN 이 살아 있는지 보세요. 응답: $($h.Body)"
} elseif ($h.Body -match '"updated"\s*:\s*"') {
    Ok "인원 저장소 정상" $h.Body
} else {
    Bad "인원 저장 기능의 응답이 이상합니다" $h.Body
}

# ─── 6. 배포 ──────────────────────────────────────

Head "6. 배포"

$tok = [Environment]::GetEnvironmentVariable("VERCEL_TOKEN", "Machine")
if (-not $tok) { $tok = $env:VERCEL_TOKEN }
if (-not $tok) {
    Bad "VERCEL_TOKEN 이 없습니다" "배포가 안 됩니다. SETUP.md 5번을 다시 보세요."
} else {
    Ok "VERCEL_TOKEN 있음" ("$($tok.Length)자 · 값은 찍지 않습니다")
}

if ($Deploy) {
    $daily = Join-Path $Base "crm4_daily.ps1"
    if (-not (Test-Path $daily)) {
        Bad "crm4_daily.ps1 이 없어 배포를 못 해봅니다"
    } else {
        Write-Host "  배포해 보는 중… (몇 분 걸릴 수 있습니다)" -ForegroundColor DarkGray
        $d = Run "powershell" @("-ExecutionPolicy", "Bypass", "-File", $daily, "-SkipCollect")
        if ($d.Code -eq 0) { Ok "배포까지 정상" (($d.Out -split "`n" | Select-Object -Last 1)) }
        else { Bad "배포가 정상 종료되지 않았습니다 (코드 $($d.Code))" (($d.Out -split "`n" | Select-Object -Last 6) -join "`n         ") }
    }
} else {
    Skip "실제 배포" "해 보려면 -Deploy 를 붙이세요."
}

# ─── 7. 자동 실행 ─────────────────────────────────

Head "7. 자동 실행"

foreach ($name in @("CRM4 일일수집", "CRM4 화면갱신")) {
    $task = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    if (-not $task) {
        Bad "$name 작업이 없습니다" "새로 깐 프로그램이 지웠거나 등록이 풀렸습니다. SETUP.md 6번으로 다시 등록하세요."
        continue
    }
    $info = Get-ScheduledTaskInfo -TaskName $name -ErrorAction SilentlyContinue
    $detail = "마지막 $($info.LastRunTime) · 결과 $($info.LastTaskResult) · 다음 $($info.NextRunTime)"
    if ($task.State -eq "Disabled") { Bad "$name 이 사용 안 함 상태입니다" $detail }
    elseif ($info -and $info.LastTaskResult -ne 0 -and $info.LastTaskResult -ne 267011) { Warn "$name 의 지난 실행이 실패로 끝났습니다" $detail }
    else { Ok $name $detail }
}

# ─── 8. 화면이 꺼지지 않는지 ──────────────────────

Head "8. 화면·절전"

function PowerIdle([string]$sub, [string]$setting) {
    try {
        $q = powercfg /query SCHEME_CURRENT $sub $setting 2>&1 | Out-String
        $m = [regex]::Match($q, '(?im)^\s*.*\bAC\b.*?:\s*0x([0-9a-fA-F]{8})')
        if ($m.Success) { return [Convert]::ToInt32($m.Groups[1].Value, 16) }
    } catch {}
    return -1
}

$video = PowerIdle "SUB_VIDEO" "VIDEOIDLE"
$sleep = PowerIdle "SUB_SLEEP" "STANDBYIDLE"

if ($video -lt 0) { Skip "화면 끄기 설정을 읽지 못했습니다" "설정 › 시스템 › 전원 에서 직접 확인하세요." }
elseif ($video -eq 0) { Ok "화면 끄기 안 함" }
else { Bad "화면이 $([int]($video/60))분 뒤 꺼집니다" "꺼지면 새벽 수집이 실패합니다. 설정 › 시스템 › 전원 에서 '안 함' 으로 바꾸세요." }

if ($sleep -lt 0) { Skip "절전 설정을 읽지 못했습니다" }
elseif ($sleep -eq 0) { Ok "절전 안 함" }
else { Bad "$([int]($sleep/60))분 뒤 절전으로 들어갑니다" "설정 › 시스템 › 전원 에서 '안 함' 으로 바꾸세요." }

$saver = (Get-ItemProperty "HKCU:\Control Panel\Desktop" -Name ScreenSaveActive -ErrorAction SilentlyContinue).ScreenSaveActive
$secure = (Get-ItemProperty "HKCU:\Control Panel\Desktop" -Name ScreenSaverIsSecure -ErrorAction SilentlyContinue).ScreenSaverIsSecure
if ($saver -eq "1" -and $secure -eq "1") {
    Bad "화면 보호기가 잠금으로 설정돼 있습니다" "잠기면 수집이 실패합니다. 화면 보호기 설정에서 '다시 시작할 때 로그온 화면 표시' 를 끄세요."
} else {
    Ok "화면 보호기 잠금 없음"
}

# ─── 9. CRM4 ──────────────────────────────────────

Head "9. CRM4"

if (-not $CRM4) {
    Skip "CRM4 창 확인" "해 보려면 통합고객목록 창을 열고 -CRM4 를 붙이세요."
} else {
    $col = Join-Path $Base "crm4_collect.ps1"
    if (-not (Test-Path $col)) {
        Bad "crm4_collect.ps1 이 없습니다"
    } else {
        $rawdir = Join-Path $Base "raw"
        New-Item -ItemType Directory -Path $rawdir -Force | Out-Null
        $c = Run "powershell" @("-ExecutionPolicy", "Bypass", "-File", $col, "-RawRoot", $rawdir, "-Diagnose")
        if ($c.Code -eq 0) { Ok "CRM4 창과 컨트롤을 찾았습니다" "클릭은 하지 않았습니다." }
        else { Bad "CRM4 를 찾지 못했습니다 (코드 $($c.Code))" (($c.Out -split "`n" | Select-Object -Last 8) -join "`n         ") }
    }
}

# ─── 정리 ─────────────────────────────────────────

Write-Host ""
Write-Host "────────────────────────────────" -ForegroundColor DarkGray
if ($fail -eq 0 -and $warn -eq 0) {
    Write-Host "다 정상입니다. 내일 아침 6시에 평소대로 돕니다." -ForegroundColor Green
} elseif ($fail -eq 0) {
    Write-Host "문제 없음 · 주의 $warn 건. 위의 [주의] 항목만 보시면 됩니다." -ForegroundColor Yellow
} else {
    Write-Host "문제 $fail 건 · 주의 $warn 건. 위의 [문제] 항목을 먼저 보세요." -ForegroundColor Red
}
if ($no) { Write-Host "생략 $no 건 (-CRM4 / -Deploy 를 붙이면 함께 봅니다)" -ForegroundColor DarkGray }
Write-Host ""
Write-Host "한글 입력기는 이 점검으로 못 잡습니다. 수집은 붙여넣기로 넣게 해 두었지만," -ForegroundColor DarkGray
Write-Host "새 프로그램이 입력기를 바꿨다면 내일 아침 로그에 저장 실패가 보일 수 있습니다." -ForegroundColor DarkGray

exit ([int]($fail -gt 0))
