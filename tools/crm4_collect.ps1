<#
통합고객목록 자동 수집기

등록일을 하루씩 밀며 [검색 → 목록 탭 → 전체선택 → 처리 탭 → Excel → 저장] 을 반복한다.
74 만 건을 한 번에 받지 않고 날짜로 자르는 이유는 화면이 그만큼을 한 번에 못 받기 때문이다.

컨트롤 식별자는 crm4_probe.ps1 덤프에서 확인한 실제 값이다.
설치할 것 없이 PowerShell 만으로 동작한다.

사용법:
    powershell -ExecutionPolicy Bypass -File crm4_collect.ps1
    powershell -ExecutionPolicy Bypass -File crm4_collect.ps1 -Start 2020-01-01 -End 2020-12-31
#>

param(
    # 수집할 등록일 범위. 기본은 어제 하루.
    [datetime]$Start = (Get-Date).Date.AddDays(-1),
    [datetime]$End   = (Get-Date).Date.AddDays(-1),

    # 저장 위치. 반드시 직접 지정한다. 스크립트가 폴더를 만들지는 않는다.
    [Parameter(Mandatory = $true)]
    [string]$OutputRoot,

    # 등록일 조건 체크박스를 스크립트가 켤지 여부. 이미 켜 둔 상태라면 -SkipDateCheck 로 끈다.
    [switch]$SkipDateCheck,

    # 처리 탭의 Excel 이 '선택한 고객' 기준으로 동작하는 경우를 대비해 전체선택을 먼저 누른다.
    [switch]$SkipSelectAll,

    # 클릭 없이 창과 컨트롤만 확인한다. 처음 쓸 때 이걸로 먼저 점검할 것.
    [switch]$DryRun,

    # 각 단계 사이 기본 대기(초). 서버가 느리면 늘린다.
    [int]$Wait = 2
)

# ─────────── 컨트롤 식별자 (probe 덤프 실측값) ───────────

$ID = @{
    Tab         = 16       # SysTabControl32
    PageFilter  = 527392   # 고객조건 페이지
    PageList    = 199984   # 목록 페이지
    PageProcess = 592892   # 처리 페이지
    DateCheck   = 68906    # '등록일' 체크박스
    DateFrom    = 68896    # SysDateTimePick32
    DateTo      = 68902    # SysDateTimePick32
    Search      = 724064   # 검색 버튼
    CountLabel  = 199986   # '전체고객 : ... 조회고객 : N 명'
    SelectAll   = 396598   # 전체선택
    Excel       = 68956    # 처리 탭 기능 그룹의 Excel 버튼
}

$TabIndex = @{ Filter = 0; List = 4; Process = 5 }

# ─────────── Win32 ───────────

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class C4 {
    public delegate bool EnumProc(IntPtr h, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, IntPtr e);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    public static string ClassOf(IntPtr h){ StringBuilder s=new StringBuilder(256); GetClassName(h,s,256); return s.ToString(); }
    public static string TextOf(IntPtr h){ StringBuilder s=new StringBuilder(1024); GetWindowText(h,s,1024); return s.ToString(); }
}
"@

$script:Log = @()
function Write-Log([string]$msg, [string]$color = "Gray") {
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg
    Write-Host $line -ForegroundColor $color
    $script:Log += $line
}

# PowerShell 스크립트블록을 델리게이트로 넘기면, 콜백 안에서 그 함수의 지역 변수가
# 보이지 않는 경우가 있다. 그래서 콜백에서는 스크립트 범위 변수에 담기만 하고
# 조건 판단은 전부 콜백 바깥에서 한다.

function Get-TopWindows {
    $script:tops = @()
    $cb = [C4+EnumProc]{
        param($h, $l)
        $o = 0
        [void][C4]::GetWindowThreadProcessId($h, [ref]$o)
        $script:tops += [pscustomobject]@{
            Handle  = $h
            OwnerPid= [int]$o
            Title   = [C4]::TextOf($h)
            Class   = [C4]::ClassOf($h)
            Visible = [C4]::IsWindowVisible($h)
        }
        return $true
    }
    [void][C4]::EnumWindows($cb, [IntPtr]::Zero)
    return $script:tops
}

function Find-MainWindow([int]$processId) {
    $m = @(Get-TopWindows | Where-Object {
        $_.OwnerPid -eq $processId -and $_.Visible -and $_.Title -like "*통합고객목록*"
    })
    if ($m.Count -gt 0) { return $m[0].Handle }
    return [IntPtr]::Zero
}

function Update-ControlMap([IntPtr]$root) {
    $script:ctrls = @{}
    $cb = [C4+EnumProc]{
        param($h, $l)
        $script:ctrls[[C4]::GetDlgCtrlID($h)] = $h
        return $true
    }
    [void][C4]::EnumChildWindows($root, $cb, [IntPtr]::Zero)
    return $script:ctrls.Count
}

function Get-Control([IntPtr]$root, [int]$controlId) {
    if (-not $script:ctrls -or -not $script:ctrls.ContainsKey($controlId)) {
        [void](Update-ControlMap $root)
    }
    if (-not $script:ctrls.ContainsKey($controlId)) {
        throw "컨트롤 $controlId 을 찾지 못했습니다."
    }
    return $script:ctrls[$controlId]
}

# 이 프로그램의 버튼은 대부분 직접 그린 컨트롤이라 BM_CLICK 이 통하지 않는다.
# 실행 시점의 좌표를 다시 읽어 실제 마우스로 누른다. 창을 옮겨도 스스로 맞춰진다.
function Click-Control([IntPtr]$h, [int]$offsetX = 0) {
    # 좌표로 누르는 방식이라, 대상 창이 맨 앞이 아니면 엉뚱한 창을 클릭하게 된다.
    # SetForegroundWindow 는 윈도우 정책상 조용히 실패할 수 있으므로 매번 확인하고,
    # 확인이 안 되면 클릭하지 않고 즉시 중단한다.
    $fg = [C4]::GetForegroundWindow()
    if ($fg -ne $script:TargetWindow) {
        throw "통합고객목록 창이 맨 앞이 아닙니다. 다른 창을 클릭할 위험이 있어 중단합니다."
    }

    $r = New-Object C4+RECT
    [void][C4]::GetWindowRect($h, [ref]$r)
    $x = if ($offsetX -gt 0) { $r.Left + $offsetX } else { [int](($r.Left + $r.Right) / 2) }
    $y = [int](($r.Top + $r.Bottom) / 2)
    [void][C4]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 150
    [C4]::mouse_event(0x02, 0, 0, 0, [IntPtr]::Zero)   # LEFTDOWN
    Start-Sleep -Milliseconds 60
    [C4]::mouse_event(0x04, 0, 0, 0, [IntPtr]::Zero)   # LEFTUP
    Start-Sleep -Milliseconds 400
}

# 탭 전환 후 해당 페이지가 실제로 보이는지로 성공을 확인한다.
function Switch-Tab([IntPtr]$win, [string]$name) {
    $tab = Get-Control $win $ID.Tab
    [void][C4]::SendMessage($tab, 0x1330, [IntPtr]$TabIndex[$name], [IntPtr]::Zero)  # TCM_SETCURFOCUS
    Start-Sleep -Milliseconds 700

    $page = Get-Control $win $ID["Page$name"]
    if (-not [C4]::IsWindowVisible($page)) { throw "$name 탭으로 전환하지 못했습니다." }
}

# 키 입력도 포커스가 엉뚱한 곳이면 그대로 그 창에 타이핑된다. 같은 확인을 거친다.
function Send-Keys([string]$keys, [IntPtr]$expectWindow) {
    if ([C4]::GetForegroundWindow() -ne $expectWindow) {
        throw "입력 대상 창이 맨 앞이 아닙니다. 엉뚱한 곳에 입력될 위험이 있어 중단합니다."
    }
    [System.Windows.Forms.SendKeys]::SendWait($keys)
}

function Set-DatePicker([IntPtr]$h, [datetime]$value) {
    Click-Control $h 12                      # 왼쪽 끝 = 연도 칸
    Send-Keys $value.ToString("yyyyMMdd") $script:TargetWindow
    Start-Sleep -Milliseconds 300
}

# 검색은 서버 왕복이라 시간이 들쭉날쭉하다. 카운트 라벨이 멈출 때까지 기다린다.
function Wait-Search([IntPtr]$win, [int]$timeoutSec = 180) {
    $label = Get-Control $win $ID.CountLabel
    $stable = 0
    $prev = $null
    $deadline = (Get-Date).AddSeconds($timeoutSec)

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 1
        $now = [C4]::TextOf($label)
        if ($now -eq $prev -and $now -match "조회고객") { $stable++ } else { $stable = 0 }
        $prev = $now
        if ($stable -ge 3) { break }
    }

    if ($prev -match "조회고객\s*:\s*([\d,]+)") { return [int]($Matches[1] -replace ",", "") }
    return -1
}

# ─────────── 실행 ───────────

$proc = Get-Process CRM4enterprise -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) { Write-Host "CRM4enterprise 가 실행 중이 아닙니다." -ForegroundColor Red; exit 1 }

$win = Find-MainWindow $proc.Id
if ($win -eq [IntPtr]::Zero) {
    Write-Host "통합고객목록 창을 열어주세요. (고객관리 > 통합고객목록)" -ForegroundColor Red; exit 1
}

$n = Update-ControlMap $win
Write-Host "컨트롤 $n 개를 찾았습니다." -ForegroundColor Gray
if ($n -lt 50) { Write-Host "컨트롤이 너무 적습니다. 통합고객목록 창이 맞는지 확인하세요." -ForegroundColor Yellow }

if (-not (Test-Path $OutputRoot)) {
    Write-Host "저장 폴더가 없습니다: $OutputRoot" -ForegroundColor Red
    Write-Host "폴더를 직접 만드시거나, -OutputRoot 로 쓸 경로를 지정하세요." -ForegroundColor Yellow
    Write-Host "예: -OutputRoot ""D:\CRM백업""" -ForegroundColor Yellow
    exit 1
}

if ($DryRun) {
    Write-Host "점검 모드 - 클릭하지 않습니다." -ForegroundColor Cyan
    foreach ($k in $ID.Keys) {
        try {
            $h = Get-Control $win $ID[$k]
            $r = New-Object C4+RECT; [void][C4]::GetWindowRect($h, [ref]$r)
            Write-Host ("  OK   {0,-10} id={1,-8} rect=({2},{3})" -f $k, $ID[$k], $r.Left, $r.Top) -ForegroundColor Green
        } catch {
            Write-Host ("  실패 {0,-10} id={1}" -f $k, $ID[$k]) -ForegroundColor Red
        }
    }
    exit 0
}

# -Start 를 직접 주지 않았으면 이미 받아둔 마지막 날 다음날부터 어제까지를 메운다.
# PC 가 며칠 꺼져 있었더라도 다음에 켜질 때 밀린 날짜가 한 번에 따라잡힌다.
if (-not $PSBoundParameters.ContainsKey('Start')) {
    $done = Get-ChildItem $OutputRoot -Filter "customers_*.xls" -ErrorAction SilentlyContinue |
            ForEach-Object { if ($_.BaseName -match '(\d{4}-\d{2}-\d{2})') { [datetime]$Matches[1] } } |
            Sort-Object -Descending | Select-Object -First 1
    if ($done) {
        $Start = $done.AddDays(1)
        Write-Host "마지막 수집일: $($done.ToString('yyyy-MM-dd')) - 그 다음날부터 이어받습니다." -ForegroundColor Gray
    }
}

if ($Start -gt $End) {
    Write-Host "받을 날짜가 없습니다. 이미 어제까지 다 받았습니다." -ForegroundColor Green
    exit 0
}

Write-Log "수집 시작: $($Start.ToString('yyyy-MM-dd')) ~ $($End.ToString('yyyy-MM-dd'))" "Cyan"
Write-Log "저장 위치: $OutputRoot" "Cyan"

$script:TargetWindow = $win

[void][C4]::SetForegroundWindow($win)
Start-Sleep -Seconds 2

if ([C4]::GetForegroundWindow() -ne $win) {
    Write-Host "통합고객목록 창을 맨 앞으로 가져오지 못했습니다." -ForegroundColor Red
    Write-Host "창을 직접 한 번 클릭해 활성화한 뒤 다시 실행하세요." -ForegroundColor Yellow
    exit 1
}

if (-not $DryRun) {
    Write-Host "5 초 뒤 시작합니다. 마우스와 키보드를 건드리지 마세요. (Ctrl+C 로 취소)" -ForegroundColor Yellow
    Start-Sleep -Seconds 5
}

Switch-Tab $win "Filter"
if (-not $SkipDateCheck) {
    Click-Control (Get-Control $win $ID.DateCheck)
    Write-Log "등록일 조건을 켰습니다. 첫 실행 시 체크 상태를 눈으로 확인하세요." "Yellow"
}

$ok = 0; $empty = 0; $failed = 0
$day = $Start.Date

while ($day -le $End.Date) {
    $stamp = $day.ToString("yyyy-MM-dd")
    $dest = Join-Path $OutputRoot "customers_$stamp.xls"

    if (Test-Path $dest) {
        Write-Log "$stamp 이미 있음, 건너뜀"
        $day = $day.AddDays(1); continue
    }

    try {
        [void][C4]::SetForegroundWindow($win)
        Switch-Tab $win "Filter"
        Set-DatePicker (Get-Control $win $ID.DateFrom) $day
        Set-DatePicker (Get-Control $win $ID.DateTo) $day

        Click-Control (Get-Control $win $ID.Search)
        $count = Wait-Search $win

        if ($count -eq 0) {
            Write-Log "$stamp 0 건, 저장 생략"
            $empty++; $day = $day.AddDays(1); continue
        }
        Write-Log "$stamp 조회 $count 건" "White"

        if (-not $SkipSelectAll) {
            Switch-Tab $win "List"
            Click-Control (Get-Control $win $ID.SelectAll)
        }

        Switch-Tab $win "Process"
        Click-Control (Get-Control $win $ID.Excel)
        Start-Sleep -Seconds $Wait

        # 저장 대화상자(#32770)가 뜨면 경로를 넣고 저장한다.
        $dlg = [IntPtr]::Zero
        $deadline = (Get-Date).AddSeconds(30)
        while ((Get-Date) -lt $deadline -and $dlg -eq [IntPtr]::Zero) {
            $cand = @(Get-TopWindows | Where-Object {
                $_.OwnerPid -eq $proc.Id -and $_.Visible -and $_.Class -eq "#32770"
            })
            if ($cand.Count -gt 0) { $dlg = $cand[0].Handle } else { Start-Sleep -Milliseconds 500 }
        }
        if ($dlg -eq [IntPtr]::Zero) { throw "저장 대화상자가 뜨지 않았습니다." }

        [void][C4]::SetForegroundWindow($dlg)
        Start-Sleep -Milliseconds 500
        Send-Keys $dest.Replace("+","{+}").Replace("^","{^}") $dlg
        Start-Sleep -Milliseconds 300
        Send-Keys "{ENTER}" $dlg
        Start-Sleep -Seconds ($Wait * 2)

        if (Test-Path $dest) {
            Write-Log "$stamp 저장 완료 ($count 건)" "Green"; $ok++
        } else {
            Write-Log "$stamp 저장 확인 실패 - 파일이 없습니다" "Red"; $failed++
        }
    }
    catch {
        Write-Log "$stamp 실패: $($_.Exception.Message)" "Red"
        $failed++
    }

    $day = $day.AddDays(1)
}

Write-Log "끝. 성공 $ok / 0건 $empty / 실패 $failed" "Cyan"
$logPath = Join-Path $OutputRoot ("log_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")
$script:Log | Out-File $logPath -Encoding utf8
Write-Host "로그: $logPath"
