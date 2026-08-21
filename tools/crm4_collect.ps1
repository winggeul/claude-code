<#
통합고객목록에서 날짜별로 CSV 를 내려받는다.

등록일을 하루씩 밀며 [검색 → 목록 탭 → 전체선택 → 처리 탭 → Excel → 저장] 을 반복한다.

컨트롤을 GetDlgCtrlID 번호로 찾지 않는다. 이 프로그램의 컨트롤 번호는 창을 새로 열 때마다
달라져서, 한 번 뽑아 둔 번호가 다음 실행에서는 맞지 않는다. 대신 창에 적힌 글자와 클래스,
그리고 서로의 위치 관계로 찾는다. 이것들은 창을 다시 열어도 그대로다.

사용:
    powershell -ExecutionPolicy Bypass -File crm4_collect.ps1 -RawRoot C:\CRM자동화\raw
    powershell -ExecutionPolicy Bypass -File crm4_collect.ps1 -RawRoot ... -Start 2026-07-01 -End 2026-08-21
    powershell -ExecutionPolicy Bypass -File crm4_collect.ps1 -RawRoot ... -Diagnose    클릭 없이 인식만 확인
#>

param(
    # 내려받은 CSV 가 쌓이는 곳. 없는 폴더면 만들지 않고 멈춘다.
    [Parameter(Mandatory = $true)]
    [string]$RawRoot,

    # 수집할 등록일 범위. -Start 를 주지 않으면 마지막으로 받은 날 다음날부터 이어받는다.
    [datetime]$Start,
    [datetime]$End = (Get-Date).Date.AddDays(-1),

    # 클릭을 전혀 하지 않고, 컨트롤을 제대로 찾는지만 확인한다.
    [switch]$Diagnose,

    # 등록일 조건 체크박스를 스크립트가 켜지 않는다. 이미 켜 둔 경우.
    [switch]$SkipDateCheck,

    # 목록 탭의 전체선택을 누르지 않는다.
    [switch]$SkipSelectAll,

    # 단계 사이 기본 대기(초). 서버가 느리면 늘린다.
    [int]$Wait = 2
)

# 탭 순서는 창 구성이라 바뀌지 않는다.
$TabIndex = @{ Filter = 0; List = 4; Process = 5 }

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
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, IntPtr e);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct SYSTEMTIME {
        public ushort wYear, wMonth, wDayOfWeek, wDay, wHour, wMinute, wSecond, wMilliseconds;
    }
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, ref SYSTEMTIME st);

    public static bool SetDate(IntPtr h, int y, int m, int d) {
        SYSTEMTIME st = new SYSTEMTIME();
        st.wYear = (ushort)y; st.wMonth = (ushort)m; st.wDay = (ushort)d;
        return SendMessage(h, 0x1002, (IntPtr)0, ref st) != IntPtr.Zero;   // DTM_SETSYSTEMTIME
    }
    public static string GetDate(IntPtr h) {
        SYSTEMTIME st = new SYSTEMTIME();
        IntPtr r = SendMessage(h, 0x1001, (IntPtr)0, ref st);              // DTM_GETSYSTEMTIME
        if (r.ToInt64() != 0) return "(비어있음)";
        return st.wYear.ToString("D4") + "-" + st.wMonth.ToString("D2") + "-" + st.wDay.ToString("D2");
    }
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

# 델리게이트 콜백 안에서는 함수의 지역 변수가 보이지 않을 수 있다.
# 콜백은 담기만 하고, 판단은 바깥에서 한다.

function Get-TopWindows {
    $script:tops = @()
    $cb = [C4+EnumProc]{
        param($h, $l)
        $o = 0
        [void][C4]::GetWindowThreadProcessId($h, [ref]$o)
        $script:tops += [pscustomobject]@{
            Handle = $h; OwnerPid = [int]$o
            Title = [C4]::TextOf($h); Class = [C4]::ClassOf($h)
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

function Get-Descendants([IntPtr]$root) {
    $script:kids = @()
    $cb = [C4+EnumProc]{
        param($h, $l)
        $r = New-Object C4+RECT
        [void][C4]::GetWindowRect($h, [ref]$r)
        $script:kids += [pscustomobject]@{
            H = $h
            P = [C4]::GetParent($h)
            Class = [C4]::ClassOf($h)
            Text  = [C4]::TextOf($h)
            L = $r.Left; T = $r.Top; R = $r.Right; B = $r.Bottom
            W = $r.Right - $r.Left; Ht = $r.Bottom - $r.Top
            Visible = [C4]::IsWindowVisible($h)
        }
        return $true
    }
    [void][C4]::EnumChildWindows($root, $cb, [IntPtr]::Zero)
    return $script:kids
}

<#
글자와 위치로 컨트롤을 집는다.

  탭          클래스가 SysTabControl32 인 것
  각 탭 페이지 글자가 '고객조건' / '목록' / '처리' 인 컨테이너
  등록일 체크  글자가 '등록일' 인 BUTTON
  날짜 두 칸   SysDateTimePick32 중 등록일 체크와 같은 줄에 있는 것, 왼쪽부터 시작/종료
  검색 버튼    탭 위쪽에 떠 있는 작은 버튼 중 가장 오른쪽
  건수 라벨    '조회고객' 이 적힌 STATIC
  전체선택     글자가 '전체선택' 인 BUTTON
  Excel       '기능' 상자 안의 버튼 두 개 중 오른쪽 (왼쪽은 SMS)
#>
# 탭 페이지들은 화면상 같은 자리에 겹쳐 있다. 위치만으로 고르면 다른 탭의 컨트롤이 걸린다.
# 부모를 거슬러 올라가 정말 그 안에 든 것인지 확인한다.
function Test-Under($ctrl, $ancestor) {
    $p = $ctrl.P
    $guard = 0
    while ($p -ne [IntPtr]::Zero -and $guard -lt 20) {
        if ($p -eq $ancestor.H) { return $true }
        $p = [C4]::GetParent($p)
        $guard++
    }
    return $false
}

function Resolve-Controls([IntPtr]$win) {
    $all = Get-Descendants $win
    $c = @{}

    $tab = $all | Where-Object { $_.Class -match "SysTabControl32" } | Select-Object -First 1
    $c.Tab = $tab

    foreach ($p in @(@("PageFilter","고객조건"), @("PageList","목록"), @("PageProcess","처리"))) {
        $c[$p[0]] = $all | Where-Object { $_.Text -eq $p[1] -and $_.W -gt 400 } | Select-Object -First 1
    }

    $chk = $all | Where-Object { $_.Class -match "BUTTON" -and $_.Text -eq "등록일" } | Select-Object -First 1
    $c.DateCheck = $chk

    if ($chk) {
        $picks = @($all | Where-Object {
            $_.Class -match "SysDateTimePick32" -and [Math]::Abs($_.T - $chk.T) -le 8
        } | Sort-Object L)
        if ($picks.Count -ge 2) { $c.DateFrom = $picks[0]; $c.DateTo = $picks[1] }
    }

    $c.SelectAll = $all | Where-Object { $_.Class -match "BUTTON" -and $_.Text -eq "전체선택" } | Select-Object -First 1

    # 건수 라벨은 검색 전에는 글자가 비어 있을 수 있다. 전체선택과 같은 줄의 오른쪽이라는 위치로 잡는다.
    if ($c.SelectAll) {
        $c.CountLabel = $all | Where-Object {
            $_.Class -match "STATIC" -and $_.L -gt $c.SelectAll.R -and
            [Math]::Abs($_.T - $c.SelectAll.T) -le 12
        } | Sort-Object L | Select-Object -First 1
    }

    if ($tab) {
        $c.Search = $all | Where-Object {
            $_.Visible -and $_.Text -eq "" -and $_.B -le $tab.T -and
            $_.W -ge 40 -and $_.W -le 140 -and $_.Ht -ge 14 -and $_.Ht -le 40
        } | Sort-Object R -Descending | Select-Object -First 1
    }

    # 처리 탭의 '기능' 상자 안에 SMS 와 Excel 이 나란히 있다. 오른쪽이 Excel.
    $grp = $all | Where-Object {
        $_.Text -eq "기능" -and ($null -eq $c.PageProcess -or (Test-Under $_ $c.PageProcess))
    } | Select-Object -First 1

    if ($grp) {
        $btns = @($all | Where-Object {
            $_.H -ne $grp.H -and $_.W -ge 30 -and $_.Ht -ge 14 -and (Test-Under $_ $grp)
        } | Sort-Object L)
        if ($btns.Count -ge 2) { $c.Excel = $btns[1] } elseif ($btns.Count -eq 1) { $c.Excel = $btns[0] }
    }

    $c.Total = $all.Count
    return $c
}

function Need($ctrls, [string]$name) {
    if (-not $ctrls[$name]) { throw "$name 을(를) 찾지 못했습니다." }
    return $ctrls[$name]
}

# 좌표로 누르는 방식이라 대상 창이 맨 앞이 아니면 엉뚱한 창을 클릭하게 된다.
# SetForegroundWindow 는 조용히 실패할 수 있으므로 매번 확인하고, 아니면 누르지 않고 멈춘다.
function Click-Ctrl($ctrl, [int]$offsetX = 0) {
    if ([C4]::GetForegroundWindow() -ne $script:TargetWindow) {
        throw "통합고객목록 창이 맨 앞이 아닙니다. 다른 창을 클릭할 위험이 있어 중단합니다."
    }
    $r = New-Object C4+RECT
    [void][C4]::GetWindowRect($ctrl.H, [ref]$r)      # 창을 옮겼을 수 있으니 지금 좌표를 다시 읽는다
    $x = if ($offsetX -gt 0) { $r.Left + $offsetX } else { [int](($r.Left + $r.Right) / 2) }
    $y = [int](($r.Top + $r.Bottom) / 2)
    [void][C4]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 150
    [C4]::mouse_event(0x02, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [C4]::mouse_event(0x04, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds 400
}

function Send-Keys([string]$keys, [IntPtr]$expect) {
    if ([C4]::GetForegroundWindow() -ne $expect) {
        throw "입력 대상 창이 맨 앞이 아닙니다. 엉뚱한 곳에 입력될 위험이 있어 중단합니다."
    }
    [System.Windows.Forms.SendKeys]::SendWait($keys)
}

# 탭을 바꾸면 그 페이지에서만 만들어지는 컨트롤이 있으므로 매번 다시 찾는다.
function Switch-Tab([IntPtr]$win, [string]$name) {
    $c = Resolve-Controls $win
    $tab = Need $c "Tab"
    [void][C4]::SendMessage($tab.H, 0x1330, [IntPtr]$TabIndex[$name], [IntPtr]::Zero)   # TCM_SETCURFOCUS
    Start-Sleep -Milliseconds 700

    $c = Resolve-Controls $win
    $page = Need $c "Page$name"
    if (-not [C4]::IsWindowVisible($page.H)) { throw "$name 탭으로 전환하지 못했습니다." }
    return $c
}

# 메시지로 넣어 보고, 되읽어서 확인하고, 안 들어갔으면 타이핑으로 다시 시도한다.
# 넣었다고 믿고 넘어가면 엉뚱한 날짜로 검색해도 알 수가 없다.
function Set-DatePicker($ctrl, [datetime]$value, [string]$label) {
    $want = $value.ToString("yyyy-MM-dd")

    [void][C4]::SetDate($ctrl.H, $value.Year, $value.Month, $value.Day)
    Start-Sleep -Milliseconds 250
    $got = [C4]::GetDate($ctrl.H)

    if ($got -ne $want) {
        Click-Ctrl $ctrl 12                    # 왼쪽 끝 = 연도 칸
        Send-Keys $value.ToString("yyyyMMdd") $script:TargetWindow
        Start-Sleep -Milliseconds 400
        $got = [C4]::GetDate($ctrl.H)
    }

    if ($got -ne $want) { throw "$label 날짜가 $want 로 안 들어갔습니다. 현재 값: $got" }
    Write-Log "  $label = $got"
}

# WinForms 체크박스는 상태를 못 읽는 경우가 있다. 읽히면 그 값을 쓰고, 안 읽히면 로그로 남긴다.
function Get-CheckState($ctrl) {
    return [C4]::SendMessage($ctrl.H, 0x00F0, [IntPtr]::Zero, [IntPtr]::Zero).ToInt32()   # BM_GETCHECK
}

# 검색은 서버 왕복이라 걸리는 시간이 일정하지 않다. 고정 대기 대신 건수 라벨이 멈출 때까지 본다.
function Wait-Search([IntPtr]$win, [int]$timeoutSec = 180) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    $stable = 0; $prev = $null

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 1
        $c = Resolve-Controls $win
        $now = if ($c.CountLabel) { [C4]::TextOf($c.CountLabel.H) } else { "" }
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
$script:TargetWindow = $win

if (-not (Test-Path $RawRoot)) {
    Write-Host "원본 저장 폴더가 없습니다: $RawRoot" -ForegroundColor Red
    Write-Host "폴더를 직접 만드시거나, -RawRoot 로 쓸 경로를 지정하세요." -ForegroundColor Yellow
    exit 1
}

# 탭을 한 번도 열지 않으면 그 페이지의 컨트롤은 아직 만들어지지 않는다.
# 점검할 때는 탭을 차례로 넘겨보며 모은다. 탭 전환은 메시지로 하므로 마우스는 움직이지 않는다.
$ctrls = Resolve-Controls $win
$tab0 = $ctrls.Tab
if (-not $tab0) { Write-Host "탭 컨트롤을 찾지 못했습니다. 통합고객목록 창이 맞는지 확인하세요." -ForegroundColor Red; exit 1 }

foreach ($name in @("List","Process","Filter")) {
    [void][C4]::SendMessage($tab0.H, 0x1330, [IntPtr]$TabIndex[$name], [IntPtr]::Zero)
    Start-Sleep -Milliseconds 600
    $found = Resolve-Controls $win
    foreach ($k in @($found.Keys)) { if ($k -ne "Total" -and $found[$k] -and -not $ctrls[$k]) { $ctrls[$k] = $found[$k] } }
}

Write-Host "찾은 결과:" -ForegroundColor Gray
foreach ($k in @("Tab","PageFilter","PageList","PageProcess","DateCheck","DateFrom","DateTo","Search","CountLabel","SelectAll","Excel")) {
    if ($ctrls[$k]) {
        Write-Host ("  OK   {0,-12} ({1},{2}) {3}" -f $k, $ctrls[$k].L, $ctrls[$k].T, $ctrls[$k].Class.Split('.')[1]) -ForegroundColor Green
    } else {
        Write-Host ("  실패 {0,-12}" -f $k) -ForegroundColor Red
    }
}

if ($Diagnose) { Write-Host "`n점검 모드였습니다. 클릭하지 않았습니다." -ForegroundColor Cyan; exit 0 }

foreach ($k in @("Tab","PageFilter","PageList","PageProcess","DateCheck","DateFrom","DateTo","Search","CountLabel","Excel")) {
    if (-not $ctrls[$k]) { Write-Host "`n$k 을(를) 찾지 못해 중단합니다." -ForegroundColor Red; exit 1 }
}

if (-not $PSBoundParameters.ContainsKey('Start')) {
    $done = Get-ChildItem $RawRoot -Filter "customers_*.csv" -ErrorAction SilentlyContinue |
            ForEach-Object { if ($_.BaseName -match '(\d{4}-\d{2}-\d{2})') { [datetime]$Matches[1] } } |
            Sort-Object -Descending | Select-Object -First 1
    if ($done) {
        $Start = $done.AddDays(1)
        Write-Host "마지막 수집일: $($done.ToString('yyyy-MM-dd')) - 그 다음날부터 이어받습니다." -ForegroundColor Gray
    } else {
        $Start = $End
    }
}
if ($Start -gt $End) { Write-Host "받을 날짜가 없습니다." -ForegroundColor Green; exit 0 }

Write-Log "수집 시작: $($Start.ToString('yyyy-MM-dd')) ~ $($End.ToString('yyyy-MM-dd'))" "Cyan"
Write-Log "저장 위치: $RawRoot" "Cyan"

[void][C4]::SetForegroundWindow($win)
Start-Sleep -Seconds 2
if ([C4]::GetForegroundWindow() -ne $win) {
    Write-Host "통합고객목록 창을 맨 앞으로 가져오지 못했습니다." -ForegroundColor Red
    Write-Host "창을 직접 한 번 클릭해 활성화한 뒤 다시 실행하세요." -ForegroundColor Yellow
    exit 1
}

Write-Host "5 초 뒤 시작합니다. 마우스와 키보드를 건드리지 마세요. (Ctrl+C 로 취소)" -ForegroundColor Yellow
Start-Sleep -Seconds 5

$ctrls = Switch-Tab $win "Filter"
if (-not $SkipDateCheck) {
    $chk = Need $ctrls "DateCheck"
    $before = Get-CheckState $chk
    if ($before -ne 1) {
        Click-Ctrl $chk
        $after = Get-CheckState $chk
        Write-Log "등록일 조건 체크: $before -> $after" "Yellow"
    } else {
        Write-Log "등록일 조건이 이미 켜져 있습니다." "Yellow"
    }
}

$ok = 0; $empty = 0; $failed = 0
$day = $Start.Date

while ($day -le $End.Date) {
    $stamp = $day.ToString("yyyy-MM-dd")
    $dest = Join-Path $RawRoot "customers_$stamp.csv"

    if (Test-Path $dest) {
        Write-Log "$stamp 이미 있음, 건너뜀"
        $day = $day.AddDays(1); continue
    }

    try {
        [void][C4]::SetForegroundWindow($win)
        Start-Sleep -Milliseconds 400

        $c = Switch-Tab $win "Filter"
        Set-DatePicker (Need $c "DateFrom") $day "시작일"
        Set-DatePicker (Need $c "DateTo") $day "종료일"

        Click-Ctrl (Need $c "Search")
        $count = Wait-Search $win

        if ($count -eq 0) {
            Write-Log "$stamp 0 건, 저장 생략"
            $empty++; $day = $day.AddDays(1); continue
        }
        Write-Log "$stamp 조회 $count 건" "White"

        if (-not $SkipSelectAll) {
            $c = Switch-Tab $win "List"
            Click-Ctrl (Need $c "SelectAll")
        }

        $c = Switch-Tab $win "Process"
        $started = Get-Date
        Click-Ctrl (Need $c "Excel")
        Start-Sleep -Seconds $Wait

        # 저장 대화상자(#32770)를 기다린다.
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

        # 프로그램이 이름이나 확장자를 제 방식대로 붙일 수 있다. 없으면 방금 생긴 파일을 찾아 맞춘다.
        if (-not (Test-Path $dest)) {
            $fresh = Get-ChildItem $RawRoot -File -ErrorAction SilentlyContinue |
                     Where-Object { $_.LastWriteTime -gt $started } |
                     Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($fresh) {
                Move-Item $fresh.FullName $dest -Force
                Write-Log "$stamp 파일명 정리: $($fresh.Name) -> $(Split-Path $dest -Leaf)"
            }
        }

        # 저장이 끝나면 '엑셀 다운로드 성공' 안내창이 뜬다. 닫지 않으면 다음 날짜로 못 넘어간다.
        $notice = @(Get-TopWindows | Where-Object {
            $_.OwnerPid -eq $proc.Id -and $_.Visible -and $_.Handle -ne $win -and
            $_.Class -ne "#32770" -and $_.Title -notlike "*통합고객목록*"
        })
        foreach ($n in @(Get-TopWindows | Where-Object { $_.OwnerPid -eq $proc.Id -and $_.Visible -and $_.Class -eq "#32770" })) {
            [void][C4]::SetForegroundWindow($n.Handle)
            Start-Sleep -Milliseconds 300
            Send-Keys "{ENTER}" $n.Handle
            Start-Sleep -Milliseconds 400
        }
        if ($notice.Count -gt 0) {
            [void][C4]::SetForegroundWindow($notice[0].Handle)
            Start-Sleep -Milliseconds 300
            Send-Keys "{ENTER}" $notice[0].Handle
            Start-Sleep -Milliseconds 400
            Write-Log "안내창을 닫았습니다."
        }

        if (Test-Path $dest) { Write-Log "$stamp 저장 완료 ($count 건)" "Green"; $ok++ }
        else { Write-Log "$stamp 저장 확인 실패 - 파일이 없습니다" "Red"; $failed++ }
    }
    catch {
        Write-Log "$stamp 실패: $($_.Exception.Message)" "Red"
        $failed++
    }

    $day = $day.AddDays(1)
}

Write-Log "끝. 성공 $ok / 0건 $empty / 실패 $failed" "Cyan"
$logPath = Join-Path $RawRoot ("log_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".txt")
$script:Log | Out-File $logPath -Encoding utf8
Write-Host "로그: $logPath"
