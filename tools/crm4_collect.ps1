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

    # 저장 위치. Synology Drive 동기화 폴더에 쓰면 NAS 로 자동 업로드된다.
    [string]$OutputRoot = "C:\SynologyDrive\CRM4",

    # 등록일 조건 체크박스를 스크립트가 켤지 여부. 이미 켜 둔 상태라면 -SkipDateCheck 로 끈다.
    [switch]$SkipDateCheck,

    # 처리 탭의 Excel 이 '선택한 고객' 기준으로 동작하는 경우를 대비해 전체선택을 먼저 누른다.
    [switch]$SkipSelectAll,

    # 각 단계 사이 기본 대기(초). 서버가 느리면 늘린다.
    [int]$Wait = 2
)

# ─────────── 컨트롤 식별자 (probe 덤프 실측값) ───────────

$ID = @{
    Tab        = 16       # SysTabControl32
    Page고객조건 = 527392
    Page목록     = 199984
    Page처리     = 592892
    등록일체크    = 68906   # '등록일' 체크박스
    등록일시작    = 68896   # SysDateTimePick32
    등록일종료    = 68902   # SysDateTimePick32
    검색         = 724064
    카운트라벨    = 199986  # '전체고객 : ... 조회고객 : N 명'
    전체선택      = 396598
    엑셀         = 68956   # 처리 탭 기능 그룹의 Excel 버튼
}

$TabIndex = @{ 고객조건 = 0; 목록 = 4; 처리 = 5 }

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

function Find-MainWindow([int]$processId) {
    $script:hit = [IntPtr]::Zero
    $cb = [C4+EnumProc]{
        param($h, $l)
        $o = 0
        [void][C4]::GetWindowThreadProcessId($h, [ref]$o)
        if ($o -eq $processId -and [C4]::IsWindowVisible($h) -and [C4]::TextOf($h) -like "*통합고객목록*") {
            $script:hit = $h; return $false
        }
        return $true
    }
    [void][C4]::EnumWindows($cb, [IntPtr]::Zero)
    return $script:hit
}

function Get-Control([IntPtr]$root, [int]$controlId) {
    $script:found = [IntPtr]::Zero
    $cb = [C4+EnumProc]{
        param($h, $l)
        if ([C4]::GetDlgCtrlID($h) -eq $controlId) { $script:found = $h; return $false }
        return $true
    }
    [void][C4]::EnumChildWindows($root, $cb, [IntPtr]::Zero)
    if ($script:found -eq [IntPtr]::Zero) { throw "컨트롤 $controlId 을 찾지 못했습니다." }
    return $script:found
}

# 이 프로그램의 버튼은 대부분 직접 그린 컨트롤이라 BM_CLICK 이 통하지 않는다.
# 실행 시점의 좌표를 다시 읽어 실제 마우스로 누른다. 창을 옮겨도 스스로 맞춰진다.
function Click-Control([IntPtr]$h, [int]$offsetX = 0) {
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

    $page = Get-Control $win $ID."Page$name"
    if (-not [C4]::IsWindowVisible($page)) { throw "'$name' 탭으로 전환하지 못했습니다." }
}

function Set-DatePicker([IntPtr]$h, [datetime]$value) {
    Click-Control $h 12                      # 왼쪽 끝 = 연도 칸
    [System.Windows.Forms.SendKeys]::SendWait($value.ToString("yyyyMMdd"))
    Start-Sleep -Milliseconds 300
}

# 검색은 서버 왕복이라 시간이 들쭉날쭉하다. 카운트 라벨이 멈출 때까지 기다린다.
function Wait-Search([IntPtr]$win, [int]$timeoutSec = 180) {
    $label = Get-Control $win $ID.카운트라벨
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

if (-not (Test-Path $OutputRoot)) { New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null }

Write-Log "수집 시작: $($Start.ToString('yyyy-MM-dd')) ~ $($End.ToString('yyyy-MM-dd'))" "Cyan"
Write-Log "저장 위치: $OutputRoot" "Cyan"

[void][C4]::SetForegroundWindow($win)
Start-Sleep -Seconds 1

Switch-Tab $win "고객조건"
if (-not $SkipDateCheck) {
    Click-Control (Get-Control $win $ID.등록일체크)
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
        Switch-Tab $win "고객조건"
        Set-DatePicker (Get-Control $win $ID.등록일시작) $day
        Set-DatePicker (Get-Control $win $ID.등록일종료) $day

        Click-Control (Get-Control $win $ID.검색)
        $count = Wait-Search $win

        if ($count -eq 0) {
            Write-Log "$stamp 0 건, 저장 생략"
            $empty++; $day = $day.AddDays(1); continue
        }
        Write-Log "$stamp 조회 $count 건" "White"

        if (-not $SkipSelectAll) {
            Switch-Tab $win "목록"
            Click-Control (Get-Control $win $ID.전체선택)
        }

        Switch-Tab $win "처리"
        Click-Control (Get-Control $win $ID.엑셀)
        Start-Sleep -Seconds $Wait

        # 저장 대화상자(#32770)가 뜨면 경로를 넣고 저장한다.
        $dlg = [IntPtr]::Zero
        $deadline = (Get-Date).AddSeconds(30)
        while ((Get-Date) -lt $deadline -and $dlg -eq [IntPtr]::Zero) {
            $script:hit = [IntPtr]::Zero
            $cb = [C4+EnumProc]{
                param($h, $l)
                $o = 0
                [void][C4]::GetWindowThreadProcessId($h, [ref]$o)
                if ($o -eq $proc.Id -and [C4]::IsWindowVisible($h) -and [C4]::ClassOf($h) -eq "#32770") {
                    $script:hit = $h; return $false
                }
                return $true
            }
            [void][C4]::EnumWindows($cb, [IntPtr]::Zero)
            $dlg = $script:hit
            if ($dlg -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 500 }
        }
        if ($dlg -eq [IntPtr]::Zero) { throw "저장 대화상자가 뜨지 않았습니다." }

        [void][C4]::SetForegroundWindow($dlg)
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait($dest.Replace("+","{+}").Replace("^","{^}"))
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
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
