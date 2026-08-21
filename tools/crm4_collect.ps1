# 통합고객목록 자동 수집기
#
# 등록일을 하루씩 밀며 검색 → 처리 탭 → Excel → 지정 경로에 저장 을 반복한다.
# 74 만 건을 한 번에 받지 않고 날짜로 잘라 받는 이유는 화면이 한 번에 그만큼을
# 감당하지 못하기 때문이다.
#
# ※ 아래 $Ids 값은 crm4_probe.ps1 결과를 보고 채워야 한다. 0 인 채로는 동작하지 않는다.

# ─────────── 설정 ───────────

# 저장 위치. 로컬 PC 가 아닌 곳을 지정할 것 (네트워크 드라이브 / NAS / 사내 서버).
$OutputRoot = "\\NAS\crm4"

# 수집할 등록일 범위
$StartDate = Get-Date "2026-08-01"
$EndDate   = Get-Date "2026-08-21"

# 컨트롤 식별자 — crm4_probe.ps1 덤프에서 확인한 값으로 교체
$Ids = @{
    등록일체크 = 0   # 고객조건 탭의 '등록일' 체크박스
    시작일     = 0   # 등록일 시작 날짜 입력칸
    종료일     = 0   # 등록일 종료 날짜 입력칸
    검색       = 0   # 우상단 검색 버튼
    탭컨트롤   = 0   # 고객조건/목록/처리 탭
    엑셀       = 0   # 처리 탭의 Excel 버튼
}

# 각 단계 사이 대기 시간(초). 서버 응답이 느리면 늘린다.
$Wait = 2

# ─────────── 내부 ───────────

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Crm4 {
    public delegate bool EnumProc(IntPtr h, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    public static string TextOf(IntPtr h) {
        StringBuilder sb = new StringBuilder(512); GetWindowText(h, sb, 512); return sb.ToString();
    }
}
"@

function Find-Window([string]$titlePart, [int]$processId) {
    $found = [IntPtr]::Zero
    $cb = [Crm4+EnumProc]{
        param($h, $l)
        $owner = 0
        [void][Crm4]::GetWindowThreadProcessId($h, [ref]$owner)
        if ($owner -eq $processId -and [Crm4]::IsWindowVisible($h) -and [Crm4]::TextOf($h) -like "*$titlePart*") {
            $script:found = $h; return $false
        }
        return $true
    }
    [void][Crm4]::EnumWindows($cb, [IntPtr]::Zero)
    return $script:found
}

function Find-Control([IntPtr]$parent, [int]$controlId) {
    $script:hit = [IntPtr]::Zero
    $cb = [Crm4+EnumProc]{
        param($h, $l)
        if ([Crm4]::GetDlgCtrlID($h) -eq $controlId) { $script:hit = $h; return $false }
        return $true
    }
    [void][Crm4]::EnumChildWindows($parent, $cb, [IntPtr]::Zero)
    return $script:hit
}

function Invoke-Click([IntPtr]$h) {
    [void][Crm4]::SendMessage($h, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)  # BM_CLICK
    Start-Sleep -Seconds $Wait
}

# ─────────── 실행 ───────────

if ($Ids.Values -contains 0) {
    Write-Host "먼저 crm4_probe.ps1 을 돌려 `$Ids 값을 채우세요." -ForegroundColor Yellow
    return
}

$proc = Get-Process CRM4enterprise -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) { Write-Host "CRM4enterprise 가 실행 중이 아닙니다." -ForegroundColor Red; return }

if (-not (Test-Path $OutputRoot)) {
    Write-Host "저장 경로에 접근할 수 없습니다: $OutputRoot" -ForegroundColor Red; return
}

$win = Find-Window "통합고객목록" $proc.Id
if ($win -eq [IntPtr]::Zero) { Write-Host "통합고객목록 창을 열어주세요." -ForegroundColor Red; return }

[void][Crm4]::SetForegroundWindow($win)

$day = $StartDate
while ($day -le $EndDate) {
    $stamp = $day.ToString("yyyy-MM-dd")
    $dest = Join-Path $OutputRoot "customers_$stamp.xls"

    if (Test-Path $dest) {
        Write-Host "$stamp 이미 있음, 건너뜀"
        $day = $day.AddDays(1); continue
    }

    Write-Host "$stamp 수집 중..." -ForegroundColor Cyan

    # 등록일 범위를 하루로 좁히고 검색
    Invoke-Click (Find-Control $win $Ids.등록일체크)
    # TODO: 날짜 입력칸 설정 (DTM_SETSYSTEMTIME) — probe 결과의 클래스 확인 후 확정
    Invoke-Click (Find-Control $win $Ids.검색)
    Start-Sleep -Seconds ($Wait * 3)

    # 처리 탭 → Excel → 저장 대화상자에 경로 입력
    Invoke-Click (Find-Control $win $Ids.엑셀)
    Start-Sleep -Seconds $Wait
    [System.Windows.Forms.SendKeys]::SendWait($dest + "{ENTER}")
    Start-Sleep -Seconds ($Wait * 2)

    $day = $day.AddDays(1)
}

Write-Host "완료. 저장 위치: $OutputRoot" -ForegroundColor Green
