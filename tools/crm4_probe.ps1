# 통합고객목록 창의 Win32 컨트롤 조사 (설치 불필요)
#
# UIA 로는 모든 컨트롤이 밋밋한 Pane 으로만 보였지만, 클래스 이름이
# WindowsForms10.EDIT / .BUTTON / .COMBOBOX 인 것에서 보이듯 실체는 진짜 Win32
# 자식 윈도우다. 그래서 여기서는 UIA 대신 user32.dll 을 직접 호출해 훑는다.
#
# 사용법: 통합고객목록 창을 띄운 상태에서 이 내용을 PowerShell 에 붙여넣는다.
#         탭을 바꿔가며 여러 번 실행하면 탭별 컨트롤이 각각 파일로 남는다.

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Win32Probe {
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }

    public static string ClassOf(IntPtr h) {
        StringBuilder sb = new StringBuilder(256); GetClassName(h, sb, 256); return sb.ToString();
    }
    public static string TextOf(IntPtr h) {
        StringBuilder sb = new StringBuilder(512); GetWindowText(h, sb, 512); return sb.ToString();
    }
}
"@

$proc = Get-Process CRM4enterprise -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) {
    Write-Host "CRM4enterprise 가 실행 중이 아닙니다." -ForegroundColor Red
    return
}

# 이 프로세스가 소유한 최상위 창을 모은다.
$tops = New-Object System.Collections.Generic.List[IntPtr]
$findTop = [Win32Probe+EnumProc]{
    param($h, $l)
    $owner = 0
    [void][Win32Probe]::GetWindowThreadProcessId($h, [ref]$owner)
    if ($owner -eq $proc.Id -and [Win32Probe]::IsWindowVisible($h)) { $tops.Add($h) }
    return $true
}
[void][Win32Probe]::EnumWindows($findTop, [IntPtr]::Zero)

$targets = @()
foreach ($h in $tops) {
    if ([Win32Probe]::TextOf($h) -like "*통합고객목록*") { $targets += $h }
}
if ($targets.Count -eq 0) {
    Write-Host "통합고객목록 창을 찾지 못했습니다. 열려 있는 창 목록:" -ForegroundColor Yellow
    foreach ($h in $tops) { Write-Host ("  - " + [Win32Probe]::TextOf($h)) }
    return
}

$out = New-Object System.Collections.Generic.List[string]
$out.Add("# 통합고객목록 Win32 덤프 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")

foreach ($top in $targets) {
    $out.Add("")
    $out.Add("===== 창: $([Win32Probe]::TextOf($top))  handle=$top =====")

    $collect = [Win32Probe+EnumProc]{
        param($h, $l)
        $rect = New-Object Win32Probe+RECT
        [void][Win32Probe]::GetWindowRect($h, [ref]$rect)

        # 부모를 거슬러 올라가 창 안에서의 깊이를 구한다.
        $depth = 0; $p = [Win32Probe]::GetParent($h)
        while ($p -ne [IntPtr]::Zero -and $p -ne $top -and $depth -lt 20) {
            $p = [Win32Probe]::GetParent($p); $depth++
        }

        $out.Add(("{0}handle={1} | id={2} | class={3} | text='{4}' | vis={5} | en={6} | rect=({7},{8},{9},{10})" -f `
            ("  " * $depth), $h, [Win32Probe]::GetDlgCtrlID($h), [Win32Probe]::ClassOf($h),
            [Win32Probe]::TextOf($h), [Win32Probe]::IsWindowVisible($h), [Win32Probe]::IsWindowEnabled($h),
            $rect.Left, $rect.Top, $rect.Right, $rect.Bottom))
        return $true
    }
    [void][Win32Probe]::EnumChildWindows($top, $collect, [IntPtr]::Zero)
}

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$path = Join-Path ([Environment]::GetFolderPath("Desktop")) "crm4_probe_$stamp.txt"
$out | Out-File -FilePath $path -Encoding utf8

Write-Host ""
Write-Host "완료. 바탕화면에 crm4_probe_$stamp.txt 저장됨 ($($out.Count) 줄)" -ForegroundColor Green
