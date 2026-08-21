# CRM4enterprise 화면 구조 덤프 (윈도우 기본 기능만 사용, 설치 불필요)
#
# 사용법: CRM4enterprise 를 실행하고 고객목록 화면을 연 상태에서
#         PowerShell 에 이 파일 내용을 통째로 붙여넣는다.
#         바탕화면에 crm4_dump.txt 가 생성된다.

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$proc = Get-Process CRM4enterprise -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) {
    Write-Host "CRM4enterprise 가 실행 중이 아닙니다. 프로그램을 먼저 켜주세요." -ForegroundColor Red
    return
}

$AE = [System.Windows.Automation.AutomationElement]
$cond = New-Object System.Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $proc.Id)
$wins = $AE::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)

Write-Host "창 $($wins.Count) 개를 찾았습니다. 읽는 중..." -ForegroundColor Cyan

$out = New-Object System.Collections.Generic.List[string]
$out.Add("# CRM4enterprise 화면 덤프 - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$gridCount = 0

foreach ($w in $wins) {
    $out.Add("")
    $out.Add("===== 창: $($w.Current.Name) =====")

    $all = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                      [System.Windows.Automation.Condition]::TrueCondition)

    foreach ($e in $all) {
        $c = $e.Current
        $out.Add(("{0} | name={1} | id={2} | class={3}" -f `
            $c.ControlType.ProgrammaticName, $c.Name, $c.AutomationId, $c.ClassName))

        # 표(Grid) 컨트롤이면 크기와 앞 3 행을 미리 뽑는다. 여기가 데이터가 든 곳이다.
        $gp = $null
        if ($e.TryGetCurrentPattern([System.Windows.Automation.GridPattern]::Pattern, [ref]$gp)) {
            $rows = $gp.Current.RowCount
            $cols = $gp.Current.ColumnCount
            $gridCount++
            $out.Add("    >>> GRID 발견: $rows 행 x $cols 열")
            for ($r = 0; $r -lt [Math]::Min($rows, 3); $r++) {
                $cells = @()
                for ($k = 0; $k -lt $cols; $k++) {
                    try { $cells += $gp.GetItem($r, $k).Current.Name } catch { $cells += "?" }
                }
                $out.Add("    >>> row $r : " + ($cells -join " | "))
            }
        }
    }
}

$path = Join-Path ([Environment]::GetFolderPath("Desktop")) "crm4_dump.txt"
$out | Out-File -FilePath $path -Encoding utf8

Write-Host ""
Write-Host "완료. 바탕화면에 crm4_dump.txt 저장됨" -ForegroundColor Green
Write-Host "표(GRID) $gridCount 개 발견, 총 $($out.Count) 줄" -ForegroundColor Green
