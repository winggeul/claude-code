# 매크로 PC 설치 안내

한 번만 하면 됩니다. 그 뒤로는 매일 새벽에 알아서 돕니다.

## 1. 폴더 만들기

`C:\CRM자동화` 를 만들고 이 파일들을 넣습니다.
C 드라이브 최상위에 폴더를 만들 때 권한을 물으면 계속 진행을 누르면 됩니다.
막히면 `C:\Users\사용자이름\CRM자동화` 도 괜찮습니다. 스크립트는 자기가 놓인 폴더를 기준으로 동작합니다.

```
crm4_collect.ps1    CRM4에서 CSV 내려받기
crm4_daily.ps1      매일 도는 전체 작업 (스케줄러가 실행할 파일)
aggregate.mjs       CSV 세어서 화면 만들기
template.html       화면 틀
```

## 2. Node.js 설치

<https://nodejs.org> 에서 LTS 버전을 받아 설치합니다. 기본 설정 그대로 다음만 누르면 됩니다.

설치 후 PowerShell을 새로 열고 확인합니다.

```powershell
node -v
```

`v22.x.x` 같은 게 나오면 됩니다.

## 3. 첫 실행 — 배포 없이 하루치만

CRM4를 켜고 `고객관리 > 통합고객목록` 창을 띄운 다음:

```powershell
cd C:\CRM자동화
powershell -ExecutionPolicy Bypass -File .\crm4_daily.ps1 -SkipDeploy
```

**실행 중에는 마우스와 키보드를 건드리지 마세요.** 스크립트가 직접 화면을 조작합니다.
5초 카운트다운 동안 `Ctrl+C` 로 취소할 수 있습니다.

끝나면 `site\index.html` 이 생깁니다. 열어서 숫자가 맞는지 확인합니다.

## 4. 과거분 채우기

7월 1일부터 한 번에 받습니다. 날짜 수만큼 반복하므로 시간이 걸립니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\crm4_daily.ps1 -Start 2026-07-01 -SkipDeploy
```

이미 받은 날짜는 건너뛰므로, 중간에 끊겨도 다시 실행하면 이어집니다.

## 5. 버셀 연결

<https://vercel.com> 가입 후 한 번만 수동으로 올립니다.

```powershell
npx vercel login
npx vercel deploy .\site --prod
```

프로젝트 이름을 물으면 아무거나 정하면 됩니다. 배포가 끝나면 주소가 나옵니다.
그 주소는 앞으로 바뀌지 않습니다.

그다음 <https://vercel.com/account/tokens> 에서 토큰을 발급받아 등록합니다.

```powershell
[Environment]::SetEnvironmentVariable("VERCEL_TOKEN", "발급받은토큰", "Machine")
```

PowerShell을 새로 열고 전체 실행이 되는지 확인합니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\crm4_daily.ps1
```

## 6. 자동 실행 등록

**관리자 권한 PowerShell**에서 실행합니다.

```powershell
$action  = New-ScheduledTaskAction -Execute "powershell.exe" `
           -Argument "-ExecutionPolicy Bypass -File C:\CRM자동화\crm4_daily.ps1" `
           -WorkingDirectory "C:\CRM자동화"
$trigger = New-ScheduledTaskTrigger -AtLogOn
$set     = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 3)

Register-ScheduledTask -TaskName "CRM4 일일수집" -Action $action -Trigger $trigger `
                       -Settings $set -RunLevel Highest -Force
```

로그온할 때마다 실행됩니다. 며칠 꺼져 있었어도 켜지는 순간 밀린 날짜를 메웁니다.

CRM4가 먼저 떠 있어야 하므로, CRM4를 **시작 프로그램에 등록**해 두는 편이 좋습니다.

## 확인과 문제 해결

| 증상 | 확인할 것 |
|---|---|
| 아무 일도 안 일어남 | `C:\CRM자동화` 의 `log_*.txt` |
| 창을 못 찾는다고 나옴 | CRM4 `통합고객목록` 창이 떠 있는지 |
| 클릭 안 하고 멈춤 | 다른 창이 앞에 있었던 것. 정상 동작이며 다시 실행하면 된다 |
| 숫자가 안 맞음 | `-KeepRaw` 로 실행해 원본 CSV 를 남긴 뒤 대조 |

원본 CSV 는 `C:\CRM자동화\raw` 에 잠깐 머물다가 집계가 끝나면 지워집니다.
집계가 실패하면 지우지 않으므로 다음 실행에서 다시 시도됩니다.
이 PC 밖으로 나가는 것은 건수 집계뿐입니다.
