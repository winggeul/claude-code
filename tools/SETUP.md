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
backfill_xlsx.mjs   하루치씩 시트로 나뉜 xlsx 를 세어 넣기 (과거분용)
template.html       화면 틀
store.json          7/1~8/20 집계본 — 이미 채워져 있습니다
```

`store.json` 은 그날그날 뽑아 두신 파일을 센 결과입니다. 건수만 들어 있어 개인정보가 없습니다.
이 파일을 같이 넣어야 8월 21일부터 이어서 쌓입니다. 지우면 과거분이 사라집니다.

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

컨트롤을 제대로 찾는지 먼저 보고 싶으면 클릭 없이 확인만 할 수 있습니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\crm4_collect.ps1 -RawRoot .\raw -Diagnose
```

**실행 중에는 마우스와 키보드를 건드리지 마세요.** 스크립트가 직접 화면을 조작합니다.
5초 카운트다운 동안 `Ctrl+C` 로 취소할 수 있습니다.

끝나면 `site\index.html` 이 생깁니다. 열어서 숫자가 맞는지 확인합니다.

## 4. 과거분 — 할 일 없음

7월 1일 ~ 8월 20일은 `store.json` 에 이미 들어 있습니다. 다시 받지 마세요.

지금 통합고객목록을 다시 뽑으면 그날 뽑았던 상태가 아니라 **오늘 기준으로 갱신된 상태**가 나옵니다.
그 사이에 유입경로가 바뀌거나 지워진 건이 섞이므로, 과거분과 앞으로 쌓일 분의 기준이 어긋납니다.
그래서 과거분은 그날 뽑아 두신 파일을 그대로 센 값만 씁니다.

수집기는 `store.json` 의 마지막 날 다음날부터 채우므로, 8월 21일부터 알아서 이어집니다.

나중에 또 하루치씩 시트로 나뉜 xlsx 를 넣어야 하면:

```powershell
node .\backfill_xlsx.mjs --xlsx .\받은파일.xlsx --store .\store.json
```

`--dry` 를 붙이면 저장하지 않고 어떻게 읽히는지만 보여줍니다.

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

새벽 정해진 시각으로 하려면 트리거 줄만 바꿉니다. 다만 그 시각에 PC가 켜져 있고
CRM4가 떠 있어야 하므로, 로그온 방식이 더 안전합니다.

```powershell
$trigger = New-ScheduledTaskTrigger -Daily -At 3am
```

CRM4 와 `통합고객목록` 창이 떠 있어야 동작하므로, CRM4 를 **시작 프로그램에 등록**하고
로그인 후 통합고객목록 창을 열어 둔 채로 두는 것이 좋습니다.

## 확인과 문제 해결

| 증상 | 확인할 것 |
|---|---|
| 아무 일도 안 일어남 | `C:\CRM자동화` 의 `log_*.txt` |
| 창을 못 찾는다고 나옴 | CRM4 `통합고객목록` 창이 떠 있는지 |
| 클릭 안 하고 멈춤 | 다른 프로그램이 앞에 있었던 것. 정상 동작이며 다시 실행하면 된다 |
| 받아온 날짜가 다르다고 나옴 | 스스로 등록일 체크박스를 뒤집고 다시 받는다. 그대로 두면 된다 |
| 숫자가 안 맞음 | `-KeepRaw` 로 실행해 원본 CSV 를 남긴 뒤 대조 |

원본 CSV 는 `C:\CRM자동화\raw` 에 잠깐 머물다가 집계가 끝나면 지워집니다.
집계가 실패하면 지우지 않으므로 다음 실행에서 다시 시도됩니다.
이 PC 밖으로 나가는 것은 건수 집계뿐입니다.
