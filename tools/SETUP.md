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
roster.mjs          담당자 → 시흥·천안 명단
template.html       화면 틀
api\headcount.js    인원을 저장하는 자리
package.json        위 파일이 쓰는 꾸러미 목록
store.json          7/1~8/26 집계본 — 이미 채워져 있습니다
```

`store.json` 은 그날그날 뽑아 두신 파일을 센 결과입니다. 건수만 들어 있어 개인정보가 없습니다.
이 파일을 같이 넣어야 8월 27일부터 이어서 쌓입니다. 지우면 과거분이 사라집니다.

## 1-1. 상담사 인원 바꾸기

화면의 시흥·천안 카드에서 인원 숫자를 눌러 고칩니다. **어느 PC 에서 고쳐도 모든 PC 에
반영됩니다.** 다른 PC 는 1분 안에 따라오고, 창을 다시 보는 순간에도 한 번 더 확인합니다.

이게 되려면 버셀에 값을 둘 자리를 한 번 만들어야 합니다. **5-1 을 보세요.**
만들기 전까지는 눌러 고쳐도 그 화면에서만 바뀌고, 표 위에 저장하지 못했다고 뜹니다.

인원을 바꾸지 않고 "10명이면 어떻게 되지" 만 보고 싶으면 주소 뒤에 붙입니다.
이건 저장하지 않으므로 그 화면에서만 그렇고, 표 위의 `되돌리기` 를 누르면 돌아옵니다.

```
https://.../?시흥=6&천안=11
```

담당자가 늘거나 옮기면 `roster.mjs` 의 명단에 이름을 넣습니다. 명단에 없는 담당자의 건은
시흥 5 대 천안 11 비율로 나눠 담기므로 합계는 언제나 맞습니다.

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

7월 1일 ~ 8월 26일은 `store.json` 에 이미 들어 있습니다. 다시 받지 마세요.

지금 통합고객목록을 다시 뽑으면 그날 뽑았던 상태가 아니라 **오늘 기준으로 갱신된 상태**가 나옵니다.
그 사이에 유입경로가 바뀌거나 지워진 건이 섞이므로, 과거분과 앞으로 쌓일 분의 기준이 어긋납니다.
그래서 과거분은 그날 뽑아 두신 파일을 그대로 센 값만 씁니다.

수집기는 **어제 하루만** 받습니다. 빠진 날을 저절로 메우지 않습니다.

같은 이유에서입니다. 어제치는 오늘 뽑아야 그날 상태 그대로이고, 그저께치를 오늘 뽑으면
이미 갱신된 값이 나와 다른 날과 기준이 어긋납니다. 그래서 못 받은 날은 그냥 비워 둡니다.

대시보드는 그런 날을 **미수집**으로 한 줄 세우고 위쪽에 `미수집 N일` 이라고 적습니다.
숫자가 0으로 떨어진 것처럼 보이는 일은 없습니다.

기준이 달라져도 좋으니 굳이 받아야 한다면 `-Backfill` 을 함께 줘야 합니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\crm4_daily.ps1 -Start 2026-08-25 -Backfill
```

나중에 또 하루치씩 시트로 나뉜 xlsx 를 넣어야 하면:

```powershell
node .\backfill_xlsx.mjs --xlsx .\받은파일.xlsx --store .\store.json
```

`--dry` 를 붙이면 저장하지 않고 어떻게 읽히는지만 보여줍니다.

## 5. 버셀 연결

<https://vercel.com> 가입 후 한 번만 수동으로 올립니다.

```powershell
npx.cmd vercel login
```

```powershell
npx.cmd vercel deploy .\site --prod
```

`npx` 가 아니라 **`npx.cmd`** 입니다. 그냥 `npx` 라고 치면 PowerShell 이 `npx.ps1` 을 찾다가
실행 정책에 막혀 *이 시스템에서 스크립트를 실행할 수 없으므로* 오류가 납니다.
`.cmd` 를 붙이면 그 검사를 지나가므로 정책을 바꿀 필요가 없습니다.

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

## 5-1. 인원을 저장할 자리 만들기

여기까지만 해도 화면은 잘 나옵니다. 이건 **인원을 어느 PC 에서 고쳐도 모든 PC 에 반영되게**
하는 부분입니다. 클릭 몇 번이고 한 번만 하면 됩니다.

1. <https://vercel.com/dashboard> 에서 방금 만든 프로젝트를 엽니다
2. 위쪽 **Storage** 탭 → **Create Database** → **Blob** 을 고릅니다

`Create Blob Store` 창이 뜨면 이렇게 둡니다.

| 칸 | 어떻게 |
|---|---|
| Store Name | 적혀 있는 그대로 두면 됩니다 |
| Region | 그대로 두면 됩니다 (파일 하나뿐이라 어디든 상관없습니다) |
| Access | **Private** 그대로 둡니다. 값은 이 화면의 서버만 읽고 씁니다 |
| Custom Environment Variable Prefix | 비워 둡니다 |
| **Add a read-write token env var to this connection** | **반드시 체크하세요** |

마지막 체크만 빠지면 값을 읽고 쓸 열쇠가 안 들어가서, 눌러 고쳐도 그 화면에서만 바뀝니다.
체크하면 `BLOB_READ_WRITE_TOKEN` 이 프로젝트에 저절로 들어갑니다. 따로 적을 것은 없습니다.

만든 뒤 **Connect Project** 로 이 프로젝트에 연결합니다.

그다음 한 번 배포하면 적용됩니다. 수집은 건너뛰고 화면만 다시 올립니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\crm4_daily.ps1 -SkipCollect
```

**확인** — 화면에서 인원을 눌러 12 로 바꾸고, 다른 PC(또는 휴대폰)에서 같은 주소를 열어
12 로 보이면 된 겁니다. 저장이 안 되면 표 위에 빨간 글씨로 왜 안 됐는지 뜹니다.

주의 — 주소를 아는 사람은 누구나 인원을 바꿀 수 있습니다. 사내에서만 쓰는 숫자라 큰 문제는
아니지만, 값이 이상하면 누가 눌러 본 것일 수 있습니다. 1~999 밖의 값은 애초에 안 들어갑니다.

## 6. 자동 실행 등록

**관리자 권한 PowerShell**에서 실행합니다.

```powershell
$action  = New-ScheduledTaskAction -Execute "powershell.exe" `
           -Argument "-ExecutionPolicy Bypass -File C:\CRM자동화\crm4_daily.ps1" `
           -WorkingDirectory "C:\CRM자동화"

# 새벽 6시에 시작하고, 실패했을 때를 대비해 2시간마다 저녁 6시까지 다시 시도한다.
# 이미 받아 둔 날이면 CRM4 를 건드리지 않고 곧바로 끝나므로 여러 번 돌아도 방해되지 않는다.
$trigger = New-ScheduledTaskTrigger -Daily -At 6am
$trigger.Repetition = (New-ScheduledTaskTrigger -Once -At 6am `
                       -RepetitionInterval (New-TimeSpan -Hours 2) `
                       -RepetitionDuration (New-TimeSpan -Hours 12)).Repetition

$set = New-ScheduledTaskSettingsSet -StartWhenAvailable `
       -ExecutionTimeLimit (New-TimeSpan -Hours 3) -MultipleInstances IgnoreNew `
       -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 15)

Register-ScheduledTask -TaskName "CRM4 일일수집" -Action $action -Trigger $trigger `
                       -Settings $set -RunLevel Highest -Force
```

매일 새벽 6시에 **어제치 하루만** 받습니다. 일요일은 영업을 하지 않으므로 건너뜁니다.

**실패했을 때의 보험이 세 겹입니다.**

| | |
|---|---|
| 하루 안 재시도 | 6시에 실패해도 8·10·12·14·16·18시에 다시 시도합니다. 어제치를 오늘 받는 것은 몇 시에 받든 기준이 같습니다 |
| 즉시 재시작 | 스크립트가 비정상 종료하면 15분 뒤 최대 2번 다시 실행합니다 |
| 놓친 실행 | 그 시각에 PC가 자고 있었으면 깨어난 뒤 바로 실행합니다 (`-StartWhenAvailable`) |

이미 받아 둔 날이면 CRM4 를 건드리지 않고 몇 초 만에 끝나므로,
업무 시간에 재시도가 돌아도 화면을 빼앗기지 않습니다.

PC 를 계속 켜 두시므로 시각 지정이 맞습니다. 로그온 방식은 재부팅을 해야 돌기 때문에 쓰지 않습니다.
시각을 바꾸려면 `-At 6am` 만 고치면 됩니다. 자정만 넘기면 언제든 어제치가 온전히 나옵니다.

그날 못 돌면 그 날짜는 비워 둡니다. 뒤늦게 받으면 기준이 달라지기 때문입니다.

### PC 를 계속 켜 둘 때 같이 해 둘 것

화면이 꺼지거나 잠기면 스크립트가 창을 앞으로 가져오지 못합니다.

1. **설정 › 시스템 › 전원** — 화면 끄기와 절전을 모두 `안 함` 으로
2. **화면 보호기 설정** — `다시 시작할 때 로그온 화면 표시` 체크 해제
3. **CRM4 를 시작 프로그램에 등록** 하고 `통합고객목록` 창을 열어 둔 채로 두기

수동 잠금(Win+L)도 하지 마세요. 잠긴 상태에서는 실패하고 그날은 비워집니다.

## 화면 갱신을 따로 돌리기

데이터는 하루 한 번이면 되지만, 화면은 고치는 대로 바로 보이는 편이 낫습니다.
그래서 작업을 둘로 나눕니다.

| 작업 | 언제 | 무엇을 |
|---|---|---|
| `CRM4 일일수집` | 매일 06:00 | CRM4에서 어제치를 받아 세고 올린다 |
| `CRM4 화면갱신` | 10분마다 | 화면 틀만 확인. 바뀌었을 때만 다시 만들어 올린다 |

화면갱신 작업은 **CRM4를 건드리지 않습니다.** 틀이 그대로면 주소만 한 번 확인하고
로그도 남기지 않은 채 곧바로 끝납니다. 바뀌었을 때만 집계본으로 화면을 다시 만들어 올립니다.
수집이 도는 중에는 알아서 다음 차례로 미룹니다.

**관리자 권한 PowerShell**에서 통째로 붙여넣으세요.

```powershell
$a = New-ScheduledTaskAction -Execute "powershell.exe" `
     -Argument "-ExecutionPolicy Bypass -File C:\CRM자동화\crm4_daily.ps1 -TemplateOnly" `
     -WorkingDirectory "C:\CRM자동화"

$t = New-ScheduledTaskTrigger -Daily -At 12am
$t.Repetition = (New-ScheduledTaskTrigger -Once -At 12am `
                 -RepetitionInterval (New-TimeSpan -Minutes 10) `
                 -RepetitionDuration (New-TimeSpan -Hours 24)).Repetition

$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
     -ExecutionTimeLimit (New-TimeSpan -Minutes 20)

Register-ScheduledTask -TaskName "CRM4 화면갱신" -Action $a -Trigger $t `
                       -Settings $s -RunLevel Highest -Force
```

등록되면 화면 수정은 올리는 즉시, 늦어도 10분 안에 반영됩니다.

확인:

```powershell
Get-ScheduledTaskInfo -TaskName "CRM4 화면갱신" | Select-Object LastRunTime, LastTaskResult, NextRunTime
```

`-RepetitionDuration` 에서 오류가 나면 `24` 를 `23` 으로 바꾸세요.

## 나중에 화면을 바꾸고 싶을 때

**보통은 아무것도 안 하셔도 됩니다.** 스크립트가 매일 돌기 전에 아래 주소에서
`template.html` 을 받아 갈아끼웁니다. 화면을 고쳐 올려 두면 다음 날 아침에 반영됩니다.

```
https://raw.githubusercontent.com/winggeul/claude-code/refs/heads/claude/crm-data-connection-2nbows/dashboard/index.html
```

받은 파일이 이상하면(너무 작거나, 데이터 자리가 없거나) 버리고 기존 것을 씁니다.
인터넷이 안 되면 그냥 넘어갑니다. 화면 때문에 그날 수집을 날리지 않습니다.
바꾸기 전 파일은 `template.html.bak` 으로 남습니다.

받아오지 않게 하려면 `crm4_daily.ps1` 위쪽의 `$TemplateUrl` 을 `""` 로 두거나
`-SkipTemplate` 을 붙여 실행하면 됩니다.

직접 바꾸실 때는:

1. 새 `template.html` 을 받아 폴더 안의 것과 바꿔치기
2. 바로 반영하려면 아래 실행 (수집은 건너뛰고 화면만 다시 만들어 올립니다)

```powershell
powershell -ExecutionPolicy Bypass -File .\crm4_daily.ps1 -SkipCollect
```

그냥 두셔도 다음 날 새벽 6시에 새 모양으로 올라갑니다.
`store.json` 은 건드리지 않으므로 쌓아 둔 숫자는 그대로입니다.

## 확인과 문제 해결

| 증상 | 확인할 것 |
|---|---|
| 아무 일도 안 일어남 | `C:\CRM자동화\logs` 의 `run_*.txt` |
| 창을 못 찾는다고 나옴 | CRM4 `통합고객목록` 창이 떠 있는지 |
| 클릭 안 하고 멈춤 | 다른 프로그램이 앞에 있었던 것. 정상 동작이며 다시 실행하면 된다 |
| 받아온 날짜가 다르다고 나옴 | 스스로 등록일 체크박스를 뒤집고 다시 받는다. 그대로 두면 된다 |
| 숫자가 안 맞음 | `-KeepRaw` 로 실행해 원본 CSV 를 남긴 뒤 대조 |

원본 CSV 는 `C:\CRM자동화\raw` 에 잠깐 머물다가 집계가 끝나면 지워집니다.
집계가 실패하면 지우지 않으므로 다음 실행에서 다시 시도됩니다.
이 PC 밖으로 나가는 것은 건수 집계뿐입니다.
