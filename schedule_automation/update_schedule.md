# update_schedule.py · check_mail_attachment.py  
<sub>2026-10-01  Jonghyun Park w/ Claude</sub>  

## 개요
캠페인 법인별 일정 파일을 자동으로 Auto 정제 파일에 붙여넣는 스크립트.

메일함 감시 → 첨부 자동 저장 → 일정 시트 정제 → Auto 워크북 반영 → 변경 셀 음영까지가 한 묶음이다.

## 파일 구성

| 파일 | 역할 |
|---|---|
| `check_mail_attachment.py` | **메일 제목** 키워드로 첨부 감지 |
| `update_schedule.py` | 최신 소스 선택 → Auto 워크북 반영 → 변경 셀 음영 → 강제 재계산 |
| `create_schtasks_v2.txt` | 작업 스케줄러 등록 명령어 모음 (감시 변형·정제판 기준 — 이 문서의 두 스크립트는 아래 "작업 스케줄러 등록" 명령을 쓴다) |

> 이 문서는 일정 자동화의 **앞 세대** 도구 2개를 다룬다 — 제목 기준 첨부 감지와, 요약 정제 단계가 없는
> 워크북 반영 스크립트. 첨부파일명·수신기간 등 조건이 다른 감시 변형과 요약 시트 자동 정제판,
> 반영 대상 워크북 예시는 같은 폴더에 있다 — [`README.md`](README.md) 참고.


> 📄 정제 대상 Auto 파일의 시트 구조 상세: [`26_Schedule_separate(Auto).md`](<26_Schedule_separate(Auto).md>) (한글판: [`26_Schedule_separate(Auto)-kr.md`](<26_Schedule_separate(Auto)-kr.md>))

## 동작 순서
1. `1.고객 법인 일정 파일/` 폴더에서 최신 파일 자동 선택
2. 소스 파일 첫 번째 시트 **B3:J** 값 읽기 (범위는 스크립트 상단 `SRC_MIN_ROW` / `SRC_MIN_COL` / `SRC_MAX_COL`)
   - `FILL_MERGED_COLS`(기본 E열 Participation) 열에서 **세로로 병합된 셀**은 병합된 모든 행에 같은 값을 채워 읽는다 —
     법인 여러 곳이 `O` 한 칸을 공유하는 일정표에서 아래쪽 법인이 빈칸으로 들어가던 문제 방지
3. **Auto 파일에 그 결과가 실제로 저장돼 있는지 확인** → 이미 반영돼 있으면 `[SKIP]` 으로 종료 (아래 "재실행 판정")
4. Auto 파일의 `고객법인일정파일` 시트 **B2:K999** 값·음영 클리어 후 **B2**부터 값 붙여넣기 (범위는 `TGT_*` 상수)
   - 대상 영역과 겹치는 **병합셀은 클리어 전에 자동 해제** — `MergedCell` 은 값 설정이 불가(read-only)해서 그냥 두면 예외로 중단됨
   - 소스가 B2B/B2C 2세트(B~N)로 확장된 일정표라면 `SRC_MAX_COL=14` / `TGT_MAX_COL=14` 로 바꾸고 `COMPARE` 에 `{8: 10, 10: 12}` 를 추가 (스크립트 상단 주석 참고)
5. 같은 시트 **D1** 에 소스 파일명 기록 (어느 원본에서 왔는지 추적용)
6. 이전 소스 파일과 비교해 변경된 셀에 노란 음영 표시
7. win32com으로 Excel을 백그라운드 실행 → `CalculateFull()` → 재계산된 상태로 저장
   - `wb_com.Close()` 는 `try/except TypeError` 로 감싸져 있다 — `%TEMP%\gen_py\<파이썬버전>\` 타입라이브러리 캐시가 없으면(파이썬을 새로 깔면 비어 있음) win32com 이 **지연 바인딩**으로 동작해 `wb_com.Close` **속성 접근만으로 COM 메서드가 실행**되고 `True` 가 반환된다 → 이어지는 `()` 가 `TypeError: 'bool' object is not callable`. 그 시점엔 저장·닫기가 이미 끝났지만 **다음 줄인 마커 기록이 안 돼서** 다음 실행이 같은 소스를 재처리한다. `%TEMP%` 캐시는 디스크 정리로 언제든 사라지므로 gen_py 재생성이 아니라 코드로 막는다.
8. 저장 직후 Auto 파일을 다시 읽어 검증 → **통과했을 때만** 마커 파일에 소스 파일명·mtime 기록

### 재실행 판정

마커 파일이 아니라 **Auto 파일에 실제로 저장된 내용**으로 판정한다 (`target_is_saved()`).

| 확인 | 내용 |
|---|---|
| D1 스탬프 | `고객법인일정파일` 시트 D1 의 소스 파일명이 지금 소스와 같은지 |
| 붙여넣기 값 | B2 부터 소스에서 읽은 값과 셀 단위로 같은지 (날짜는 date 로, 문자열은 앞뒤 공백·줄바꿈을 맞춰 비교) |
| 잔재 행 | 붙여넣은 마지막 행 아래 ~ `TGT_MAX_ROW` 에 옛 데이터가 남아 있지 않은지 |

- 하나라도 어긋나면 `[갱신 필요] <사유>` 를 찍고 다시 붙여넣는다. 마커가 '처리 완료'였다면 `[경고]` 도 같이 찍힌다.
- 종전(마커만 비교)에는 저장 뒤 Auto 파일이 Excel·OneDrive 에 의해 되돌려지거나 붙여넣기가 어긋난 채 남으면 계속 SKIP 됐다.
- 내용이 같아도 다시 붙여넣으려면 `python update_schedule.py --force`.
- 마커 파일은 기록·경고용으로만 남는다. 지워도 판정에 영향이 없다.

## 최신 파일 선택 기준
파일명에서 아래 정렬 키를 추출하고, 가장 큰 튜플을 가진 파일을 최신으로 판단:

```python
(doc_date, hhmm, ver_float, ver_int, mail_date)
```

### A형 — 8자리 날짜(YYYYMMDD) 포함

| 패턴 | 예시 | 정렬 키 |
|---|---|---|
| `YYYYMMDD_HHMM[_YYMMDD[_HHMM]]` | `20260415_1543_260417` | `(20260415, 1543, 0, 0, 2604170000)` |
| `YYYYMMDD_vN[_YYMMDD[_HHMM]]`   | `20260421_v2`          | `(20260421, 0, 0, 2, 0)`             |
| `YYYYMMDD_YYMMDD[_HHMM]`        | `20260401_260406`      | `(20260401, 0, 0, 0, 2604060000)`    |
| `YYYYMMDD`                      | `20260421`             | `(20260421, 0, 0, 0, 0)`             |

**메일 수신일시(마지막 성분)**: `check_mail_attachment.py` 가 같은 파일명 재수신 시 붙이는 suffix.
`_YYMMDD` → `_YYMMDD_HHMM` 로 확장됐고, 키에서는 `YYMMDD × 10000 + HHMM` 정수로 합쳐 비교.

```
···20260421_v1_260722       → (20260421, 0, 0, 1, 2607220000)   시각 없는 옛 파일
···20260421_v1_260722_1432  → (20260421, 0, 0, 1, 2607221432)
···20260421_v1_260722_1840  → (20260421, 0, 0, 1, 2607221840)  ← 최신
```

확장 전에는 뒤 `_HHMM` 을 무시해 같은 날 두 번 온 파일이 **동점**이 됐고, 최신 판정이 파일 나열 순서에 좌우됐음.
옛 `_YYMMDD` 형식은 그대로 파싱되므로 이미 저장된 파일의 재정렬 문제 없음.

전체 예시 (오름차순 정렬):
```
···20260401_260406.xlsx       → (20260401, 0, 0, 0, 2604060000)
···20260407_v2_260408.xlsx    → (20260407, 0, 0, 2, 2604080000)
···20260409_260409.xlsx       → (20260409, 0, 0, 0, 2604090000)
···20260409_v1_260409.xlsx    → (20260409, 0, 0, 1, 2604090000)
···20260409_v2_260409.xlsx    → (20260409, 0, 0, 2, 2604090000)
···20260409_v2_260414.xlsx    → (20260409, 0, 0, 2, 2604140000)
···20260415_1543.xlsx         → (20260415, 1543, 0, 0, 0)
···20260415_1543_260417.xlsx  → (20260415, 1543, 0, 0, 2604170000)
···20260420_0906.xlsx         → (20260420, 906, 0, 0, 0)
···20260421.xlsx              → (20260421, 0, 0, 0, 0)
···20260421_v2.xlsx           → (20260421, 0, 0, 2, 0)  ← 최신
```

### B형 — 6자리 날짜(YYMMDD) + 소수점 버전

| 키 | 정규식 | 예시 | 파싱 결과 |
|---|---|---|---|
| `date6`   | `(?<!\d)\d{6}(?!\d)` | `_260325`  | `260325` |
| `version` | `_v(\d+\.\d+)`        | `_v0.441`  | `0.441`  |
| `suffix`  | `_(\d{1,5})$`         | `_2` (끝)  | `2`      |

> **주의**: 끝번호 정규식 `_(\d{1,5})$` 이 메일 suffix 의 시각(`_1432`)을 버전 끝번호로 오인하므로,
> 끝의 `_YYMMDD_HHMM`(날짜+시각이 **둘 다** 있는 형태 = 스크립트가 붙인 것)을 먼저 떼어낸 뒤 판정한다.
> 날짜만 있는 `_YYMMDD` 는 문서날짜일 수 있어 떼지 않음.

```
_v0.44_260319   → (260319, 0, 0.44, 0, 0)
_v0.48_260420   → (260420, 0, 0.48, 0, 0)
_v0.49_260420   → (260420, 0, 0.49, 0, 0)  ← 최신
```

## 값 변환 규칙
| 소스 셀 타입 | 변환 결과 |
|---|---|
| `datetime` | `date` 객체 (number_format: `YYYY-MM-DD`) |
| WEEKNUM 수식 결과 (정수) | `W01` 형식 문자열 |
| 나머지 | 그대로 |

## 전후 비교 음영 표시
- 매칭 기준: **C열 (Subs)**
- 비교 대상: **E열** (Participation), **F열** (Start at), **H열** (End at)
- 이전 파일 없으면 스킵
- 매 실행마다 기존 음영 초기화 후 재적용

## 파일 경로
| 구분 | 경로 |
|---|---|
| 소스 폴더 | `02. SCHEDULE/1.고객 법인 일정 파일/` |
| 업데이트 대상 | `02. SCHEDULE/*Auto*.xlsx` (자동 탐색) |
| 마커 파일 | `BASE/schedule_last_source.txt` — `파일명\|mtime` 형식으로 저장 (Auto 파일과 같은 폴더에 두어 프로젝트별 독립 관리. 다른 캠페인으로 fork 시 BASE만 교체하면 마커도 따라감). **기록·경고용** — SKIP 판정은 Auto 파일 내용으로 한다 |

## 실행 방법
```bash
python update_schedule.py
python update_schedule.py --force   # 판정과 무관하게 다시 붙여넣기
```

## 스케줄 작업
- 작업 이름: `cmp_schedule_update`
- 실행 주기: 20분마다
- 시작: 10:00
- 기간: ~ 2026-05-15 (예시 — `/ed` 는 캠페인 종료일이므로 실제 값으로 교체)
- 조건: 로그온 중일 때만 실행 (`/it`)
- 배터리: 배터리 전원에서도 실행 (schtasks 등록 후 PowerShell 추가 설정 필요 → 아래 참고)

## 작업 스케줄러 등록

Python 스크립트를 작업 스케줄러에 직접 등록. bat/vbs 래퍼 불필요.

> 💡 **창 없이 실행하려면 `pythonw.exe` 사용**  
> `python.exe`는 콘솔 앱이라 실행 시 cmd 창이 팝업됨.  
> `pythonw.exe`(같은 경로에 있음)를 대신 쓰면 창이 전혀 뜨지 않음.  
> 단, stdout/stderr 출력이 버려지므로 로그가 필요하면 스크립트 내에서 파일로 기록해야 함.

> ⚠️ `/tr` 경로에 공백이 있으면 schtasks가 경로를 잘라 오류 발생 (0x80070005).  
> python.exe 경로와 py 파일 경로를 각각 `\"...\"` 로 감싸야 공백이 포함된 경로도 처리 가능.

### CLI 등록 명령어

```bat
schtasks /create /tn cmp_mail_check ^
  /tr "\"C:\Python3xx\pythonw.exe\" \"C:\Users\user_name\OneDrive - company_name\user_id\...\check_mail_attachment.py\"" ^
  /sc minute /mo 20 /st 09:55 /ed 2026/05/15 /it /f

schtasks /create /tn cmp_schedule_update ^
  /tr "\"C:\Python3xx\pythonw.exe\" \"C:\Users\user_name\OneDrive - company_name\user_id\...\update_schedule.py\"" ^
  /sc minute /mo 20 /st 10:00 /ed 2026/05/15 /it /f
```

같은 폴더의 `create_schtasks_v2.txt` 는 감시 변형·정제판 기준 명령어 모음이다 (형식은 동일).

| 옵션 | 설명 |
|---|---|
| `/sc minute /mo 20` | 20분마다 실행 |
| `/st 09:55 / 10:00` | 각 스크립트 시작 시각 |
| `/it` | 로그온 중일 때만 실행 |
| `/f` | 동일 이름 작업 덮어쓰기 |

### 배터리 모드 허용 (등록 후 추가 실행)

`schtasks`는 배터리 조건을 직접 지정할 수 없으므로, 등록 후 아래 명령어를 별도로 실행.

**[권장] PowerShell 창에서 직접 실행** — 래퍼 없이 한 줄, 빠름:

```powershell
$names = 'cmp_mail_check','cmp_schedule_update'; foreach ($n in $names) { $t = Get-ScheduledTask -TaskName $n; $t.Settings.DisallowStartIfOnBatteries = $false; $t.Settings.StopIfGoingOnBatteries = $false; Set-ScheduledTask -InputObject $t }
```

또는 가독성 버전:

```powershell
$names = 'cmp_mail_check', 'cmp_schedule_update'
foreach ($n in $names) {
    $t = Get-ScheduledTask -TaskName $n
    $t.Settings.DisallowStartIfOnBatteries = $false
    $t.Settings.StopIfGoingOnBatteries = $false
    Set-ScheduledTask -InputObject $t
}
```

**[cmd/bat에서 실행할 때만] `powershell -Command` 래퍼 사용**:

```bat
powershell -Command "$names = 'cmp_mail_check','cmp_schedule_update'; foreach ($n in $names) { $t = Get-ScheduledTask -TaskName $n; $t.Settings.DisallowStartIfOnBatteries = $false; $t.Settings.StopIfGoingOnBatteries = $false; Set-ScheduledTask -InputObject $t }"
```

⚠ **PowerShell 창에서 래퍼 형태를 쓰지 말 것** — 바깥 PS가 `$names`, `$n`, `$t`를 자기 변수로 먼저 치환(빈 값)해버려 `foreach 뒤에 변수 이름이 없습니다` 오류가 발생합니다.

### Python 경로 확인

```bat
where pythonw
:: 예) C:\Python314\pythonw.exe
```

### GUI 등록

1. `taskschd.msc` 실행
2. **작업 만들기** → 일반 탭: 이름 입력, "사용자가 로그온할 때만 실행" 선택
3. **트리거** 탭 → 새로 만들기 → 반복 주기 설정
4. **동작** 탭 → 새로 만들기:
   - 프로그램/스크립트: `C:\Python3xx\pythonw.exe` (창 없이 실행; 일반 python.exe 쓰면 cmd 창 팝업됨)
   - 인수 추가: `"C:\Users\user_name\OneDrive - company_name\...\update_schedule.py"`
5. **조건** 탭 → 전원 섹션 → **"AC 전원이 연결된 경우에만 작업 시작" 체크 해제**

## check_mail_attachment.py

Outlook 수신함을 폴링해 일정 파일 첨부를 자동으로 로컬에 저장하는 스크립트.  
`update_schedule.py`와 쌍으로 작업 스케줄러에 등록해 주기적으로 실행합니다.  
새 소스 파일이 도착하면 다음 `update_schedule.py` 실행 시 자동 반영됩니다.

### 동작 순서

1. `win32com.client`로 Outlook 접근 — 기본 프로필이 아니라 **`STORE_NAME` 과 이름이 일치하는 메일함(store)** 을 찾아 그 받은편지함을 연다. 공유 메일함을 쓰므로 실행 전 스크립트 상단 `STORE_NAME` 을 본인 환경의 메일함 이름으로 반드시 교체할 것 (못 찾으면 그대로 종료)
2. 제목 필터 조건에 맞는 메일만 선별 (스크립트 상단 `SUBJECT_KEYS` 참고)
3. 처리 이력(`campaign_mail_processed_ids.txt`)에 없는 메일만 처리
4. `.xlsx` 첨부파일을 지정 폴더에 저장
5. 처리한 메일의 EntryID를 이력 파일에 추가

### 실행 방법

```bash
python check_mail_attachment.py
```

작업 스케줄러 등록 방법은 위의 **작업 스케줄러 등록** 섹션 참고 (`cmp_mail_check` 작업).

### 재처리 방지 (EntryID 마커)

처리 완료한 메일의 `EntryID`를 마커 파일에 기록해, 다음 실행 시 동일 메일을 건너뜀.

```
C:\Users\user_name\Documents\campaign_mail_processed_ids.txt
```

- 파일이 없으면 자동 생성
- 새 첨부파일이 저장되거나 이미 처리된 것으로 간주된 경우에만 EntryID 기록

### 중복 파일명 처리

동일한 파일명의 첨부파일이 재전송되면 수신일시(`_yymmdd_HHMM`)를 붙여 별도 저장.

```
원본 파일명: [앞부분]···20260409_v2.xlsx  (이미 저장됨)
신규 수신 시: [앞부분]···20260409_v2_260414_1432.xlsx  (수신 260414 14:32)
```

- 날짜만(`_yymmdd`) 붙이면 같은 날 같은 이름으로 두 번 올 때 붙일 이름이 같아져 두 번째가
  **스킵되어 유실**되고, `update_schedule.py` 의 최신 판정도 동점이 됨 → 시각까지 붙인다.
- 수신일시까지 붙인 파일도 이미 존재하면 저장 없이 스킵 (같은 메일 = 같은 분 단위 시각).

### MAX_PATH 우회

저장 경로가 Windows 260자 제한을 초과하면 `att.SaveAsFile()` 오류 발생.  
임시 파일에 먼저 저장 후 `shutil.move()`로 이동하는 방식으로 우회.

> 레지스트리 설정(`LongPathsEnabled = 1`)도 병행 적용 필요 → [`enable_long_path.md`](enable_long_path.md) 참고

