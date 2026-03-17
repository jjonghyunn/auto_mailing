# run_mailer — xlsx 자동 발송 스크립트

`Downloads/mailing_test` 폴더의 xlsx 파일을 Outlook으로 자동 발송합니다.

---

## 파일 구성

| 파일 | 역할 |
|------|------|
| `run_mailer.py` | 메인 스크립트 (설정 + 발송 로직) |
| `run_mailer.md` | 이 문서 |
| `run_mailer.log` | 실행 로그 (자동 생성) |
| `run_mailer_latest.bat` | 최신 날짜 파일 발송용 배치 |
| `run_mailer_today_only.bat` | 오늘 날짜 파일만 발송용 배치 |

---

## 설정 (`run_mailer.py` 상단)

### 수신자

엔터로 구분, 빈 줄 자동 무시:

```python
TO = """
jonghyun.park@test.com
another@example.com
"""

CC = """
manager@example.com
"""

BCC = """
"""
```

### 메일 내용

```python
SUBJECT_TEMPLATE = "[Auto] {filename} 발송"   # {filename} → 첨부 파일명 자동 삽입

BODY_TEMPLATE = """\
안녕하세요,
...
파일명: {filename}
날짜  : {file_date}
발송시: {now}
"""
```

### 첨부 파일 소스 폴더

```python
ATTACH_DIR = r"C:\Users\user_name\Downloads\mailing_test"
```

- `_YYMMDD.xlsx` 패턴 파일만 인식 (예: `report_260317.xlsx`)
- 패턴 없는 파일은 무시

### 발송 방식

```python
SEND_METHOD = "outlook"   # 로컬 Outlook 앱 사용 (로그인 불필요)
# SEND_METHOD = "smtp"    # SMTP 직접 발송 (하단 SMTP 설정 필요)
```

---

## 실행 방법

### 직접 실행

```bash
# 기본: 가장 최신 날짜 파일 발송
python run_mailer.py

# 오늘 날짜 파일만 발송 (없으면 발송 안 함)
python run_mailer.py --today-only

# 발송 없이 결과 미리보기
python run_mailer.py --dry-run
python run_mailer.py --today-only --dry-run

# 수신자 임시 덮어쓰기 (스크립트 설정 무시)
python run_mailer.py --to a@company.com b@company.com --cc c@company.com
```

### 배치파일 실행

```
run_mailer_latest.bat       → python run_mailer.py
run_mailer_today_only.bat   → python run_mailer.py --today-only
```

---

## 작업 스케줄러 등록 / 관리

> `/tn` = 작업 이름 (Task Name). 조회·수정·삭제 시 동일하게 사용.
> 로그인된 세션이 있을 때만 실행됨 (화면 잠금 상태도 세션 유지이므로 실행됨).

### 1회성 등록

```cmd
schtasks /create /tn "mailer_test_once" /tr "\"C:\Users\user_name\OneDrive - test Corporation\pjh\2.data\99.PY,SQL-250429\00.py_notebook\run_mailer_latest.bat\"" /sc once /st HH:MM /it /f
```

```cmd
:: 확인
schtasks /query /tn "mailer_test_once"

:: 삭제
schtasks /delete /tn "mailer_test_once" /f
```

### 매일 반복 등록

```cmd
schtasks /create /tn "auto_mailing_260317" /tr "\"C:\Users\user_name\OneDrive - test Corporation\pjh\2.data\99.PY,SQL-250429\00.py_notebook\run_mailer_latest.bat\"" /sc daily /st HH:MM /it /f
```

예시 — 2026/03/17 14시 시작, 2026/03/23 14시 마지막 발송:

```cmd
schtasks /create /tn "auto_mailing_260317" /tr "\"C:\Users\user_name\OneDrive - test Corporation\pjh\2.data\99.PY,SQL-250429\00.py_notebook\run_mailer_latest.bat\"" /sc daily /st 14:00 /sd 2026/03/17 /ed 2026/03/23 /it /f
```

### 관리 명령어

```cmd
:: 확인
schtasks /query /tn "auto_mailing_260317"

:: 실행 시간 변경
schtasks /change /tn "auto_mailing_260317" /st HH:MM

:: 종료일 설정
schtasks /change /tn "auto_mailing_260317" /ed YYYY/MM/DD

:: 비활성화 (일시 중지)
schtasks /change /tn "auto_mailing_260317" /disable

:: 다시 활성화
schtasks /change /tn "auto_mailing_260317" /enable

:: 삭제 (완전 제거)
schtasks /delete /tn "auto_mailing_260317" /f
```

---

## 파일 선택 로직

```
_YYMMDD 패턴 xlsx 스캔
│
├── --today-only 옵션
│   ├── 오늘 날짜 파일 있음 → 발송
│   └── 없음 → 발송 안 함, 정상 종료 (exit 0)
│
└── 기본 (옵션 없음)
    └── 가장 최신 날짜 파일 → 발송
```

---

## 로그

`run_mailer.log` 에 매 실행 결과 누적 기록:

```
2026-03-17 09:00:01 [INFO] ============================================================
2026-03-17 09:00:01 [INFO] run_mailer 시작  today-only=False  dry-run=False
2026-03-17 09:00:01 [INFO] 발견된 파일 3개:
2026-03-17 09:00:01 [INFO]   2026-03-15  report_260315.xlsx
2026-03-17 09:00:01 [INFO]   2026-03-16  report_260316.xlsx
2026-03-17 09:00:01 [INFO]   2026-03-17  report_260317.xlsx
2026-03-17 09:00:01 [INFO] 선택된 파일: report_260317.xlsx  (날짜: 2026-03-17)
2026-03-17 09:00:02 [INFO] Outlook으로 발송 완료.
2026-03-17 09:00:02 [INFO] 발송 성공.
```

---

## 의존성

```bash
pip install pywin32   # SEND_METHOD = "outlook" 사용 시 필요
```
