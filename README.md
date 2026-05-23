# Auto Mailing

특정 폴더의 엑셀 파일을 날짜 기준으로 자동 선택하여 이메일로 발송하는 Python 스크립트입니다.

## Features

- 파일명의 `_YYMMDD` 패턴을 인식하여 최신 파일 또는 오늘 날짜 파일 자동 선택
- Outlook COM 또는 SMTP (Office365/Gmail) 두 가지 발송 방식 지원
- Windows 작업 스케줄러 연동용 `.bat` 파일 포함
- `--dry-run` 모드로 발송 전 미리보기 가능
- 실행 로그 자동 기록 (`run_mailer.log`)

## Requirements

- Windows
- Python 3
- pywin32 (Outlook COM 방식 사용 시)

```bash
pip install pywin32
```

## Usage

```bash
# 가장 최신 날짜 파일 발송
python run_mailer.py

# 오늘 날짜 파일만 발송 (없으면 skip)
python run_mailer.py --today-only

# 발송 미리보기
python run_mailer.py --dry-run

# 수신자 직접 지정
python run_mailer.py --to user@example.com --cc cc@example.com
```

## Configuration

`run_mailer.py` 상단에서 설정:

| 항목 | 설명 |
|---|---|
| `TO` / `CC` / `BCC` | 수신자 이메일 |
| `SUBJECT_TEMPLATE` | 메일 제목 (`{filename}`, `{file_date}`, `{now}` 플레이스홀더) |
| `BODY_TEMPLATE` | 메일 본문 |
| `ATTACH_DIR` | 엑셀 파일 폴더 경로 |
| `SEND_METHOD` | `"outlook"` 또는 `"smtp"` |

## 스케줄러 등록

Windows 작업 스케줄러에 `.bat` 파일을 등록하면 자동 발송됩니다.

- `run_mailer_latest.bat` — 최신 파일 발송
- `run_mailer_today_only.bat` — 오늘 파일만 발송

자세한 설정은 `run_mailer.md`를 참고하세요.

## License

MIT
