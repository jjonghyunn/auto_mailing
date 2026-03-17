# -*- coding: utf-8 -*-
"""
run_mailer.py
──────────────────────────────────────────────────────────────────────
Downloads/mailing_test 폴더에서 xlsx 파일 중 _YYMMDD 패턴을 가진
가장 최신 파일(또는 오늘 날짜 파일)을 이메일로 발송합니다.

사용법:
  python run_mailer.py                  # 최신 날짜 파일 발송 (기본)
  python run_mailer.py --today-only     # 오늘 날짜 파일만 발송 (없으면 중단)
  python run_mailer.py --dry-run        # 발송 없이 파일 선택 결과만 확인
  python run_mailer.py --today-only --dry-run

발송 방식:
  SEND_METHOD = "outlook"  → 로컬 Outlook 앱 사용 (win32com, 인증 불필요)
  SEND_METHOD = "smtp"     → SMTP 직접 발송 (SMTP_* 설정 필요)
"""

import re
import sys
import argparse
import smtplib
import logging
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path

# ════════════════════════════════════════════════════════════════════
#  ★ 설정 구역 — 여기만 수정하세요
# ════════════════════════════════════════════════════════════════════

# ── 수신자 (엔터로 구분, 빈 줄 무시) ─────────────────────────────
TO = """
jonghyun.park@test.com
"""

CC = """
test1@test.com
test2@test.com
"""

BCC = """
"""

# ── 메일 내용 ─────────────────────────────────────────────────────
SUBJECT_TEMPLATE = "[Auto] {filename} 발송"   # {filename} 자리에 파일명 자동 삽입
BODY_TEMPLATE    = """\
안녕하세요,
작업 스케줄러 auto mailing test입니다.
첨부파일은 오늘 날짜와 같은 파일을 우선 첨부하여 발송합니다.


파일명: {filename}
날짜  : {file_date}
발송시: {now}

감사합니다.
"""

# ── 첨부 파일 소스 폴더 ───────────────────────────────────────────
ATTACH_DIR = r"C:\Users\user_name\Downloads\mailing_test"

# ── 발송 방식: "outlook" 또는 "smtp" ──────────────────────────────
SEND_METHOD = "outlook"   # 사내 PC 환경이면 "outlook" 권장. 하단 SMTP 설정 필요 없음. (윈도우 + Outlook 설치 환경에서만 작동)

# ── SMTP 설정 (SEND_METHOD = "smtp" 일 때만 사용) ─────────────────
SMTP_HOST     = "smtp.office365.com"   # Office365: smtp.office365.com / Gmail: smtp.gmail.com
SMTP_PORT     = 587
SMTP_USER     = "jonghyun.park@test.com"
SMTP_PASSWORD = "your_password"        # 앱 비밀번호 권장
FROM_ADDR     = "jonghyun.park@test.com"

# ════════════════════════════════════════════════════════════════════
#  내부 로직 (수정 불필요)
# ════════════════════════════════════════════════════════════════════

def _parse_addr(s):
    """멀티라인 문자열 또는 리스트 → 이메일 주소 리스트 (빈 줄/공백 제거)."""
    if isinstance(s, list):
        return [x.strip() for x in s if x.strip()]
    return [line.strip() for line in s.splitlines() if line.strip()]

TO  = _parse_addr(TO)
CC  = _parse_addr(CC)
BCC = _parse_addr(BCC)

LOG_FILE = Path(__file__).parent / "run_mailer.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)

DATE_PATTERN = re.compile(r"_(\d{6})$", re.IGNORECASE)  # p.stem 기준 (확장자 제거 후)


def parse_date(yymmdd: str) -> date:
    """YYMMDD 문자열을 date 객체로 변환 (26xxxx → 2026xx)."""
    yy, mm, dd = int(yymmdd[:2]), int(yymmdd[2:4]), int(yymmdd[4:6])
    year = 2000 + yy
    return date(year, mm, dd)


def find_xlsx_files(directory: str):
    """
    directory 내 xlsx 파일 중 _YYMMDD 패턴이 있는 파일 목록 반환.
    [(date, Path), ...] 형태, 날짜 오름차순 정렬.
    """
    results = []
    for p in Path(directory).glob("*.xlsx"):
        m = DATE_PATTERN.search(p.stem)
        if m:
            try:
                d = parse_date(m.group(1))
                results.append((d, p))
            except ValueError:
                pass
    results.sort(key=lambda x: x[0])
    return results


def select_file(today_only: bool):
    """
    조건에 따라 발송할 파일 선택.
    today_only=True  → 오늘 날짜 파일 (없으면 None)
    today_only=False → 가장 최신 날짜 파일
    """
    files = find_xlsx_files(ATTACH_DIR)
    if not files:
        log.error(f"파일 없음: {ATTACH_DIR} 에서 _YYMMDD 패턴 xlsx 파일을 찾지 못했습니다.")
        return None, None

    log.info(f"발견된 파일 {len(files)}개:")
    for d, p in files:
        log.info(f"  {d}  {p.name}")

    if today_only:
        today = date.today()
        matched = [(d, p) for d, p in files if d == today]
        if not matched:
            log.warning(f"오늘({today}) 날짜 파일 없음. 발송 중단.")
            return None, None
        file_date, file_path = matched[-1]   # 오늘 날짜 중 마지막
    else:
        file_date, file_path = files[-1]     # 가장 최신

    log.info(f"선택된 파일: {file_path.name}  (날짜: {file_date})")
    return file_date, file_path


# ── Outlook COM 발송 ──────────────────────────────────────────────

def send_via_outlook(to, cc, bcc, subject, body, attach_path):
    try:
        import win32com.client
    except ImportError:
        log.error("win32com 없음. `pip install pywin32` 후 재시도하거나 SEND_METHOD를 'smtp'로 변경하세요.")
        raise

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail    = outlook.CreateItem(0)  # 0 = olMailItem

    mail.To      = "; ".join(to)
    mail.CC      = "; ".join(cc)
    mail.BCC     = "; ".join(bcc)
    mail.Subject = subject
    mail.Body    = body

    mail.Attachments.Add(str(attach_path.resolve()))
    mail.Send()
    log.info("Outlook으로 발송 완료.")


# ── SMTP 발송 ─────────────────────────────────────────────────────

def send_via_smtp(to, cc, bcc, subject, body, attach_path):
    msg = MIMEMultipart()
    msg["From"]    = FROM_ADDR
    msg["To"]      = ", ".join(to)
    msg["CC"]      = ", ".join(cc)
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    with open(attach_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f'attachment; filename="{attach_path.name}"',
    )
    msg.attach(part)

    all_recipients = to + cc + bcc
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(FROM_ADDR, all_recipients, msg.as_string())
    log.info("SMTP으로 발송 완료.")


# ── 메인 ──────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Auto Mailer - xlsx 자동 발송")
    parser.add_argument(
        "--today-only", action="store_true",
        help="오늘 날짜(_YYMMDD)와 일치하는 파일만 발송 (없으면 발송 안 함)"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="파일 선택 결과만 출력, 실제 발송 안 함"
    )
    parser.add_argument(
        "--to", nargs="+", metavar="EMAIL",
        help="수신자 이메일 지정 (공백으로 구분, 스크립트 TO 설정 덮어씀)"
    )
    parser.add_argument(
        "--cc", nargs="+", metavar="EMAIL",
        help="참조 이메일 지정"
    )
    args = parser.parse_args()

    # 커맨드라인 인수로 수신자 덮어쓰기
    if args.to:
        TO[:] = args.to
    if args.cc:
        CC[:] = args.cc

    log.info("=" * 60)
    log.info(f"run_mailer 시작  today-only={args.today_only}  dry-run={args.dry_run}")

    file_date, file_path = select_file(today_only=args.today_only)
    if file_path is None:
        log.info("발송할 파일 없음. 종료.")
        sys.exit(0)

    now_str      = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    subject      = SUBJECT_TEMPLATE.format(filename=file_path.name)
    body         = BODY_TEMPLATE.format(
        filename=file_path.name,
        file_date=str(file_date),
        now=now_str,
    )

    log.info(f"수신자: {TO}")
    log.info(f"제목  : {subject}")
    log.info(f"첨부  : {file_path}")

    if args.dry_run:
        log.info("[DRY-RUN] 실제 발송 생략.")
        print("\n─── DRY-RUN 결과 ───")
        print(f"  파일     : {file_path.name}")
        print(f"  파일 날짜: {file_date}")
        print(f"  수신자   : {TO}")
        print(f"  참조     : {CC}")
        print(f"  제목     : {subject}")
        print(f"  발송방식 : {SEND_METHOD}")
        print("────────────────────")
        return

    try:
        if SEND_METHOD == "outlook":
            send_via_outlook(TO, CC, BCC, subject, body, file_path)
        elif SEND_METHOD == "smtp":
            send_via_smtp(TO, CC, BCC, subject, body, file_path)
        else:
            log.error(f"알 수 없는 SEND_METHOD: {SEND_METHOD!r}. 'outlook' 또는 'smtp'로 설정하세요.")
            sys.exit(1)
        log.info("발송 성공.")
    except Exception as e:
        log.exception(f"발송 실패: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
