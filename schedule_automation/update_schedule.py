"""
update_schedule.py
2026-10-01  Jonghyun Park w/ Claude  — SKIP 판정을 마커 대신 Auto 파일 실제 저장내용으로 전환 (target_is_saved)
                                       + 저장 직후 재검증을 통과했을 때만 마커 기록 + --force 옵션
                                       + 소스의 세로 병합 셀(여러 법인이 Participation 'O' 한 칸을 공유) 값을
                                         병합된 모든 행에 채움 (FILL_MERGED_COLS)
2026-08-11  Jonghyun Park w/ Claude  — win32com 지연 바인딩에서 wb_com.Close() 가 TypeError 로 죽는 문제 방어
(이전 버전 이력은 git history / GitHub Releases 참조 — 헤더에는 최근 2개 항목만 남긴다)

1. 1.고객 법인 일정 파일/ 폴더에서 최신 파일 자동 선택
   - 정렬 기준: 파일명 내 날짜(YYMMDD) → 버전(_vX.XX) → 끝 번호(_2 등) → 메일수신 일시
2. 소스 파일 첫 번째 시트 B3:J(마지막 데이터 행) 값 읽기  ※ 범위는 상단 SRC_* 상수
   - datetime → yyyy-mm-dd 문자열 변환
   - WEEKNUM 수식 셀 → W01 형식 변환
   - FILL_MERGED_COLS 열의 세로 병합 셀은 병합된 모든 행에 같은 값을 채운다
3. Auto 파일에 그 결과가 실제로 저장돼 있는지 확인 → 이미 반영돼 있으면 SKIP
   (마커가 아니라 파일 내용으로 판정 — 아래 '재실행 판정' 참조)
4. Auto 파일의 '고객법인일정파일' 시트 B2:K999 클리어 후 B2부터 값 붙여넣기 (서식 제외)
   ※ 범위는 상단 TGT_* 상수
5. Excel COM 으로 전체 재계산 후 저장 → 저장 직후 재검증 → 통과했을 때만 마커 기록

재실행 판정:
  종전엔 마커(schedule_last_source.txt)가 최신이면 무조건 SKIP 했다. 그런데 저장이 끝난 뒤
  Auto 파일이 외부(열려 있던 Excel / OneDrive 옛 버전 복원)에 의해 되돌려지거나 붙여넣기가 어긋난 채 남으면
  마커만 '처리 완료'인 상태로 굳어 스케줄러가 계속 SKIP 한다.
  → target_is_saved() 가 Auto 파일의 D1 스탬프 + 붙여넣기 영역 값 + 잔재 행을 실제로 대조해 판정한다.
    마커는 기록·경고용으로만 남는다.
  → 내용이 같아도 다시 붙여넣고 싶으면:  python update_schedule.py --force
"""

import re
import sys
import time
import datetime as dt
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import win32com.client

CHANGED_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
NO_FILL      = PatternFill(fill_type=None)


# ── 최신 파일 정렬 키 ────────────────────────────────────────
def mail_stamp_key(date6: str | None, hhmm: str | None) -> int:
    """메일수신 일시(YYMMDD[_HHMM])를 정렬 가능한 정수 하나로 합침.

    YYMMDD*10000 + HHMM → 260722_1432 = 2607221432, 260722(시각없음) = 2607220000
    시각 없는 옛 파일이 같은 날 시각 있는 파일보다 항상 앞(=오래된 것)으로 정렬됨.
    """
    return int(date6 or 0) * 10000 + int(hhmm or 0)


def latest_file_key(f: Path):
    """파일명에서 정렬 키 (문서날짜, HHMM, 버전float, 버전int, 메일수신일시) 반환.

    ※ 마지막 성분(메일수신일시)은 check_mail_attachment.py 가 **같은 파일명이 재수신될 때만**
      덧붙이는 suffix. _YYMMDD → _YYMMDD_HHMM (시각 포함) 으로 확장됐고,
      옛 _YYMMDD 형식도 그대로 파싱되므로 기존 파일 재정렬 문제 없음.

    A형 (8자리 날짜):
      YYYYMMDD_HHMM[_YYMMDD[_HHMM]]  → (doc_date, hhmm, 0, 0, mail)
      YYYYMMDD_vN[_YYMMDD[_HHMM]]    → (doc_date, 0, 0, ver_int, mail)
      YYYYMMDD_YYMMDD[_HHMM]         → (doc_date, 0, 0, 0, mail)
      YYYYMMDD                       → (doc_date, 0, 0, 0, 0)

    B형 (6자리 날짜 + vX.XX 버전):
      _vX.XX_YYMMDD[_YYMMDD_HHMM]    → (doc_date6, 0, ver_float, suffix, mail)
    """
    name = f.stem

    # 메일 suffix 공통 꼬리: _YYMMDD 또는 _YYMMDD_HHMM (둘 다 없어도 됨)
    MAIL_TAIL = r'(?:_(\d{6})(?:_(\d{4}))?)?'

    m8 = re.search(r'(?<!\d)(\d{8})(?!\d)', name)
    if m8:
        doc_date = int(m8.group(1))

        # YYYYMMDD_HHMM[_메일꼬리]: 뒤에 4자리 숫자가 오되 그 직후 숫자 없을 때
        m = re.search(r'(?<!\d)\d{8}_(\d{4})' + MAIL_TAIL + r'(?!\d)', name)
        if m:
            return (doc_date, int(m.group(1)), 0.0, 0, mail_stamp_key(m.group(2), m.group(3)))

        # YYYYMMDD_vN[_메일꼬리]
        m = re.search(r'(?<!\d)\d{8}_v(\d+)' + MAIL_TAIL, name)
        if m:
            return (doc_date, 0, 0.0, int(m.group(1)), mail_stamp_key(m.group(2), m.group(3)))

        # YYYYMMDD_YYMMDD[_HHMM]
        m = re.search(r'(?<!\d)\d{8}_(\d{6})(?:_(\d{4}))?(?!\d)', name)
        if m:
            return (doc_date, 0, 0.0, 0, mail_stamp_key(m.group(1), m.group(2)))

        return (doc_date, 0, 0.0, 0, 0)

    # ── B형: _vX.XX_YYMMDD ──
    # ⚠ 끝번호 정규식 `_(\d{1,5})$` 이 메일 suffix 의 시각(_1432)을 버전 끝번호로 오인하므로,
    #   끝의 `_YYMMDD_HHMM`(날짜+시각이 둘 다 있는 형태 = 스크립트가 붙인 것) 을 먼저 떼어낸 뒤 판정.
    #   날짜만 있는 `_YYMMDD` 는 문서날짜일 수 있어 떼지 않음 (종전 동작 유지).
    m_tail   = re.search(r'_(\d{6})_(\d{4})$', name)
    mail_key = mail_stamp_key(m_tail.group(1), m_tail.group(2)) if m_tail else 0
    core     = name[:m_tail.start()] if m_tail else name

    date6   = int(m.group()) if (m := re.search(r'(?<!\d)\d{6}(?!\d)', core)) else 0
    version = float(m.group(1)) if (m := re.search(r'_v(\d+\.\d+)', core)) else 0.0
    suffix  = int(m.group(1)) if (m := re.search(r'_(\d{1,5})$', core)) else 0
    return (date6, 0, version, suffix, mail_key)


# ── 경로 설정 ────────────────────────────────────────────────
BASE = Path(
    r"C:\Users\user_name\OneDrive - company_name"
    r"\Project_team_name - 1 company_name - 02 part_name"
    r"\part_name\2026\# CAMPAIGN_PROJECTS\02. CAMPAIGN NAME\02. SCHEDULE"
)

SOURCE_FOLDER    = BASE / "1.고객 법인 일정 파일"
TARGET_SHEET     = "고객법인일정파일"
LAST_SOURCE_FILE = BASE / "schedule_last_source.txt"  # 마커: Auto 파일과 같은 폴더 (프로젝트별 독립 관리. 다른 캠페인으로 fork 시 BASE만 교체하면 마커도 따라감)

# ── 읽기 / 붙여넣기 범위 ─────────────────────────────────────
# 소스 일정표의 열 구성이 바뀌면 여기 숫자만 고치면 된다 (함수 본문엔 숫자를 박지 않는다).
SRC_MIN_COL   = 2    # B (Global)
SRC_MAX_COL   = 10   # J — 1세트 구성 기준
SRC_MIN_ROW   = 3    # 소스 데이터 시작 행
TGT_START_ROW = 2    # 타겟 붙여넣기 시작 행 (소스 행 - 1)
TGT_MAX_ROW   = 999  # 클리어 범위 하단
TGT_MIN_COL   = 2    # B
TGT_MAX_COL   = 11   # K

# 전후 비교(노란 음영) 대상 — {src_data 인덱스: 타겟 열번호},  타겟 열번호 = 인덱스 + 2
COMPARE = {
    3: 5,   # E Participation
    4: 6,   # F Starts at
    6: 8,   # H Ends at
}

# ※ 소스가 B2B/B2C 2세트(B~N)로 확장된 일정표라면 위 값을 아래로 바꿔 쓴다:
#     SRC_MAX_COL = 14 (N) / TGT_MAX_COL = 14
#     COMPARE 에 { 8: 10 (J 2번째 Starts at), 10: 12 (L 2번째 Ends at) } 추가
#   (포맷이 바뀐 직후 1회는 전후 비교에서 성격이 다른 열끼리 비교돼 음영이 과하게 찍힐 수 있음)

# ── 세로 병합 셀 채움 ────────────────────────────────────────
# 소스에서 여러 행이 한 칸으로 병합돼 있으면(예: 법인 3곳이 Participation 'O' 한 칸을 공유)
# 값은 맨 위 행에만 있고 나머지는 빈칸으로 읽힌다 → 이 열들은 병합된 모든 행에 같은 값을 채운다 (소스 열 번호).
# ※ B열(Region)도 그룹 단위로 병합돼 있지만 '그룹 시작 행에만 값' 인 기존 형태를 유지하려고 넣지 않았다.
FILL_MERGED_COLS = [5]   # E열 Participation


def merged_fill_map(ws) -> dict:
    """{(행, 열): 병합 영역의 왼쪽 위 값} — FILL_MERGED_COLS 열 안에서 세로로만 병합된 영역이 대상.

    data_only 로 연 시트에서 불러야 값(수식 결과)이 들어온다. 가로 병합(헤더 등)은 건드리지 않는다.
    """
    fill = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_col != mr.max_col or mr.min_col not in FILL_MERGED_COLS or mr.min_row == mr.max_row:
            continue
        top = ws.cell(row=mr.min_row, column=mr.min_col).value
        if top is None:
            continue
        for r in range(mr.min_row + 1, mr.max_row + 1):
            fill[(r, mr.min_col)] = top
    return fill


# ── 재실행 판정 설정 ─────────────────────────────────────────
STAMP_ROW         = 1      # 붙여넣은 소스 파일명을 기록·대조하는 셀 (D1) — 재실행 판정의 1차 키
STAMP_COL         = 4
FORCE_FLAG        = "--force"   # 이 인자를 주면 판정과 무관하게 다시 붙여넣는다
# 저장 직후 재검증 재시도 — Excel 이 막 저장하고 핸들을 놓기 전 찰나에 읽으면 PermissionError 가 난다
VERIFY_RETRIES    = 3
VERIFY_WAIT_SEC   = 2


# ── 재실행 판정 ──────────────────────────────────────────────
def _norm(v):
    """저장값·소스값 비교용 정규화 — datetime → date, 문자열 줄바꿈 통일·앞뒤 공백 제거, 빈 문자열 == 빈칸."""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, str):
        v = v.replace("\r\n", "\n").strip()
        return v or None
    return v


def target_is_saved(output_file: Path, source_name: str, src_data: list) -> tuple[bool, str]:
    """Auto 파일에 현재 소스의 붙여넣기 결과가 **실제로 저장돼 있는지** 확인.

    (True, "")    = 이미 반영됨 → 실행 불필요
    (False, 사유) = 미반영·유실 → 실행 필요

    마커(로그)를 믿지 않고 파일 내용으로만 판정한다. 저장이 끝난 뒤 외부(Excel/OneDrive)가
    파일을 되돌려도 다음 실행이 스스로 알아채고 재처리하게 하는 게 목적.
    """
    try:
        wb = openpyxl.load_workbook(output_file, data_only=True, read_only=True)
    except Exception as e:                  # 사용 중·손상 등 — 못 읽으면 '미반영' 으로 보고 진행
        return False, f"Auto 파일을 읽을 수 없음 ({type(e).__name__}: {e})"

    try:
        if TARGET_SHEET not in wb.sheetnames:
            return False, f"'{TARGET_SHEET}' 시트 없음"
        ws = wb[TARGET_SHEET]
        # read_only 모드는 ws.cell() 랜덤 접근이 느리므로 한 번에 훑어 dict 로 받는다
        rows = {r_idx: row for r_idx, row in enumerate(
            ws.iter_rows(min_row=1, max_row=TGT_MAX_ROW,
                         min_col=TGT_MIN_COL, max_col=TGT_MAX_COL,
                         values_only=True), start=1)}
    finally:
        wb.close()

    def cell(r_idx: int, col: int):
        row = rows.get(r_idx)
        idx = col - TGT_MIN_COL
        return row[idx] if row and idx < len(row) else None

    # 1) D1 소스 파일명 스탬프
    stamp = _norm(cell(STAMP_ROW, STAMP_COL))
    if stamp != source_name:
        return False, f"D1 소스명 불일치 (저장됨 {stamp!r} != 현재 {source_name!r})"

    # 2) 붙여넣기 영역 값 대조 (B2~)
    for r_offset, want_row in enumerate(src_data):
        r_idx = TGT_START_ROW + r_offset
        for c_offset, want in enumerate(want_row):
            got = cell(r_idx, TGT_MIN_COL + c_offset)
            if _norm(got) != _norm(want):
                col = get_column_letter(TGT_MIN_COL + c_offset)
                return False, f"{col}{r_idx} 값 불일치 (저장됨 {got!r} != 소스 {want!r})"

    # 3) 영역 밖 잔재 — 행 수가 줄었는데 옛 행이 남은 경우
    paste_end = TGT_START_ROW + len(src_data) - 1
    for r_idx, row in rows.items():
        if r_idx <= paste_end:              # 스탬프 행(1행) + 붙여넣기 영역
            continue
        if any(_norm(v) is not None for v in row):
            return False, f"{r_idx}행에 옛 데이터 잔재"

    return True, ""


# ── Auto 파일 자동 탐색 ──────────────────────────────────────
auto_files = list(BASE.glob("*Auto*.xlsx"))
if not auto_files:
    raise FileNotFoundError(f"Auto 파일을 찾을 수 없습니다: {BASE}")
output_file = auto_files[0]
print(f"[업데이트 대상] {output_file.name}")

# ── 소스 폴더에서 최신 파일 선택 ────────────────────────────
xlsx_files = sorted(SOURCE_FOLDER.glob("*.xlsx"), key=latest_file_key)
if not xlsx_files:
    raise FileNotFoundError(f"소스 폴더에 xlsx 파일이 없습니다: {SOURCE_FOLDER}")

source_file = xlsx_files[-1]
print(f"[소스 파일] {source_file.name}")

# ── Pass 1: WEEKNUM 수식이 있는 셀 위치 파악 ─────────────────
src_wb_raw = openpyxl.load_workbook(source_file, data_only=False)
src_ws_raw = src_wb_raw.worksheets[0]

weeknum_cells = set()
for row in src_ws_raw.iter_rows(min_row=SRC_MIN_ROW, min_col=SRC_MIN_COL, max_col=SRC_MAX_COL):
    for cell in row:
        if (
            cell.value
            and isinstance(cell.value, str)
            and "WEEKNUM" in cell.value.upper()
        ):
            weeknum_cells.add((cell.row, cell.column))

src_wb_raw.close()
print(f"[WEEKNUM 셀] {len(weeknum_cells)}개 감지")

# ── Pass 2: 실제 값 읽기 ─────────────────────────────────────
src_wb = openpyxl.load_workbook(source_file, data_only=True)
src_ws = src_wb.worksheets[0]
print(f"[소스 시트] {src_ws.title}")

src_fill = merged_fill_map(src_ws)
if src_fill:
    print(f"[병합 셀 채움] {len(src_fill)}칸")

src_data = []
for row in src_ws.iter_rows(min_row=SRC_MIN_ROW, min_col=SRC_MIN_COL, max_col=SRC_MAX_COL):
    if all(cell.value is None for cell in row):
        continue  # 완전 빈 행 스킵
    row_data = []
    for cell in row:
        v = cell.value
        if v is None:
            v = src_fill.get((cell.row, cell.column))   # 세로 병합의 아래쪽 칸
        # WEEKNUM 수식 셀 → W01 형식
        if (cell.row, cell.column) in weeknum_cells and v and isinstance(v, (int, float)):
            v = f"W{int(v):02d}"
        # datetime → date로 변환 (시간 정보 제거, Excel 날짜값 유지)
        elif isinstance(v, dt.datetime):
            v = v.date()
        row_data.append(v)
    src_data.append(row_data)

src_wb.close()
print(f"[읽은 행 수] {len(src_data)}행")

# ── 재실행 판정 — 마커(로그)가 아니라 Auto 파일에 실제 저장된 내용으로 ──
# ※ 마커는 '처리 완료' 기록·경고용으로만 남긴다. 판정에 쓰면 저장이 유실됐을 때 영원히 SKIP 된다.
src_mtime = int(source_file.stat().st_mtime)
current_marker = f"{source_file.name}|{src_mtime}"
marker_says_done = (LAST_SOURCE_FILE.exists()
                    and LAST_SOURCE_FILE.read_text(encoding="utf-8").strip() == current_marker)

if FORCE_FLAG in sys.argv[1:]:
    print(f"[강제 실행] {FORCE_FLAG} - 판정을 건너뛰고 다시 붙여넣습니다")
else:
    saved, reason = target_is_saved(output_file, source_file.name, src_data)
    if saved:
        print(f"[SKIP] Auto 파일에 이미 반영돼 있습니다 ({source_file.name})")
        if not marker_says_done:
            LAST_SOURCE_FILE.write_text(current_marker, encoding="utf-8")
            print("[알림] 내용은 최신이라 마커만 뒤늦게 동기화했습니다.")
        exit(0)
    if marker_says_done:
        print("[경고] 마커는 '처리 완료'인데 Auto 파일 내용은 최신이 아닙니다 - 저장이 유실된 것으로 보고 재실행합니다.")
        print("        (Auto 파일이 Excel 에서 열려 있었거나 OneDrive 가 옛 버전으로 되돌렸을 수 있습니다)")
    print(f"[갱신 필요] {reason}")

# ── 이전 파일 읽기 (전후 비교용) ──────────────────────────────
prev_data = {}
if len(xlsx_files) >= 2:
    prev_file = xlsx_files[-2]
    prev_wb_raw = openpyxl.load_workbook(prev_file, data_only=False)
    prev_ws_raw = prev_wb_raw.worksheets[0]
    prev_weeknum_cells = set()
    for row in prev_ws_raw.iter_rows(min_row=SRC_MIN_ROW, min_col=SRC_MIN_COL, max_col=SRC_MAX_COL):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "WEEKNUM" in cell.value.upper():
                prev_weeknum_cells.add((cell.row, cell.column))
    prev_wb_raw.close()

    prev_wb = openpyxl.load_workbook(prev_file, data_only=True)
    prev_ws = prev_wb.worksheets[0]
    prev_fill = merged_fill_map(prev_ws)
    for row in prev_ws.iter_rows(min_row=SRC_MIN_ROW, min_col=SRC_MIN_COL, max_col=SRC_MAX_COL):
        if all(cell.value is None for cell in row):
            continue
        row_data = []
        for cell in row:
            v = cell.value
            if v is None:
                v = prev_fill.get((cell.row, cell.column))   # 세로 병합의 아래쪽 칸
            if (cell.row, cell.column) in prev_weeknum_cells and v and isinstance(v, (int, float)):
                v = f"W{int(v):02d}"
            elif isinstance(v, dt.datetime):
                v = v.date()
            row_data.append(v)
        subs_key = row_data[1]  # C열
        if subs_key:
            prev_data[subs_key] = row_data
    prev_wb.close()
    print(f"[이전 파일] {prev_file.name} ({len(prev_data)}행 로드)")

# ── 타겟 파일 업데이트 ───────────────────────────────────────
try:
    tgt_wb = openpyxl.load_workbook(output_file)
except PermissionError:
    print(f"[SKIP] 파일이 사용 중입니다. 다음 실행 시 재시도합니다: {output_file.name}")
    exit(0)

if TARGET_SHEET not in tgt_wb.sheetnames:
    raise ValueError(f"'{TARGET_SHEET}' 시트를 찾을 수 없습니다. 시트 목록: {tgt_wb.sheetnames}")

tgt_ws = tgt_wb[TARGET_SHEET]

# D1에 소스 파일명 기록 (재실행 판정의 1차 키 — target_is_saved() 가 이 값을 대조한다)
tgt_ws.cell(row=STAMP_ROW, column=STAMP_COL, value=source_file.name)

# 대상 영역과 겹치는 병합셀 해제 — MergedCell 은 value 설정이 불가(read-only)해서
# 클리어·붙여넣기에서 예외가 난다. 이 영역은 어차피 소스값으로 덮어쓰므로 해제해도 무방.
for rng in list(tgt_ws.merged_cells.ranges):
    if (rng.max_row >= TGT_START_ROW and rng.min_row <= TGT_MAX_ROW
            and rng.max_col >= TGT_MIN_COL and rng.min_col <= TGT_MAX_COL):
        tgt_ws.unmerge_cells(str(rng))

# 값·음영 클리어 (서식 유지)
for row in tgt_ws.iter_rows(min_row=TGT_START_ROW, max_row=TGT_MAX_ROW,
                            min_col=TGT_MIN_COL, max_col=TGT_MAX_COL):
    for cell in row:
        cell.value = None
        cell.fill  = NO_FILL

# B2부터 값 붙여넣기
for r_idx, row_data in enumerate(src_data, start=TGT_START_ROW):
    for c_idx, value in enumerate(row_data, start=TGT_MIN_COL):  # B열=2
        cell = tgt_ws.cell(row=r_idx, column=c_idx, value=value)
        if isinstance(value, dt.date):
            cell.number_format = "YYYY-MM-DD"

# 변경 셀 음영 표시 (대상 = 상단 COMPARE)
if prev_data:
    changed_count = 0
    for r_idx, row_data in enumerate(src_data, start=TGT_START_ROW):
        subs_key = row_data[1]  # C열
        if not subs_key or subs_key not in prev_data:
            continue
        prev_row = prev_data[subs_key]
        for src_idx, tgt_col in COMPARE.items():
            cur_val  = row_data[src_idx] if src_idx < len(row_data) else None
            prv_val  = prev_row[src_idx] if src_idx < len(prev_row) else None
            if cur_val != prv_val:
                tgt_ws.cell(row=r_idx, column=tgt_col).fill = CHANGED_FILL
                changed_count += 1
    print(f"[변경 셀] {changed_count}개 음영 표시")

try:
    tgt_wb.save(output_file)
    tgt_wb.close()
except PermissionError:
    tgt_wb.close()
    print(f"[SKIP] 저장 중 파일이 잠겼습니다. 다음 실행 시 재시도합니다: {output_file.name}")
    exit(0)

# Excel로 열어서 전체 재계산 후 저장 (FILTER/SORT 등 동적 배열 함수 반영)
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False      # Close/Quit 시 저장 확인 팝업 방지 (스케줄러 실행 대비)
try:
    wb_com = excel.Workbooks.Open(str(output_file.resolve()))
    excel.CalculateFull()
    wb_com.Save()
    # ※ gen_py 캐시가 없으면(파이썬 새 버전 설치 직후 등) win32com 이 지연 바인딩으로 동작해
    #   `wb_com.Close` 는 속성 접근만으로 COM 메서드가 실행되고 bool(True) 을 돌려준다
    #   → 이어지는 `()` 가 TypeError: 'bool' object is not callable 로 죽는다.
    #   그 시점엔 이미 저장·닫기가 끝난 상태이므로 TypeError 만 무시한다.
    #   (조기 바인딩이면 아래 호출이 정상 동작 — 양쪽 다 안전)
    #   같은 지연 바인딩 상태에서 AttributeError('Open.Close') 로 나오는 경우도 있다 - 저장은 끝난 뒤라 함께 무시.
    try:
        wb_com.Close(SaveChanges=False)
    except (TypeError, AttributeError):
        pass
finally:
    excel.Quit()

# 저장 직후 재검증 — 마커는 여기를 통과했을 때만 기록한다.
# ※ 저장이 외부(Excel/OneDrive)에 의해 되돌려진 경우를 즉시 드러내기 위한 단계.
# ※ Excel 이 막 놓은 파일을 곧바로 읽으면 PermissionError 가 난다 — 그건 '되돌려짐' 이 아니라 단순 타이밍이므로
#   **읽기 실패 사유일 때만** 잠시 기다렸다 다시 본다. 값 불일치·잔재는 재시도해도 그대로라 즉시 판정한다.
for _ in range(VERIFY_RETRIES):
    ok, why = target_is_saved(output_file, source_file.name, src_data)
    if ok or "읽을 수 없음" not in why:
        break
    time.sleep(VERIFY_WAIT_SEC)
if ok:
    LAST_SOURCE_FILE.write_text(current_marker, encoding="utf-8")
    print(f"[완료] {output_file.name} 저장 완료")
else:
    print(f"[경고] 저장 직후 검증 실패 - {why}")
    print("        다른 프로그램(Excel/OneDrive)이 파일을 되돌렸을 수 있습니다.")
    print("        마커를 기록하지 않았으므로 다음 실행이 다시 처리합니다.")
