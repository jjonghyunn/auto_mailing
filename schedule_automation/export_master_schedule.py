"""
export_master_schedule.py
2026-11-23  Jonghyun Park w/ Claude  - 받은편지함에 메일이 없는 고객 파일(폴더에 직접 넣은 수기본)도 처리 (NO_MAIL_AS_MANUAL). Outlook 조회 실패 회차만 보류.
2026-11-19  Jonghyun Park w/ Claude  - GRO Status (Pre) 와 법인별 Early 비교를 team_note 에 추가 (GRO_PRE_*, phase 1, Early 없으면 phase 2). GRO 열은 헤더 이름으로 찾음.
(이전 버전 이력은 git history / GitHub Releases 참조 — 헤더에는 최근 2개 항목만 남긴다)

고객 법인 일정 파일(`1.고객 법인 일정 파일/yymmdd_hhmm_*.xlsx`)을 직접 읽어 site_code 별 Start/End 를 계산하고
`campaign, site_code, phase, start_date, end_date, source_subs, note, team_note` 8칼럼 CSV 로 저장한다.
- phase = 0 (Handraiser) · 1 (Pre-sale / Early Access) · 2 (Main). 0·1 행은 참여(O)이고 시작/종료 중 하나라도 있는 site 만 —
  Main 만 있는 site 는 2 행 하나뿐이다. 2 행은 미참여(X)·매칭 실패까지 모든 site 에 항상 나온다. site 안에서 0 → 1 → 2 순서다.
- source_subs = 그 site 의 일정을 가져온 고객 파일 C열(Subs) 원본값 — 검수용. 매칭 실패(일정 없음)면 빈칸.
- note = 같은 고객 파일 행의 Note 칼럼 값 (Auto `고객법인일정파일` N열) 원문. 매칭 실패·Note 없음이면 빈칸.
- team_note = 정제·비교하며 남기는 메모. 해당 없으면 빈칸.
  phase 2 행은 그 site 의 GRO Live/End 가 다르면 `[GRO yymmdd_hhmm] site Live m/d ~ End m/d (법인별 …)`.
  (2026-11-16 꺼짐 — EARLY_GAP_EXTEND_ENABLED) phase 1 행은 Early 와 Main 사이가 떠서 Early Ends 를 Main 직전까지 늘렸으면 `[정제] Main 전 공백을 포함해 Early Ends 변경 (m/d → m/d).`
  phase 1·2 행 둘 다, Early 가 Main 을 침범해 Early Ends 를 Main 전날로 잘랐으면 `[정제] Main 침범분을 잘라 Early Ends 변경 (m/d → m/d).`

- 파일명 = schedule_<yymmdd>_<hhmm>.csv — 고객 파일명 앞 수신일시 prefix(check_mail_attachment.py 가 붙임) 그대로
  → 법인 일정 파일의 메일 시각과 1:1 매칭. 아직 CSV 가 없는 고객 파일 + **최신 고객 파일**을 처리한다
  (최신 회차는 매번 다시 계산해 내용이 바뀌었을 때만 덮어씀. 지난 회차는 그 시점 스냅샷으로 남김 — 다시 뽑으려면 --all).
- 같은 CSV 를 OUTPUT_COPY_DIR(기본 = 스크립트 폴더 하위 3.main_raw_schedule)에도 backup본으로 저장한다. None 이면 backup 끔.
- 법인 → site 매핑 = 같은 폴더 `site_subs_map.csv` (site_code, subs, region, country — 행 순서 = CSV 출력 순서).
  Auto 파일 `태깅기획site_code` 시트에서 뽑은 것. 매핑이 바뀌면 이 CSV 를 고치거나 `--build-map` 으로 다시 뽑는다.
- site 별 계산은 Auto 파일 `고객법인일정파일` 시트 O~S 열 수식 + MASTER J/K 수식을 그대로 옮겼다 (아래 '수식 이식' 참조).
- 미참여는 start 에 'X', 일정 없음(매칭 실패)은 빈칸 — MASTER 와 같은 값.
- 제외 site = 상단 EXCLUDE_SITES (수기) ∪ EXCLUDE_SOURCE_XLSX 의 MASTER 'api수기제외' 칼럼이 O 인 site (자동, 매 회차 읽음).
  CSV 와 --check 대조 양쪽에서 뺀다. 제외 소스는 공유모드로 읽어 다른 프로세스가 잡고 있어도 읽힌다.
  그래도 못 읽으면 마지막으로 읽은 목록(EXCLUDE_CACHE_FILE)을 쓰고, 캐시도 없으면 그 회차는 SKIP (제외가 빠진 CSV 를 쓰지 않도록).
- 팀원 발신 제외: 고객 파일을 보낸 메일의 발신 주소가 TEAM_SENDER_DOMAINS 이면 그 파일은 없는 것으로 본다 (CSV 를 만들지 않고 최신 회차에서도 빠짐).
  발신자는 수신일시 prefix + 첨부명으로 Outlook 받은편지함에서 찾고 SENDER_CACHE_FILE 에 남긴다 (새 파일일 때만 Outlook 조회).
  받은편지함에 메일이 없는 파일은 폴더에 직접 넣은 수기 파일로 보고 처리한다 (NO_MAIL_AS_MANUAL, 캐시에 manual 로 기록).
  Outlook 미연결 등으로 조회 자체가 실패한 새 파일만 그 회차에 보류 — CSV 가 이미 있는 지난 회차는 그대로 둔다.

옵션:
  --all        : CSV 가 이미 있는 고객 파일도 다시 계산 (내용이 다르면 덮어씀)
  --check      : 최신 고객 파일 결과를 Auto MASTER 와 site 별로 대조 (Auto D1 스탬프가 같은 파일일 때만)
  --build-map  : Auto 파일 `태깅기획site_code` 시트에서 site_subs_map.csv 를 다시 만든다
"""

import csv
import io
import re
import sys
import zipfile
import datetime as dt
from pathlib import Path
import openpyxl
import pywintypes
import win32com.client
import win32file

# ════ 사용자가 바꿔야 하는 부분 ════

# ─── 캠페인 식별자 (CSV campaign 칼럼값) ───
CAMPAIGN_ID = "campaign_id"

# ─── 수기 제외 site (한 줄에 site_code 하나, 대소문자 무시. 빈 줄·# 주석 무시) ───
EXCLUDE_SITES = """
"""

# ─── 자동 제외 소스 (이 파일 MASTER 시트의 'api수기제외' 칼럼이 O 인 site 를 매 회차 제외) ───
# None 이면 자동 제외 끔 (EXCLUDE_SITES 만 사용)
EXCLUDE_SOURCE_XLSX = Path(
    r"C:\Users\user_name\OneDrive - company_name"
    r"\Project_team_name - 1 company_name - 02 part_name"
    r"\part_name\2026\# CAMPAIGN_PROJECTS\02. CAMPAIGN NAME\02. SCHEDULE"
    r"\26_Schedule(Auto).xlsx"
)
EXCLUDE_SOURCE_SHEET      = "MASTER"
EXCLUDE_SOURCE_HEADER_ROW = 4
EXCLUDE_SOURCE_SITE_COL   = "G"            # Site Code
EXCLUDE_SOURCE_FLAG_COL   = "N"            # api수기제외 (2026-11-05 Handraiser H/I 추가로 L → N)
EXCLUDE_SOURCE_FLAG_HDR   = "api수기제외"   # 헤더가 이 값이 아니면 구조가 바뀐 것 → 그 회차 SKIP
EXCLUDE_SOURCE_FLAG_MARK  = "O"            # 대소문자 무시
# 제외 소스를 읽을 때마다 site 목록을 여기 저장 — 다음 회차에 소스를 못 읽으면 이 목록으로 CSV 를 만든다 (None 이면 끔 → 못 읽으면 SKIP)
EXCLUDE_CACHE_FILE = Path(__file__).resolve().parent / "_exclude_source_cache.txt"

# ─── 팀원 발신 메일 제외 ───
# 고객 파일을 보낸 메일의 발신 주소 도메인이 이 중 하나면 팀원으로 보고 처리하지 않는다 (빈 리스트면 끔)
TEAM_SENDER_DOMAINS = ["company_name.com"]
MAIL_STORE_NAME     = "part_name mailbox"    # check_mail_attachment.py 의 STORE_NAME 과 같은 메일함
SENDER_CACHE_FILE   = Path(__file__).resolve().parent / "_source_sender_cache.csv"   # file, sender_name, sender_addr
# 받은편지함에 메일이 없는 파일(폴더에 직접 넣은 수기 파일)을 처리할지. Outlook 조회 자체가 실패한 회차는 여전히 보류한다
NO_MAIL_AS_MANUAL   = True
MANUAL_SENDER_NAME  = "폴더 직접(메일 없음)"
MANUAL_SENDER_ADDR  = "manual"

# ─── 경로 ───
BASE = Path(
    r"C:\Users\user_name\OneDrive - company_name"
    r"\Project_team_name - 1 company_name - 02 part_name"
    r"\part_name\2026\# CAMPAIGN_PROJECTS\02. CAMPAIGN NAME\02. SCHEDULE"
)
SOURCE_FOLDER   = BASE / "1.고객 법인 일정 파일"
OUTPUT_DIR      = BASE / "3.main_raw_schedule"
OUTPUT_COPY_DIR = Path(__file__).resolve().parent / "3.main_raw_schedule"   # backup본 — 원본 OUTPUT_DIR 과 같은 CSV 를 로컬 코드 폴더에도 저장 (None 이면 끔)
OUTPUT_PREFIX   = "schedule_"
OUTPUT_ENCODING = "utf-8-sig"          # Excel 에서 바로 열리게 BOM 포함
SITE_MAP_CSV    = Path(__file__).resolve().parent / "site_subs_map.csv"

# ─── GRO ↔ 법인별 일정 날짜 차이 → phase 2 team_note 에 기록 (update_schedule.py 의 GRO_NOTE_* 와 같은 규칙) ───
# 그 고객 파일보다 나중에 도착한 GRO Status 중 (다음 고객 파일 도착 전) 최신 1개와 site 별 Main Starts/Ends 를 비교한다.
#   - GRO 칸이 비어 있으면 그 칸은 비교하지 않는다. 글자가 섞인 칸은 m/d 부분만 (연도 = GRO 수신 prefix 연도).
#   - GRO 폴더 `<GRO파일명>_site_dates.csv` 캐시가 있으면 그걸 읽는다 (update_schedule.py 가 만든다 — 없으면 여기서 만든다).
GRO_NOTE_ENABLED   = True
GRO_FOLDER         = BASE / "2.GRO Status"
GRO_NAME_KEYS      = ["gro", "status"]          # 파일명(소문자)에 모두 포함
GRO_EXTS           = (".xlsx", ".xlsm")
GRO_EXCLUDE_SUFFIX = "_일정"                     # update_schedule.py 가 만든 정제 사본 — GRO 원본 아님
GRO_SHEET          = "GRO Status"
GRO_SUBHEADER_ROWS = 1
# 열은 **헤더 이름**으로 찾는다 (2026-11-19 고정 열 F/K/M → 헤더). 헤더행 = Site Code 가 있는 행, 그 아래 1행은 보조 헤더
GRO_HDR_SITE       = "Site Code"
GRO_HDR_LIVE       = "Live date"                # GRO Status: Live date (Local) / Pre: 'Pre-sale Date' 아래 보조 헤더
GRO_HDR_END        = "End date"
GRO_HDR_SCAN_ROWS  = 20
GRO_NOTE_FORMAT    = "[GRO {stamp}] {site} Live {g_start} ~ End {g_end} (법인별 {c_start} ~ {c_end})"
GRO_NOTE_DATE_FMT  = "%m/%d"
GRO_SITE_DATES_SUFFIX = "_site_dates"
# ─── GRO Status (Pre) ↔ 법인별 Early (update_schedule.py 의 GRO_PRE_* 와 같은 규칙 — 설명은 그쪽 주석 참조) ───
# Note 줄은 phase 1(Early) 행 team_note 에, Early 행이 없으면 phase 2 행에 붙인다.
GRO_PRE_ENABLED        = True
GRO_PRE_SHEET          = "GRO Status (Pre)"
GRO_PRE_HDR_URL        = "Pre-sale URL"
GRO_PRE_HDR_REMARK     = "Remark"
# Note 칸은 읽지 않는다 (2026-11-19) - 법인별 일정 Note 를 옮겨 적은 내용이라 GRO 조사 근거가 아님. Remark 와 날짜 칸만 쓴다
GRO_PRE_HR_URL_KEYS    = ("register", "handraiser", "reserve")
GRO_PRE_EARLY_TEXT_KEYS = ("early", "pre-sale")
GRO_PRE_HR_TEXT_KEYS   = ("handraiser", "hand-raiser", "teaser")
GRO_PRE_MAIN_TEXT_KEYS = ("메인 캠페인", "main campaign", "메인 페이지")
GRO_PRE_ROWS_SUFFIX    = "_pre_rows"
GRO_PRE_NOTE_FORMAT    = "[GRO(Pre) {stamp}] {site} {reason}: GRO Pre-sale {g} / 법인별 Early {c}{main}"
GRO_PRE_MAIN_FORMAT    = " (Main 시작 GRO {gm} / 법인별 {cm})"

# ─── Early ↔ Main 공백 정제 (update_schedule.py 의 EARLY_GAP_* 와 같은 규칙) ───
# Early Ends 다음 날이 Main Starts 가 아니고 사이가 뜨면 Early Ends 를 Main Starts 전날로 늘린다.
# 2026-11-16 사용자 요청으로 끔 — 고객 Early Ends 를 그대로 두고 공백을 남긴다.
EARLY_GAP_EXTEND_ENABLED = False
EARLY_GAP_NOTE_FORMAT  = "[정제] Main 전 공백을 포함해 Early Ends 변경 ({old} → {new})."
EARLY_GAP_DATE_FMT     = "%m/%d"
# Early Ends 가 Main Starts 와 같거나 뒤면 Main Starts 전날로 자른다. 기록은 phase 1·2 team_note 둘 다.
EARLY_OVERLAP_TRIM_ENABLED = True
EARLY_OVERLAP_NOTE_FORMAT  = "[정제] Main 침범분을 잘라 Early Ends 변경 ({old} → {new})."

# ─── 고객 파일 읽기 (update_schedule.py 와 같은 규칙 — Auto 시트에 붙는 모양을 그대로 재현) ───
SRC_FIRST_ROW    = 3      # 헤더 포함 읽기 시작 행
SRC_FIRST_COL    = 2      # B열
SRC_LAST_COL     = 15     # O열
NOTE_HEADER      = "Note"
NOTE_TARGET_COL  = 14     # N열 — Auto `고객법인일정파일` 의 Note 자리
CUSTOMER_NAME_KEYS = ["법인별", "일정"]   # 고객 파일로 인정할 파일명 키워드 (모두 포함) — 같은 메일의 Home KV 계획 등 제외
# 열 매핑 — Auto B~O 각 칸에 넣을 소스 열. 아래 헤더 자동 감지가 실패한 칸만 이 값을 쓴다 (기본값 = 1029_v2 배치).
#                     B  C  D    E              F       G    H     I    J        K    L      M    N     O
#                  Global Subs PIC Participation Starts  주차  Ends  주차  E.Starts 주차  E.Ends 주차  Note  Days
CUSTOMER_SRC_COLS = ["B", "C", "D", "E", "F", "G", "H", "I", "K", "L", "M", "N", "S", "J"]
MAIN_KEYWORD      = "main"   # 기간 그룹 헤더 — 이 셀 아래의 Starts/Ends 를 쓴다 (대소문자 무시, 단어 일치)
EARLY_KEYWORDS    = ["early", "pre-sale"]   # Pre-sale (Early Access etc) 그룹 헤더 — 없으면 Early 칸은 빈칸
HANDRAISER_KEYWORDS = ["handraiser"]     # Handraiser 그룹 헤더 — 없으면 phase 0 행을 만들지 않는다
START_HEADER      = "start"  # 이 글자로 시작하는 헤더 (Starts / Starts at)
END_HEADER        = "end"
DAYS_HEADER       = "days"
HEADER_SCAN_ROWS  = 8        # 시트 1행부터 몇 행까지 헤더로 보고 찾을지
FILL_MERGED_COLS = [5]    # E열 Participation 세로 병합 채움
SUBS_HEADER      = "Subs" # C열이 이 값인 행 = 헤더행. 데이터는 그 다음 행부터 (Auto 수식 범위 C4:C200 과 같음)

# ─── Auto 파일 (--check 대조 / --build-map 에만 사용) ───
AUTO_FILE_NAME    = "26_Schedule(Auto).xlsx"
MASTER_SHEET      = "MASTER"
MASTER_HEADER_ROW = 4
MASTER_COLS       = {"site": "G", "early_start": "J", "early_end": "K", "start": "L", "end": "M"}   # H/I = Handraiser (--check 대상 아님)
STAMP_SHEET       = "고객법인일정파일"
STAMP_CELL        = "D1"
TAG_SHEET         = "태깅기획site_code"
TAG_HEADER_ROW    = 3
TAG_COLS          = {"no": "B", "region": "C", "subs": "D", "country": "E", "site": "F"}

# ─── 특수 표기 (Auto 수식에 박혀 있던 값) ───
HQ_SUBS_KR    = "한총"    # 고객 파일 C열 법인명이 이 값(앞부분)이면 본사(HQ_SUBS_CLEAN) 법인으로 본다
HQ_SUBS_CLEAN = "hq"
PART_YES      = "O"
PART_NO       = "X"

# ─── 출력 칼럼 ───
OUTPUT_COLUMNS = ["campaign", "site_code", "phase", "start_date", "end_date", "source_subs", "note", "team_note"]
PHASE_HANDRAISER = 0      # Handraiser
PHASE_EARLY    = 1        # Pre-sale (Early Access)
PHASE_MAIN     = 2        # Main (본 캠페인 기간)
MAP_COLUMNS    = ["site_code", "subs", "region", "country"]
DATE_FORMAT    = "%Y-%m-%d"

# ════ 내부 사용 ════

MAIL_PREFIX_PAT = re.compile(r'^(\d{6})_(\d{4})_')   # check_mail_attachment.py 저장명 규칙
GRO_TEXT_DATE_PAT = re.compile(r"(\d{1,2})\s*/\s*(\d{1,2})")   # 글자가 섞인 GRO 날짜 칸에서 m/d 만
GRO_PRE_MD_PAT    = re.compile(r"(?<![\d/])(\d{1,2})/(\d{1,2})(?![\d/])")   # Remark 안 m/d (yyyy/mm/dd 제외)
FLAG_ALL, FLAG_CHECK, FLAG_BUILD_MAP = "--all", "--check", "--build-map"
SENDER_COLUMNS = ["file", "sender_name", "sender_addr"]
OUTLOOK_INBOX  = 6                                    # olFolderInbox


def _col_idx(letter: str) -> int:
    return openpyxl.utils.column_index_from_string(letter)


def _fmt(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime(DATE_FORMAT)
    if v is None:
        return ""
    return str(v).strip()


def _ci(v) -> str:
    """Excel COUNTIF/MATCH 처럼 대소문자 무시 비교용 키."""
    return "" if v is None else str(v).casefold()


def _blank(v) -> bool:
    return v is None or (isinstance(v, str) and v == "")


# ── 고객 파일 → Auto `고객법인일정파일` 시트에 붙는 행 (B~O) ──────────────
def _group_cols(hdr: dict, merged: list, max_col: int, keywords: list):
    """기간 그룹 헤더(keywords 중 하나가 단어로 든 셀) 아래의 (그룹 위치, Starts 열, Ends 열). 그룹이 없으면 (None, None, None)."""
    pat = re.compile("|".join(rf"\b{re.escape(k)}\b" for k in keywords))
    grp = next(((r, c) for (r, c), v in sorted(hdr.items()) if pat.search(v)), None)
    if grp is None:
        return None, None, None
    r0, c0 = grp
    span_end = next((m.max_col for m in merged if m.min_row <= r0 <= m.max_row and m.min_col == c0), None)
    if span_end is None:
        span_end = min([c for (r, c) in hdr if r == r0 and c > c0], default=max_col + 1) - 1
    below = sorted((c, v) for (r, c), v in hdr.items() if r > r0 and c0 <= c <= span_end)
    start_col = next((c for c, v in below if v.startswith(START_HEADER)), None)
    end_col = next((c for c, v in below if v.startswith(END_HEADER)), None)
    return grp, start_col, end_col


def detect_col_map(ws) -> list:
    """헤더에서 Main·Pre-sale 아래 Starts/Ends, Note, Days 위치를 찾아 CUSTOMER_SRC_COLS 형태의 매핑을 만든다 (update_schedule.py 와 같은 규칙).
    Pre-sale 그룹이 없는 옛 파일은 Early 4칸을 None(빈칸)으로 둔다."""
    cols = list(CUSTOMER_SRC_COLS)
    hdr = {}
    for row in ws.iter_rows(min_row=1, max_row=min(HEADER_SCAN_ROWS, ws.max_row)):
        for c in row:
            if isinstance(c.value, str) and c.value.strip():
                hdr[(c.row, c.column)] = c.value.strip().lower()
    merged = list(ws.merged_cells.ranges)
    _, start_col, end_col = _group_cols(hdr, merged, ws.max_column, [MAIN_KEYWORD])
    early, e_start, e_end = _group_cols(hdr, merged, ws.max_column, EARLY_KEYWORDS)
    _, h_start, h_end = _group_cols(hdr, merged, ws.max_column, HANDRAISER_KEYWORDS)

    def find(name: str):
        hits = sorted(c for (r, c), v in hdr.items() if v == name)
        return next((c for c in hits if end_col and c > end_col), hits[0] if hits else None)

    note_col, days_col = find(NOTE_HEADER.lower()), find(DAYS_HEADER)
    letter = openpyxl.utils.get_column_letter
    for i, col in ((4, start_col), (6, end_col), (8, e_start), (10, e_end)):
        if col:
            cols[i], cols[i + 1] = letter(col), letter(col + 1)
    if note_col:
        cols[12] = letter(note_col)
    if days_col:
        cols[13] = letter(days_col)
    missing = [k for k, v in {"Starts": start_col, "Ends": end_col, "Note": note_col, "Days": days_col}.items() if not v]
    if missing:
        print(f"[경고] 헤더 감지 실패 {missing} - CUSTOMER_SRC_COLS 값 사용")
    if not early:
        cols[8:12] = [None] * 4
        print(f"[알림] 헤더에 Pre-sale 그룹({'/'.join(EARLY_KEYWORDS)})이 없어 Early 칸은 비워 둡니다")
    elif not (e_start and e_end):
        print("[경고] Pre-sale Starts/Ends 감지 실패 - CUSTOMER_SRC_COLS 값 사용")
    return cols + [letter(h_start) if h_start else None, letter(h_end) if h_end else None]   # Auto B~O 뒤에 Handraiser 2칸


def read_customer_rows(src: Path) -> list:
    """update_schedule.py 의 읽기 규칙(빈 행 스킵 / E열 병합 채움 / 헤더 감지 열 매핑)을 따라 Auto B~O 모양의 행 리스트를 만든다."""
    wb = openpyxl.load_workbook(src, data_only=True)
    try:
        ws = wb.worksheets[0]
        cols = detect_col_map(ws)
        idx = [_col_idx(c) - SRC_FIRST_COL if c else None for c in cols]
        last_col = max(_col_idx(c) for c in cols if c)
        fill = {}
        for mr in ws.merged_cells.ranges:
            if mr.min_col != mr.max_col or mr.min_col not in FILL_MERGED_COLS or mr.min_row == mr.max_row:
                continue
            top = ws.cell(row=mr.min_row, column=mr.min_col).value
            if top is None:
                continue
            for r in range(mr.min_row + 1, mr.max_row + 1):
                fill[(r, mr.min_col)] = top
        rows = []
        for row in ws.iter_rows(min_row=SRC_FIRST_ROW, min_col=SRC_FIRST_COL, max_col=last_col):
            if all(c.value is None for c in row):
                continue
            vals = []
            for c in row:
                v = c.value if c.value is not None else fill.get((c.row, c.column))
                vals.append(v.date() if isinstance(v, dt.datetime) else v)
            rows.append([vals[i] if i is not None else None for i in idx])
    finally:
        wb.close()

    # 헤더행(C열 'Subs') 다음부터가 데이터 — Auto 수식이 보는 C4:C200 과 같은 범위
    hdr = next((i for i, r in enumerate(rows) if isinstance(r[1], str) and r[1].strip() == SUBS_HEADER), None)
    return rows[hdr + 1:] if hdr is not None else rows


# ── 수식 이식 ────────────────────────────────────────────────
# B=0 C=1 D=2 E=3 F=4 G=5 H=6 … J=8 L=10 N=12 (고객 행 인덱스) — read_customer_rows 가 Auto B~O 순서로 재배치해 둔다
IB, IC, IE, IF_, IH = 0, 1, 3, 4, 6
IJ, IL = 8, 10            # Early Starts / Early Ends (Auto 고객 V/W → MASTER J/K)
IHS, IHE = 14, 15         # Handraiser Starts / Ends (Auto B~O 뒤에 붙인 칸)
INOTE = NOTE_TARGET_COL - SRC_FIRST_COL


def ar_clean(row) -> str:
    """O열 AR_clean: Region/Total/빈 행은 '', 하이픈 행(SUBS-XX 처럼 특정 site 지정)도 '' — 그 외 C열의 ' ' 앞부분. '한총' 이면 HQ_SUBS_CLEAN.
    하이픈 행을 남기면 Subs-SITE·Subs 단독 행이 없는 같은 법인의 다른 site 가 그 나라 일정을 받는다 (2026-10-20 수정)."""
    c = row[IC]
    if not _blank(row[IB]) or _blank(c) or c == "Total":
        return ""
    c = str(c)
    if "-" in c:
        return ""
    if " " in c:
        v = c[:c.find(" ")]
    else:
        v = c
    return HQ_SUBS_CLEAN if v == HQ_SUBS_KR else v


def ar_key(subs: str, site: str, c_keys: set, o_to_c: dict, c_first: dict):
    """P열 AR_key: 태깅 site 에 대응하는 고객 C열 값 (3단계 우선순위 + fallback). 못 찾으면 ''."""
    if not site:
        return ""
    base = site.split("_", 1)[0]
    k1 = f"{subs}-{base.upper()}"
    if _ci(k1) in c_keys:
        return k1
    if "_" in site:
        lang = site.split("_", 1)[1][:20]
        k2 = f"{subs} ({base.title()}_{lang.upper()})"
        if _ci(k2) in c_keys:
            return k2
    if _ci(subs) in c_keys:
        return subs
    if "_" in site:
        return ""
    if _ci(subs) in o_to_c:                      # O열(정제값) 매칭 → 그 행의 C 값
        return o_to_c[_ci(subs)]
    if "-" in subs:
        head = subs[:subs.find("-")]
        return c_first.get(_ci(head), "")
    return ""                                    # (수식의 SCA 분기는 '_' 있는 site 에서만 의미가 있어 위에서 이미 '' — 도달 불가)


def filled_early_end(early_end, main_start):
    """Early 와 Main 사이가 뜨면 Main 전날을, 아니면 None 을 돌려준다 (둘 중 하나라도 날짜가 아니면 None)."""
    if not (EARLY_GAP_EXTEND_ENABLED and isinstance(early_end, dt.date) and isinstance(main_start, dt.date)):
        return None
    new_end = main_start - dt.timedelta(days=1)
    return new_end if new_end > early_end else None


def trimmed_early_end(early_start, early_end, main_start):
    """Early 가 Main 을 침범하면 Main 전날을, 아니면 None 을 돌려준다.
    자르면 Early 가 없어지는 경우(Main 전날 < Early Starts)는 자르지 않는다."""
    if not (EARLY_OVERLAP_TRIM_ENABLED and isinstance(early_end, dt.date) and isinstance(main_start, dt.date)):
        return None
    new_end = main_start - dt.timedelta(days=1)
    if new_end >= early_end or (isinstance(early_start, dt.date) and new_end < early_start):
        return None
    return new_end


def compute_sites(cust_rows: list, site_map: list) -> list:
    """[(site, phase, start, end, source_subs, note, team_note)] — phase 2 start/end 는 MASTER L/M, phase 1 은 MASTER J/K 와 같은 값, phase 0 은 고객 파일 Handraiser,
    source_subs·note 는 매칭된 고객 행의 C열 원본·Note(N열) 값, team_note 는 phase 1 의 Early 연장·phase 1·2 의 Early 자름 기록 (GRO 줄은 apply_gro_notes 가 더한다)."""
    c_first, o_to_c = {}, {}
    for r in cust_rows:
        if not _blank(r[IC]):
            c_first.setdefault(_ci(r[IC]), r)
        o = ar_clean(r)
        if o:
            o_to_c.setdefault(_ci(o), r[IC])
    c_keys = set(c_first)
    c_row = {k: r for k, r in c_first.items()}

    out = []
    for m in site_map:
        key = ar_key(m["subs"], m["site_code"], c_keys, o_to_c,
                     {k: r[IC] for k, r in c_first.items()})
        start = end = e_start = e_end = h_start = h_end = src_subs = note = e_line = o_line = ""
        if str(key).strip():
            r = c_row.get(_ci(key))
            src_subs = r[IC] if r is not None else ""
            note = r[INOTE] if r is not None and len(r) > INOTE else ""
            v = str(r[IE]).upper() if r is not None and r[IE] is not None else ""
            part = PART_YES if "O" in v else (PART_NO if "X" in v else "")
            if part == PART_NO:
                start = PART_NO
            elif part == PART_YES and r is not None:
                start, end = r[IF_], r[IH]
                e_start, e_end = r[IJ], r[IL]     # 미참여(X)면 MASTER J/K 도 빈칸
                h_start, h_end = r[IHS], r[IHE]
                new_end = filled_early_end(e_end, start)
                if new_end:
                    e_line = EARLY_GAP_NOTE_FORMAT.format(old=e_end.strftime(EARLY_GAP_DATE_FMT),
                                                          new=new_end.strftime(EARLY_GAP_DATE_FMT))
                    e_end = new_end
                    print(f"[Early 연장] {m['site_code']} ({src_subs}): {e_line}")
                new_end = trimmed_early_end(e_start, e_end, start)
                if new_end:
                    o_line = EARLY_OVERLAP_NOTE_FORMAT.format(old=e_end.strftime(EARLY_GAP_DATE_FMT),
                                                              new=new_end.strftime(EARLY_GAP_DATE_FMT))
                    e_end = new_end
                    print(f"[Early 자름] {m['site_code']} ({src_subs}): {o_line}")
        if not (_blank(h_start) and _blank(h_end)):
            out.append((m["site_code"], PHASE_HANDRAISER, h_start, h_end, src_subs, note, ""))
        if not (_blank(e_start) and _blank(e_end)):
            out.append((m["site_code"], PHASE_EARLY, e_start, e_end, src_subs, note, e_line or o_line))
        out.append((m["site_code"], PHASE_MAIN, start, end, src_subs, note, o_line))
    return out


# ── GRO ↔ 법인별 일정 날짜 차이 ────────────────────────────────
def _prefix(f: Path) -> str:
    m = MAIL_PREFIX_PAT.match(f.name)
    return f"{m.group(1)}_{m.group(2)}" if m else ""


def gro_files() -> list:
    if not GRO_FOLDER.exists():
        return []
    return sorted((f for f in GRO_FOLDER.iterdir()
                   if f.is_file() and f.suffix.lower() in GRO_EXTS and not f.name.startswith("~$")
                   and not f.stem.endswith(GRO_EXCLUDE_SUFFIX) and MAIL_PREFIX_PAT.match(f.name)
                   and all(k in f.name.lower() for k in GRO_NAME_KEYS)), key=_prefix)


def _gro_cell_date(v, year: int):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and (m := GRO_TEXT_DATE_PAT.search(v)):
        try:
            return dt.date(year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None
    return None


def _cell(row, i):
    return row[i] if i is not None and i < len(row) else None


def _hdr_norm(v) -> str:
    return re.sub(r"\s+", " ", str(v)).strip().casefold() if v is not None else ""


def _gro_header_cols(rows: list, names: dict, sheet: str):
    """(헤더행 index, {key: 열 index}) — 헤더행 = GRO_HDR_SITE 가 있는 행. 이름은 헤더행 → 바로 아래 보조 헤더행 순으로 찾는다."""
    want_site = _hdr_norm(GRO_HDR_SITE)
    hdr = next((i for i, r in enumerate(rows[:GRO_HDR_SCAN_ROWS]) if any(_hdr_norm(v) == want_site for v in r)), None)
    if hdr is None:
        raise ValueError(f"'{sheet}' 헤더행을 찾을 수 없음 ('{GRO_HDR_SITE}' 없음)")
    out = {}
    for key, name in names.items():
        n = _hdr_norm(name)
        for r in rows[hdr:hdr + 2]:
            j = next((j for j, v in enumerate(r) if _hdr_norm(v) == n), None)
            if j is not None:
                out[key] = j
                break
        else:
            raise ValueError(f"'{sheet}' 에서 헤더 '{name}' 를 찾을 수 없음")
    return hdr, out


def gro_site_dates(gro_file: Path) -> dict:
    """{site(소문자): (Live, End)} — 캐시 CSV(update_schedule.py 와 같은 형식)가 원본보다 새것이면 재사용."""
    cache = gro_file.with_name(f"{gro_file.stem}{GRO_SITE_DATES_SUFFIX}.csv")
    iso = lambda s: dt.date.fromisoformat(s) if s else None
    if cache.exists() and cache.stat().st_mtime >= gro_file.stat().st_mtime:
        with cache.open(encoding="utf-8-sig", newline="") as f:
            return {r["site"]: (iso(r["live"]), iso(r["end"])) for r in csv.DictReader(f)}
    year = 2000 + int(_prefix(gro_file)[:2])
    wb = openpyxl.load_workbook(gro_file, data_only=True, read_only=True, keep_vba=False)
    try:
        rows = list(wb[GRO_SHEET].iter_rows(values_only=True))
    finally:
        wb.close()
    hdr, ci = _gro_header_cols(rows, dict(site=GRO_HDR_SITE, start=GRO_HDR_LIVE, end=GRO_HDR_END), GRO_SHEET)
    out = {}
    for r in rows[hdr + 1 + GRO_SUBHEADER_ROWS:]:
        s = _cell(r, ci["site"])
        if isinstance(s, str) and s.strip():
            out[s.strip().lower()] = (_gro_cell_date(_cell(r, ci["start"]), year), _gro_cell_date(_cell(r, ci["end"]), year))
    tmp = cache.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["site", "live", "end"])
        for s, (a, b) in out.items():
            w.writerow([s, a.isoformat() if a else "", b.isoformat() if b else ""])
    tmp.replace(cache)
    return out


def gro_pre_rows(gro_file: Path) -> dict:
    """{site(소문자): [(kind 'hr'|'early', start, end, text)]} — update_schedule.py 와 같은 캐시 CSV 를 같이 쓴다."""
    cache = gro_file.with_name(f"{gro_file.stem}{GRO_PRE_ROWS_SUFFIX}.csv")
    iso = lambda s: dt.date.fromisoformat(s) if s else None
    out = {}
    if cache.exists() and cache.stat().st_mtime >= gro_file.stat().st_mtime:
        with cache.open(encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                out.setdefault(r["site"], []).append((r["kind"], iso(r["start"]), iso(r["end"]), r["text"]))
        return out
    year = 2000 + int(_prefix(gro_file)[:2])
    wb = openpyxl.load_workbook(gro_file, data_only=True, read_only=True, keep_vba=False)
    try:
        if GRO_PRE_SHEET not in wb.sheetnames:
            raise ValueError(f"'{GRO_PRE_SHEET}' 시트 없음")
        rows = list(wb[GRO_PRE_SHEET].iter_rows(values_only=True))
    finally:
        wb.close()
    hdr, ci = _gro_header_cols(rows, dict(site=GRO_HDR_SITE, start=GRO_HDR_LIVE, end=GRO_HDR_END, url=GRO_PRE_HDR_URL,
                                          remark=GRO_PRE_HDR_REMARK), GRO_PRE_SHEET)
    site = None
    for r in rows[hdr + 1 + GRO_SUBHEADER_ROWS:]:
        s = _cell(r, ci["site"])
        if isinstance(s, str) and s.strip():
            site = s.strip().lower()
        a, b = _gro_cell_date(_cell(r, ci["start"]), year), _gro_cell_date(_cell(r, ci["end"]), year)
        if site is None or not (a or b):
            continue
        url = str(_cell(r, ci["url"]) or "").lower()
        text = str(_cell(r, ci["remark"]) or "").strip()
        low = text.lower()
        if any(k in url for k in GRO_PRE_HR_URL_KEYS):
            kind = "hr"
        elif any(k in low for k in GRO_PRE_EARLY_TEXT_KEYS):
            kind = "early"
        elif any(k in low for k in GRO_PRE_HR_TEXT_KEYS):
            kind = "hr"
        else:
            kind = "early"
        out.setdefault(site, []).append((kind, a, b, text))
    tmp = cache.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["site", "kind", "start", "end", "text"])
        for s, items in out.items():
            for kind, a, b, text in items:
                w.writerow([s, kind, a.isoformat() if a else "", b.isoformat() if b else "", text])
    tmp.replace(cache)
    return out


def _pre_text_dates(line: str, year: int) -> list:
    out = []
    for m in GRO_PRE_MD_PAT.finditer(line):
        try:
            out.append(dt.date(year, int(m.group(1)), int(m.group(2))))
        except ValueError:
            pass
    return out


def gro_pre_verdict(items: list, es, ee, ms, gm, year: int):
    """GRO Pre 행 ↔ 법인별 Early 판정 (update_schedule.py 와 같은 규칙). 일치면 None, 다르면 (사유, GRO 시작, GRO End, Main 표시)."""
    early = [x for x in items if x[0] != "hr"]
    if not es and not early:
        return None
    ev_starts, ev_ranges, main_txt = set(), set(), set()
    for kind, _, _, text in items:
        for line in text.splitlines():
            low = line.lower()
            ds = _pre_text_dates(line, year)
            if not ds:
                continue
            if any(k in low for k in GRO_PRE_EARLY_TEXT_KEYS):
                ev_starts.add(ds[0])
                if len(ds) >= 2:
                    ev_ranges.add((ds[0], ds[1]))
            if kind != "hr" and any(k in low for k in GRO_PRE_MAIN_TEXT_KEYS):
                main_txt.add(ds[0])
    gs = min((x[1] for x in early if x[1]), default=None)
    ge = None if any(x[2] is None for x in early) else max((x[2] for x in early), default=None)
    if not es:
        return ("법인별 Early 없음", gs, ge, False)
    if not early:
        return None if (es, ee) in ev_ranges else ("GRO Pre-sale 없음", None, None, False)
    s_ok = gs == es or es in ev_starts
    if ge is None:
        e_ok = gm is None or ms is None or gm == ms
    else:
        e_ok = (ge == ee
                or (ee is not None and ge > ee and gm is not None and gm == ms)
                or ee in {b for _, b in ev_ranges}
                or (ee is not None and ge < ee and ms in main_txt))
    if s_ok and e_ok:
        return None
    reason = " / ".join(r for r, bad in (("시작 다름", not s_ok), ("End 다름", not e_ok)) if bad)
    return (reason, gs, ge, not e_ok)


def _as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    return v if isinstance(v, dt.date) else None


def apply_gro_notes(results: list, src: Path, next_src, gros: list) -> list:
    """src 이후(next_src 이전) 도착한 최신 GRO 와 Main 날짜가 다른 site 의 phase 2 team_note 에 한 줄을 쓴다."""
    lo, hi = _prefix(src), _prefix(next_src) if next_src else None
    cands = [g for g in gros if _prefix(g) > lo and (hi is None or _prefix(g) < hi)]
    if not cands:
        return results
    gro_file = cands[-1]
    gro = gro_site_dates(gro_file)
    fmt = lambda d: d.strftime(GRO_NOTE_DATE_FMT) if isinstance(d, dt.date) else (_fmt(d) or "-")
    out, n = [], 0
    for site, phase, start, end, src_subs, note, team_note in results:
        if phase == PHASE_MAIN and site.lower() in gro and str(src_subs).strip():
            g_start, g_end = gro[site.lower()]
            c_start, c_end = ("", "") if start == PART_NO else (start, end)
            if (g_start or g_end) and not ((g_start is None or g_start == c_start) and (g_end is None or g_end == c_end)):
                line = GRO_NOTE_FORMAT.format(stamp=_prefix(gro_file), site=site, g_start=fmt(g_start), g_end=fmt(g_end),
                                              c_start=fmt(c_start), c_end=fmt(c_end))
                team_note = f"{team_note}\n{line}" if team_note else line
                n += 1
        out.append((site, phase, start, end, src_subs, note, team_note))
    print(f"[GRO Note] {gro_file.name} 와 날짜가 다른 site {n}개")

    # ── GRO Status (Pre) ↔ 법인별 Early ──
    if not GRO_PRE_ENABLED:
        return out
    try:
        pre = gro_pre_rows(gro_file)
    except (OSError, ValueError, KeyError) as e:
        print(f"[GRO Pre Note] 생략 ({type(e).__name__}: {e})")
        return out
    year = 2000 + int(_prefix(gro_file)[:2])
    info = {}   # site → {phase: (start, end, src_subs)}
    for site, phase, start, end, src_subs, *_ in out:
        info.setdefault(site, {})[phase] = (start, end, src_subs)
    pre_lines = {}
    for site, ph in info.items():
        main = ph.get(PHASE_MAIN)
        if main is None or main[0] == PART_NO or not str(main[2]).strip():
            continue                                  # 미참여 / 고객 행 매칭 없음
        es, ee = (_as_date(x) for x in ph.get(PHASE_EARLY, (None, None, ""))[:2])
        ms = _as_date(main[0])
        gm = gro.get(site.lower(), (None, None))[0]
        v = gro_pre_verdict(pre.get(site.lower(), []), es, ee, ms, gm, year)
        if v is None:
            continue
        reason, gs, ge, show_main = v
        g = f"{fmt(gs)} ~ {fmt(ge)}" if gs or ge else "없음"
        c = f"{fmt(es)} ~ {fmt(ee)}" if es or ee else "없음"
        main_txt = GRO_PRE_MAIN_FORMAT.format(gm=fmt(gm), cm=fmt(ms)) if show_main else ""
        pre_lines[(site, PHASE_EARLY if PHASE_EARLY in ph else PHASE_MAIN)] = GRO_PRE_NOTE_FORMAT.format(
            stamp=_prefix(gro_file), site=site, reason=reason, g=g, c=c, main=main_txt)
    final = []
    for site, phase, start, end, src_subs, note, team_note in out:
        line = pre_lines.get((site, phase))
        if line:
            team_note = f"{team_note}\n{line}" if team_note else line
        final.append((site, phase, start, end, src_subs, note, team_note))
    print(f"[GRO Pre Note] {gro_file.name} (Pre) 와 Early 가 다른 site {len(pre_lines)}개")
    return final


# ── site 매핑 CSV ───────────────────────────────────────────
def manual_excluded_sites() -> set:
    return {ln.split("#", 1)[0].strip().casefold() for ln in EXCLUDE_SITES.splitlines()} - {""}


def _read_bytes_shared(path: Path) -> bytes:
    """다른 프로세스가 잡고 있어도 읽히도록 공유모드 READ|WRITE|DELETE 로 연다.
    (Python open()/shutil.copy 는 DELETE 공유를 주지 않아, Auto 를 잡은 숨은 Excel 등에 PermissionError 가 난다)"""
    try:
        h = win32file.CreateFile(
            str(path), win32file.GENERIC_READ,
            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
            None, win32file.OPEN_EXISTING, 0, None)
    except pywintypes.error as e:
        raise OSError(f"{path.name} 열기 실패 ({e.winerror}: {e.strerror})") from None
    try:
        chunks = []
        while True:
            _, data = win32file.ReadFile(h, 1 << 20)
            if not data:
                break
            chunks.append(bytes(data))
        return b"".join(chunks)
    except pywintypes.error as e:
        raise OSError(f"{path.name} 읽기 실패 ({e.winerror}: {e.strerror})") from None
    finally:
        h.Close()


def _load_wb_shared(path: Path):
    """공유모드로 읽은 바이트를 메모리에서 연다 (임시 사본 파일 없음)."""
    return openpyxl.load_workbook(io.BytesIO(_read_bytes_shared(path)), data_only=True, read_only=True)


def save_exclude_cache(sites: set) -> None:
    if EXCLUDE_CACHE_FILE is None:
        return
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    text = f"# {stamp} {EXCLUDE_SOURCE_XLSX.name} MASTER api수기제외=O\n" + "".join(f"{s}\n" for s in sorted(sites))
    old = EXCLUDE_CACHE_FILE.read_text(encoding="utf-8") if EXCLUDE_CACHE_FILE.exists() else ""
    if old.split("\n", 1)[-1] != text.split("\n", 1)[-1]:   # 목록이 바뀐 때만 씀 (20분마다 OneDrive 새 버전 방지)
        EXCLUDE_CACHE_FILE.write_text(text, encoding="utf-8")


def load_exclude_cache():
    """(목록, 저장시각 줄) — 캐시가 없으면 None."""
    if EXCLUDE_CACHE_FILE is None or not EXCLUDE_CACHE_FILE.exists():
        return None
    lines = EXCLUDE_CACHE_FILE.read_text(encoding="utf-8").splitlines()
    head = lines[0].lstrip("# ").strip() if lines and lines[0].startswith("#") else ""
    return {ln.strip().casefold() for ln in lines if ln.strip() and not ln.startswith("#")}, head


# ── 팀원 발신 메일 제외 ───────────────────────────────────────
def _load_sender_cache() -> dict:
    if SENDER_CACHE_FILE is None or not SENDER_CACHE_FILE.exists():
        return {}
    with SENDER_CACHE_FILE.open(encoding="utf-8-sig", newline="") as f:
        return {r["file"]: (r.get("sender_name") or "", r.get("sender_addr") or "")
                for r in csv.DictReader(f) if r.get("file")}


def _save_sender_cache(cache: dict) -> None:
    if SENDER_CACHE_FILE is None:
        return
    with SENDER_CACHE_FILE.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(SENDER_COLUMNS)
        for name in sorted(cache):
            w.writerow([name, *cache[name]])


def _mail_sender_addr(mail) -> str:
    """Exchange 내부 발신자는 SenderEmailAddress 가 '/O=EXCHANGE…' 형태라 SMTP 주소로 바꿔 읽는다."""
    try:
        user = mail.Sender.GetExchangeUser()
        if user is not None and user.PrimarySmtpAddress:
            return user.PrimarySmtpAddress
    except (pywintypes.com_error, AttributeError):
        pass
    return mail.SenderEmailAddress or ""


def _lookup_senders_outlook(files: list) -> dict:
    """{파일명: (발신자 이름, 주소)} — check_mail_attachment.py 저장명(수신일시 prefix + 원본 첨부명)을 거꾸로 맞춰 메일을 찾는다."""
    want = {}
    for f in files:
        m = MAIL_PREFIX_PAT.match(f.name)
        want[(_prefix(f), f.name[m.end():].lower())] = f.name
    oldest = min(p for p, _ in want)
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    store = next((s for s in ns.Stores if MAIL_STORE_NAME.lower() in s.DisplayName.lower()), None)
    if store is None:
        raise OSError(f"메일함을 찾을 수 없음: {MAIL_STORE_NAME}")
    items = store.GetDefaultFolder(OUTLOOK_INBOX).Items
    items.Sort("[ReceivedTime]", True)          # 최신순 — 가장 오래된 파일보다 이전이면 멈춘다
    found = {}
    for mail in items:
        try:
            stamp = mail.ReceivedTime.strftime("%y%m%d_%H%M")
            atts = [a.FileName.lower() for a in mail.Attachments]
        except (pywintypes.com_error, AttributeError):
            continue                            # 일정 요청 등 ReceivedTime/첨부가 없는 항목
        if stamp < oldest:
            break
        hits = [want[(stamp, a)] for a in atts if (stamp, a) in want]
        if hits:
            sender = (mail.SenderName or "", _mail_sender_addr(mail))
            found.update({h: sender for h in hits})
            if len(found) == len(want):
                break
    return found


def source_senders(sources: list) -> dict:
    """{고객 파일명: (발신자 이름, 주소)} — 캐시에 없는 파일만 Outlook 에서 찾아 캐시에 더한다. 못 찾은 파일은 빠진다."""
    cache = _load_sender_cache()
    missing = [f for f in sources if f.name not in cache]
    if missing:
        try:
            found = _lookup_senders_outlook(missing)
        except (pywintypes.com_error, OSError) as e:
            print(f"[경고] Outlook 에서 발신자를 찾을 수 없음 ({type(e).__name__}: {e})")
            found = {}
        else:
            # Outlook 조회는 끝났는데 메일이 없는 파일 = 폴더에 직접 넣은 수기 파일 → 수기 발신자로 기록해 처리한다
            if NO_MAIL_AS_MANUAL:
                for f in missing:
                    if f.name not in found:
                        found[f.name] = (MANUAL_SENDER_NAME, MANUAL_SENDER_ADDR)
                        print(f"[수기] 받은편지함에 메일 없음 - 폴더 파일로 처리: {f.name}")
        if found:
            cache.update(found)
            _save_sender_cache(cache)
    return cache


def filter_team_sources(sources: list) -> list:
    """팀원 도메인이 보낸 파일은 빼고, 발신자를 모르는 새 파일(CSV 없음)은 이번 회차 보류."""
    senders = source_senders(sources)
    doms = tuple("@" + d.lower().lstrip("@") for d in TEAM_SENDER_DOMAINS)
    kept = []
    for f in sources:
        out_name = f"{OUTPUT_PREFIX}{_prefix(f)}.csv"
        sender = senders.get(f.name)
        if sender and sender[1].lower().endswith(doms):
            print(f"[제외] 팀원 발신 {f.name} ({sender[0]} {sender[1]})")
            for d in (OUTPUT_DIR, OUTPUT_COPY_DIR):
                if d and (d / out_name).exists():
                    print(f"[경고] 팀원 발신 파일로 만든 CSV 가 남아 있음 — 수동 삭제 필요: {d / out_name}")
            continue
        if sender is None and not (OUTPUT_DIR / out_name).exists():
            print(f"[보류] 발신자 확인 전 — 다음 회차에 다시 확인: {f.name}")
            continue
        kept.append(f)
    return kept


def source_excluded_sites() -> set:
    """EXCLUDE_SOURCE_XLSX MASTER 'api수기제외' = O 인 site. 읽기 실패·헤더 불일치는 예외로 올린다 (호출부가 캐시 사용 또는 SKIP)."""
    if EXCLUDE_SOURCE_XLSX is None:
        return set()
    wb = _load_wb_shared(EXCLUDE_SOURCE_XLSX)
    try:
        ws = wb[EXCLUDE_SOURCE_SHEET]
        si, fi = _col_idx(EXCLUDE_SOURCE_SITE_COL) - 1, _col_idx(EXCLUDE_SOURCE_FLAG_COL) - 1
        rows = list(ws.iter_rows(min_row=EXCLUDE_SOURCE_HEADER_ROW, max_col=max(si, fi) + 1, values_only=True))
    finally:
        wb.close()
    hdr = _fmt(rows[0][fi]) if rows else ""
    if hdr != EXCLUDE_SOURCE_FLAG_HDR:
        raise ValueError(f"{EXCLUDE_SOURCE_SHEET}!{EXCLUDE_SOURCE_FLAG_COL}{EXCLUDE_SOURCE_HEADER_ROW} 헤더가 "
                         f"{EXCLUDE_SOURCE_FLAG_HDR!r} 아님 ({hdr!r}) — 칼럼 위치 확인")
    mark = EXCLUDE_SOURCE_FLAG_MARK.casefold()
    return {_fmt(r[si]).casefold() for r in rows[1:]
            if not _blank(r[si]) and _fmt(r[fi]).casefold() == mark}


def load_site_map() -> list:
    if not SITE_MAP_CSV.exists():
        raise FileNotFoundError(f"site 매핑 CSV 없음: {SITE_MAP_CSV} — `python {Path(__file__).name} {FLAG_BUILD_MAP}` 로 만든다")
    with SITE_MAP_CSV.open(encoding=OUTPUT_ENCODING, newline="") as f:
        rows = [r for r in csv.DictReader(f) if (r.get("site_code") or "").strip()]
    if not rows:
        raise ValueError(f"site 매핑 CSV 가 비어 있음: {SITE_MAP_CSV}")
    for r in rows:
        r["site_code"], r["subs"] = r["site_code"].strip(), (r.get("subs") or "").strip()
    return rows


def build_site_map() -> int:
    auto = BASE / AUTO_FILE_NAME
    wb = _load_wb_shared(auto)
    try:
        ws = wb[TAG_SHEET]
        idx = {k: _col_idx(v) - 1 for k, v in TAG_COLS.items()}
        rows = list(ws.iter_rows(min_row=TAG_HEADER_ROW + 1, max_col=max(idx.values()) + 1, values_only=True))
    finally:
        wb.close()
    out = [{"site_code": str(r[idx["site"]]).strip(), "subs": _fmt(r[idx["subs"]]),
            "region": _fmt(r[idx["region"]]), "country": _fmt(r[idx["country"]])}
           for r in rows if not _blank(r[idx["no"]]) and not _blank(r[idx["site"]])]
    with SITE_MAP_CSV.open("w", encoding=OUTPUT_ENCODING, newline="") as f:
        w = csv.DictWriter(f, fieldnames=MAP_COLUMNS, lineterminator="\n")
        w.writeheader()
        w.writerows(out)
    print(f"[매핑 생성] {SITE_MAP_CSV.name} ({len(out)} site) ← {auto.name} '{TAG_SHEET}'")
    return 0


# ── 출력 ─────────────────────────────────────────────────────
def csv_text(results) -> str:
    # QUOTE_MINIMAL: 콤마·줄바꿈·" 가 든 값만 양끝을 "…" 로 감싸고 안쪽 " 는 "" 로 바꾼다 → Note 문장에 콤마가 있어도 열이 밀리지 않는다.
    # (지난 회차 CSV 와 바이트가 같게 — 이런 문자가 없는 값은 따옴표 없이 그대로 둔다)
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n", quoting=csv.QUOTE_MINIMAL, doublequote=True)
    w.writerow(OUTPUT_COLUMNS)
    for site, phase, start, end, src_subs, note, team_note in results:
        w.writerow([CAMPAIGN_ID, site, phase, _fmt(start), _fmt(end), _fmt(src_subs), _fmt(note), _fmt(team_note)])
    return buf.getvalue()


def write_csv(out_file: Path, text: str, n: int, label: str = "") -> None:
    if out_file.exists():
        if out_file.read_text(encoding=OUTPUT_ENCODING).replace("\r\n", "\n") == text:
            print(f"[SKIP] 이미 저장됨: {label}{out_file.name}")
            return
        action = "갱신"
    else:
        action = "저장"
    out_file.parent.mkdir(parents=True, exist_ok=True)
    out_file.write_text(text, encoding=OUTPUT_ENCODING, newline="")
    print(f"[{action}] {label}{out_file.name} ({n} 행)")


def check_against_master(src: Path, results, excl: set) -> None:
    """Auto MASTER(답안지)와 site 별 대조 — Auto D1 스탬프가 이 고객 파일일 때만."""
    auto = BASE / AUTO_FILE_NAME
    try:
        wb = _load_wb_shared(auto)
    except (OSError, zipfile.BadZipFile) as e:   # 저장 도중이면 BadZipFile
        print(f"[대조 생략] Auto 파일을 읽을 수 없음 ({type(e).__name__}: {e})")
        return
    try:
        stamp = wb[STAMP_SHEET][STAMP_CELL].value
        idx = {k: _col_idx(v) - 1 for k, v in MASTER_COLS.items()}
        rows = list(wb[MASTER_SHEET].iter_rows(min_row=MASTER_HEADER_ROW + 1,
                                               max_col=max(idx.values()) + 1, values_only=True))
    finally:
        wb.close()
    if str(stamp or "").strip() != src.name:
        print(f"[대조 생략] Auto 에 반영된 소스가 다름 (Auto {stamp!r} / 이번 {src.name!r})")
        return
    master = {}   # (site, phase) → (start, end) — Early 는 H/I 중 하나라도 있을 때만
    for r in rows:
        site = str(r[idx["site"]]).strip() if not _blank(r[idx["site"]]) else ""
        if not site or site.casefold() in excl:
            continue
        es, ee = _fmt(r[idx["early_start"]]), _fmt(r[idx["early_end"]])
        if es or ee:
            master[(site, PHASE_EARLY)] = (es, ee)
        master[(site, PHASE_MAIN)] = (_fmt(r[idx["start"]]), _fmt(r[idx["end"]]))
    mine = {(s, p): (_fmt(a), _fmt(b)) for s, p, a, b, *_ in results if p != PHASE_HANDRAISER}   # MASTER 에 없는 phase
    diff = [(k, mine.get(k), master.get(k)) for k in sorted(set(mine) | set(master)) if mine.get(k) != master.get(k)]
    if diff:
        print(f"[대조 불일치] {len(diff)} 행 (site/phase: 계산값 / MASTER)")
        for (s, p), a, b in diff:
            print(f"   {s}/{p}: {a} / {b}")
    else:
        print(f"[대조 일치] MASTER 와 {len(mine)} 행 모두 같음 (Early {sum(p == PHASE_EARLY for _, p in mine)})")


def main() -> int:
    args = sys.argv[1:]
    if FLAG_BUILD_MAP in args:
        return build_site_map()

    try:
        site_map = load_site_map()
    except (OSError, ValueError) as e:
        print(f"[ERROR] {e}")
        return 1

    manual = manual_excluded_sites()
    try:
        auto_excl = source_excluded_sites()
        save_exclude_cache(auto_excl)
    except (OSError, KeyError, ValueError, zipfile.BadZipFile) as e:
        # 공유모드로도 못 읽음 / 저장 중(BadZipFile) / 시트·칼럼 구조 변경 → 마지막으로 읽은 목록으로 진행.
        # 캐시도 없으면 제외가 빠진 CSV 를 쓰지 않도록 이번 회차는 건너뛴다
        cached = load_exclude_cache()
        if cached is None:
            print(f"[SKIP] 제외 소스를 읽을 수 없고 캐시도 없음 ({type(e).__name__}: {e})")
            return 0
        auto_excl, head = cached
        print(f"[경고] 제외 소스를 읽을 수 없음 ({type(e).__name__}: {e}) → 이전 목록 사용 ({head})")
    excl = manual | auto_excl
    if excl:
        print(f"[제외] 수기 {sorted(manual)} / 소스 api수기제외 {sorted(auto_excl)}")
    unknown = excl - {m["site_code"].casefold() for m in site_map}
    if unknown:
        print(f"[경고] 제외 site 중 매핑표에 없는 site: {sorted(unknown)}")
    site_map = [m for m in site_map if m["site_code"].casefold() not in excl]

    sources = sorted((f for f in SOURCE_FOLDER.glob("*.xlsx")
                      if not f.name.startswith("~$") and MAIL_PREFIX_PAT.match(f.name)
                      and all(k in f.name for k in CUSTOMER_NAME_KEYS)), key=lambda f: f.name)
    if not sources:
        print(f"[SKIP] 수신일시 prefix 가 붙은 고객 법인 일정 파일이 아직 없습니다: {SOURCE_FOLDER}")
        return 0
    if TEAM_SENDER_DOMAINS:
        sources = filter_team_sources(sources)     # 팀원 발신 파일은 최신 회차 판정에서도 빠진다
        if not sources:
            print("[SKIP] 팀원 발신·보류를 빼면 처리할 고객 파일이 없습니다")
            return 0

    gros = gro_files() if GRO_NOTE_ENABLED else []
    for i, src in enumerate(sources):
        m = MAIL_PREFIX_PAT.match(src.name)
        out_file = OUTPUT_DIR / f"{OUTPUT_PREFIX}{m.group(1)}_{m.group(2)}.csv"
        is_latest = src == sources[-1]
        targets = [out_file] + ([OUTPUT_COPY_DIR / out_file.name] if OUTPUT_COPY_DIR else [])
        if out_file.exists() and FLAG_ALL not in args and not is_latest:
            # 지난 회차는 원본을 다시 계산하지 않는다 (그 시점 스냅샷 유지) — backup본이 없으면 원본을 그대로 복사만
            for t in targets[1:]:
                if not t.exists():
                    t.parent.mkdir(parents=True, exist_ok=True)
                    t.write_bytes(out_file.read_bytes())
                    print(f"[복사] (backup) {t.name}")
            continue
        try:
            results = compute_sites(read_customer_rows(src), site_map)
            print(f"[소스] {src.name}")
            if gros:
                try:
                    results = apply_gro_notes(results, src, sources[i + 1] if i + 1 < len(sources) else None, gros)
                except (OSError, ValueError, KeyError) as e:
                    # 부가 기능 — GRO 를 못 읽어도 CSV 는 고객 파일 기준으로 쓴다
                    print(f"[GRO Note] 생략 ({type(e).__name__}: {e})")
            text = csv_text(results)
            for t in targets:
                write_csv(t, text, len(results), "" if t == out_file else "(backup) ")
        except (OSError, ValueError, IndexError) as e:
            # OneDrive 동기화 중 잠김 / 포맷이 다른 파일 — 다음 회차가 재시도한다
            print(f"[SKIP] {src.name} 처리 실패 ({type(e).__name__}: {e})")
            continue
        if FLAG_CHECK in args and is_latest:
            check_against_master(src, results, excl)
    return 0


if __name__ == "__main__":
    sys.exit(main())
