"""
check_mail_attachment.py
2026-10-12  Jonghyun Park w/ Claude  — 저장명 앞에 수신일시(yymmdd_hhmm_) prefix 를 항상 붙인다 (기존: 중복일 때만 뒤에 _yymmdd_HHMM)
                                       + RECEIVED_FROM: 지정 날짜(당일 포함) 이후 받은 메일만 처리 — 캠페인 폴더를 복사해 쓸 때 지난 캠페인 메일 재수집 방지
2026-07-22  Jonghyun Park w/ Claude  — 중복 suffix _yymmdd → _yymmdd_HHMM
(이전 버전 이력은 git history / GitHub Releases 참조 — 헤더에는 최근 2개 항목만 남긴다)

[mailbox] 받은편함에서 CAMPAIGN NAME 법인별 일정 첨부파일을 감지해 저장.
- 이미 처리한 메일은 EntryID로 기록해 재처리 방지
- 저장명 앞에 수신일시(yymmdd_hhmm_) 를 붙여 저장 (같은 이름이 다시 와도 수신 시각으로 구분)
- SaveAsFile은 MAX_PATH 제한으로 임시폴더에 저장 후 shutil.move로 이동
"""

import win32com.client
import shutil
import tempfile
from pathlib import Path
from datetime import datetime, date

SAVE_FOLDER  = Path(
    r"C:\Users\user_name\OneDrive - company_name"
    r"\Project_team_name - 1 company_name - 02 part_name"
    r"\part_name\2026\# CAMPAIGN_PROJECTS\03. CAMPAIGN NAME"
    r"\02. SCHEDULE\1.고객 법인 일정 파일"
)
STORE_NAME   = "part_name mailbox"
SUBJECT_KEYS = ["campaign name", "법인별"]  # 모두 포함된 메일만
MARKER_FILE  = Path(r"C:\Users\user_name\Documents\campaign_mail_processed_ids.txt")
# 받은 날짜가 이 날짜 이상인 메일만 처리 — **이 날짜 당일 포함** (예: 9/29 → 9/29 받은 메일부터). None = 전체
# 지난 캠페인 폴더를 복사해 새 캠페인을 시작할 때, 같은 키워드의 옛 메일을 다시 집어오지 않게 캠페인 시작일을 넣는다.
RECEIVED_FROM = date(2026, 1, 1)

# ── 긴 경로 존재 확인 (\\?\ 접두사 사용) ──────────────────────
def long_path_exists(p: Path) -> bool:
    lp = Path("\\\\?\\" + str(p.resolve()))
    return lp.exists()

# ── 처리된 EntryID 로드 ───────────────────────────────────────
if MARKER_FILE.exists():
    processed_ids = set(MARKER_FILE.read_text(encoding="utf-8").splitlines())
else:
    processed_ids = set()

# ── Outlook 연결 ──────────────────────────────────────────────
outlook = win32com.client.Dispatch("Outlook.Application")
ns      = outlook.GetNamespace("MAPI")

target_store = None
for store in ns.Stores:
    if STORE_NAME.lower() in store.DisplayName.lower():
        target_store = store
        break

if target_store is None:
    raise RuntimeError(f"메일함을 찾을 수 없습니다: {STORE_NAME}")

inbox = target_store.GetDefaultFolder(6)  # 6 = 받은편함
print(f"[메일함] {target_store.DisplayName} / 받은편함 ({inbox.Items.Count}개)")

# ── 메일 순회 ────────────────────────────────────────────────
saved = 0
skipped = 0
new_ids = []

for mail in inbox.Items:
    try:
        subject  = mail.Subject.lower()
        entry_id = mail.EntryID
    except Exception:
        continue

    if not all(k in subject for k in SUBJECT_KEYS):
        continue

    if RECEIVED_FROM:
        rt = mail.ReceivedTime
        if date(rt.year, rt.month, rt.day) < RECEIVED_FROM:
            continue

    if entry_id in processed_ids:
        skipped += 1
        continue

    mail_saved = False
    received = datetime.strftime(mail.ReceivedTime, "%y%m%d_%H%M")

    for att in mail.Attachments:
        name = att.FileName
        if not name.lower().endswith(".xlsx"):
            continue

        # 저장명 = 수신일시(yymmdd_hhmm_) + 원본 첨부명 — 같은 이름이 다시 와도 수신 시각으로 구분된다
        suffix = Path(name).suffix
        dest   = SAVE_FOLDER / f"{received}_{Path(name).stem}{suffix}"

        # 같은 수신일시 + 같은 이름이 이미 있음 = 같은 메일 재처리 → 스킵
        if long_path_exists(dest):
            print(f"[스킵] 이미 처리됨: {dest.name}")
            skipped += 1
            mail_saved = True   # 이미 저장된 것으로 간주 → EntryID 기록
            continue

        # 임시폴더 저장 후 이동
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp_path = Path(tmp.name)
        att.SaveAsFile(str(tmp_path))
        shutil.move(str(tmp_path), str(dest))
        print(f"[저장] {name} → {dest.name}")

        saved += 1
        mail_saved = True

    if mail_saved:
        new_ids.append(entry_id)

# ── 처리된 EntryID 저장 ───────────────────────────────────────
if new_ids:
    with MARKER_FILE.open("a", encoding="utf-8") as f:
        f.write("\n".join(new_ids) + "\n")

print(f"\n완료 - 저장 {saved}개 / 스킵 {skipped}개")
