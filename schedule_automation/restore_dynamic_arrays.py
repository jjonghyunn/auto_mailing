"""
restore_dynamic_arrays.py
2026-11-02  Jonghyun Park w/ Claude  — 신규. openpyxl 저장으로 고정 범위 CSE 배열로 바뀐 동적 배열 수식을 다시 동적 배열(Formula2)로 넣는다

배경: update_schedule.py 가 Auto 를 openpyxl 로 저장하던 시절, 동적 배열 메타데이터(셀 cm="1" + xl/metadata.xml)가
      버려져 FILTER/HSTACK/LET 수식이 CSE 배열로 바뀌고 #VALUE! 가 났다. 2026-11-02 부터 update_schedule.py 는 COM 으로 쓰므로
      더 깨지지 않지만, 이미 깨진 파일은 이 스크립트로 1회 복구한다.

대상 = DYNAMIC_ANCHORS (정상 파일에서 cm="1" 이던 앵커). --reference <정상 xlsx> 를 주면 그 파일의 cm="1" 앵커를 대신 쓴다.
※ 나머지 여러 칸 배열(MASTER AB5:AK85, RAW_URL H2 등)은 정상 파일에서도 CSE 라 건드리지 않는다.

사용:
  python restore_dynamic_arrays.py                       # dry-run: 대상 앵커 상태만 출력
  python restore_dynamic_arrays.py --apply               # 복구 후 저장
  python restore_dynamic_arrays.py --target <xlsx>       # 다른 파일 대상 (테스트용)
  python restore_dynamic_arrays.py --reference <xlsx>    # 정상 파일에서 대상 앵커를 뽑음
  python restore_dynamic_arrays.py --anchors "api용!B5,Appendix_URL!B10"   # 앵커 직접 지정 (다른 캠페인 Auto 용)
"""

import re
import sys
import time
import zipfile
from pathlib import Path
import pywintypes
import win32com.client

# ════ 사용자가 바꿔야 하는 부분 ════
AUTO_FILE = Path(
    r"C:\Users\user_name\OneDrive - company_name"
    r"\Project_team_name - 1 company_name - 02 part_name"
    r"\part_name\2026\# CAMPAIGN_PROJECTS\02. CAMPAIGN NAME\02. SCHEDULE"
    r"\26_Schedule_separate(Auto).xlsx"
)
# 정상 파일에서 cm="1" 이던 앵커 — 워크북마다 다르므로 --reference <정상 xlsx> 또는 --anchors 로 바꿔 쓸 것
DYNAMIC_ANCHORS = [
    ("api용", "B5"),
    ("Appendix_Date", "B12"),
    ("Appendix_Date", "X12"),
    ("Appendix_Date", "Y12"),
    ("Appendix_URL", "B10"),
]

# ════ 내부 사용 ════
APPLY_FLAG, TARGET_FLAG, REF_FLAG, ANCHORS_FLAG = "--apply", "--target", "--reference", "--anchors"
RPC_E_CALL_REJECTED = -2147418111
COM_RETRIES, COM_RETRY_WAIT_SEC = 10, 3
XL_ERRORS = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
             -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826238: "#SPILL!",
             -2146826233: "#CALC!"}


def _arg(flag):
    a = sys.argv[1:]
    return a[a.index(flag) + 1] if flag in a and a.index(flag) + 1 < len(a) else None


def anchors_from_reference(ref: Path) -> list:
    z = zipfile.ZipFile(ref)
    wb = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    out = []
    for name, rid in re.findall(r'<sheet name="([^"]+)" sheetId="\d+" r:id="(rId\d+)"', wb):
        tgt = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels).group(1)
        x = z.read("xl/" + tgt.replace("/xl/", "").lstrip("/")).decode()
        out += [(name.replace("&amp;", "&"), c) for c in re.findall(r'<c r="([A-Z]+\d+)"[^>]*cm="1"', x)]
    return out


def retry(fn):
    """막 열린 Excel 이 바쁘면 '피호출자가 호출을 거부했습니다'(RPC_E_CALL_REJECTED) — 잠시 뒤 다시 시도."""
    for attempt in range(1, COM_RETRIES + 1):
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult != RPC_E_CALL_REJECTED or attempt == COM_RETRIES:
                raise
            time.sleep(COM_RETRY_WAIT_SEC)


def error_counts(wb) -> dict:
    """시트별 오류 셀 수 (UsedRange Value2 의 오류코드)."""
    out = {}
    for i in range(1, retry(lambda: wb.Worksheets.Count) + 1):
        ws = retry(lambda: wb.Worksheets(i))
        v = retry(lambda: ws.UsedRange.Value2)
        rows = v if isinstance(v, tuple) else ((v,),)
        n = sum(1 for r in rows for c in (r if isinstance(r, tuple) else (r,)) if isinstance(c, int) and c in XL_ERRORS)
        if n:
            out[ws.Name] = n
    return out


def main() -> int:
    target = Path(_arg(TARGET_FLAG) or AUTO_FILE)
    ref = _arg(REF_FLAG)
    given = _arg(ANCHORS_FLAG)   # "시트!A1,시트!B2" — 시트명에 콤마가 있으면 안 됨
    if given:
        anchors = [tuple(a.rsplit("!", 1)) for a in given.split(",") if a.strip()]
    else:
        anchors = anchors_from_reference(Path(ref)) if ref else DYNAMIC_ANCHORS
    apply = APPLY_FLAG in sys.argv[1:]
    print(f"[대상] {target}\n[앵커] {anchors}{' (reference)' if ref else ''}\n[모드] {'APPLY' if apply else 'dry-run'}")

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(target.resolve()))
        if wb.ReadOnly:
            print("[SKIP] 파일이 다른 곳에서 열려 있어 읽기 전용으로 열렸습니다")
            wb.Close(False)
            return 1
        retry(xl.CalculateFull)
        before = error_counts(wb)
        print(f"[전] 오류 셀 {before}")
        changed = 0
        for sheet, addr in anchors:
            rng = wb.Worksheets(sheet).Range(addr)
            if not rng.HasFormula:
                print(f"  {sheet}!{addr}: 수식 없음 - 건너뜀")
                continue
            dyn = rng.HasSpill
            area = rng.CurrentArray.Address if rng.HasArray else ""
            print(f"  {sheet}!{addr}: {'동적배열(정상)' if dyn and not rng.HasArray else 'CSE ' + area} 값={rng.Text!r}")
            if not rng.HasArray or not apply:
                continue
            f = rng.Formula2
            rng.CurrentArray.ClearContents()
            rng.Formula2 = f
            changed += 1
        if apply and changed:
            retry(xl.CalculateFull)
            after = error_counts(wb)
            print(f"[후] 오류 셀 {after}")
            for sheet, addr in anchors:
                rng = wb.Worksheets(sheet).Range(addr)
                spill = rng.SpillingToRange.Address if rng.HasSpill else "-"
                print(f"  {sheet}!{addr}: HasArray={rng.HasArray} 펼침={spill} 값={rng.Text!r}")
            wb.Save()
            print(f"[저장] {changed}개 앵커 복구")
        wb.Close(False)
    except pywintypes.com_error as e:
        print(f"[ERROR] COM: {e}")
        return 1
    finally:
        try:
            xl.Quit()
        except (TypeError, AttributeError):
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
