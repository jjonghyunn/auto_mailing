# Windows LongPath 활성화 방법  
<sub>2026-07-01  Jonghyun Park w/ Claude</sub>  

## 왜 필요한가

`check_mail_attachment.py`에서 Outlook 첨부파일을 OneDrive 하위 깊은 경로에 저장할 때,
Windows 기본 경로 길이 제한(MAX_PATH = 260자)을 초과하면 `att.SaveAsFile()`이
`pywintypes.com_error: -2147024893 (ERROR_PATH_NOT_FOUND)` 오류로 실패한다.

저장 대상 경로가 ~260자에 육박하는 경우 아래 설정을 적용해야 한다.

---

## 방법 1: 관리자 PowerShell에서 직접 실행 (권장)

1. 시작 메뉴 → "powershell" 검색 → 우클릭 → "관리자 권한으로 실행"
2. 아래 명령어 붙여넣기:

```powershell
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1
```

3. 적용 확인:

```powershell
Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled"
# LongPathsEnabled : 1  이면 정상
```

---

## 방법 2: 레지스트리 직접 수정

1. `Win+R` → `regedit` 실행
2. 경로 이동: `HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem`
3. `LongPathsEnabled` 값을 `1`로 변경
