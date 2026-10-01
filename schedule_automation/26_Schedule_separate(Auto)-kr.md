# 26_Schedule_separate(Auto).xlsx  
<sub>2026-07-10  Jonghyun Park w/ Claude</sub>  

캠페인 일정 관리 템플릿 파일.
실제 데이터는 더미값으로 대체한 참고용 정제 버전.

---

## 시트 구성

### 1. MASTER
캠페인 실행 요약 시트. 모든 시트의 핵심 기준 데이터.

| 열 | 헤더 | 설명 |
|---|---|---|
| B | Tier | 사이트 티어 (1 = Tier 1, 2 = Tier 2) |
| C | Region | 지역 본사 (예: MENA, EU, LATAM) |
| D | Subs | 법인명 |
| E | COUNTRY | 국가명 |
| F | RS | 리포팅 세그먼트 |
| G | Site Code | 소문자 사이트 코드 (예: `us`, `de`) |
| H | Start Date | TY 캠페인 시작일 |
| I | End Date | TY 캠페인 종료일 |
| J | api대상 여부 | API 추출 대상 사이트 여부 |
| K | Start Date | LY 시작일 |
| L | End Date | LY 종료일 |
| M | Note | LY 비고 |
| N | URL | TY 캠페인 URL (RAW_URL,PageName 시트 수식) |
| O | Page Name | TY 페이지네임 (RAW_URL,PageName 시트 수식) |

- 데이터 시작 행: **5행**
- 우측 블록 (W열~): 탑캠사이트_code 시트의 보조 참조 목록

---

### 2. Appendix_URL
사이트별 배포용 URL 별첨 시트. 공식 캠페인 별첨 양식을 그대로 미러링.

| 열 | 헤더 | 출처 |
|---|---|---|
| B | Tier | MASTER 시트 수식 |
| C | Region | MASTER 시트 수식 |
| D | Subs | MASTER 시트 수식 |
| E | Country | MASTER 시트 수식 |
| F | Site Code | MASTER 시트 수식 (대문자) |
| G | This Year | RAW_URL,PageName 시트 XLOOKUP (TY) |
| H | Last Year | RAW_URL,PageName 시트 XLOOKUP (LY) |

- 헤더 행: **9행**
- 데이터 시작 행: **10행**
- C7: 최종 업데이트일 (수동 입력 또는 `=TODAY()`)

**URL 수식 (G열, This Year):**
```excel
=IFERROR(XLOOKUP(1,
  (LOWER('RAW_URL,PageName'!$C$2:$C$500)=LOWER(F10))
  *('RAW_URL,PageName'!$B$2:$B$500="TY"),
  'RAW_URL,PageName'!$D$2:$D$500,""),"")
```

---

### 3. RAW_tier
site_code별 티어 자동 생성 조회 테이블.

| 열 | 헤더 |
|---|---|
| A | # (자동 인덱스) |
| B | site_code |
| C | Tier (MASTER 시트 수식) |

---

### 4. RAW_URL,PageName
핵심 입력 시트. 1행 = URL/PageName 항목 1건.

| 열 | 헤더 | 비고 |
|---|---|---|
| A | # (자동 인덱스) | |
| B | This/Last Year | `TY` 또는 `LY` |
| C | site_code | 소문자 |
| D | url | `https://`로 시작해야 함 |
| E | pagename | 애널리틱스 페이지네임 |
| F | note | 선택 메모 |
| G | TY ref | TY 기준 MASTER 사이트 참조 (56개) |
| H | LY ref | LY 기준 MASTER 사이트 참조 (73개) |
| I | TY | TY site_code 교차 참조 |
| J | LY | LY site_code 교차 참조 |

- 데이터 시작 행: **2행**, 최대 500행 (수식 범위)

---

### 5. api대상csv
사이트별 API 추출 날짜 범위. Data Cutoff 셀이 모든 날짜 계산의 기준.

| 구분 | 열 | 설명 |
|---|---|---|
| This Year | B–E | 사이트 / 시작일 / 종료일 / 일수 |
| Prior | F–I | 이전 기간 |
| Last Year | J–M | YoY 비교 |

- **Data Cutoff** 셀: `D2` — 이 셀 값을 변경하면 모든 날짜 범위 자동 갱신.
- 데이터 시작 행: **5행**

---

### 6. 고객법인일정파일
고객/법인 일정 참조 시트.
`update_schedule.py`가 클라이언트 공유 최신 엑셀 파일에서 자동 업데이트.

| 범위 | 내용 |
|---|---|
| D1 | 소스 파일명 (스크립트 자동 기입) |
| B2:K999 | 일정 데이터 (업데이트 시 초기화 후 재작성) |

---

### 7. 탑캠사이트_code
마스터 사이트 코드 참조 목록.

| 열 | 헤더 |
|---|---|
| B | # |
| C | Region |
| D | Subsidiary |
| E | Country |
| F | Site Code |
| G | Base URL |
| H | Shop Platform |
| I | App 설치 |
| J | RS |
| K | App 여부 |
| L | 캠페인_CampaignPart |

- 헤더 행: **3행**
- 데이터 시작 행: **4행**

---

## 자동화

| 스크립트 | 역할 |
|---|---|
| `update_schedule.py` | 최신 클라이언트 일정 xlsx 읽기 → 고객법인일정파일 시트 업데이트 |
| `check_mail_attachment.py` | Outlook 수신함 확인 → 신규 첨부파일 로컬 폴더에 저장 |

**예약 작업:** 두 스크립트 모두 Windows 작업 스케줄러를 통해 매 20분마다 실행 (`/it` — 로그온 시에만 실행).

스크립트 상세 내용은 `README.md` 참고.
