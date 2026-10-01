# 26_Schedule_separate(Auto).xlsx  
<sub>2026-07-10  Jonghyun Park w/ Claude</sub>  

Campaign schedule management template file.
Sanitized version for reference — actual data replaced with dummy values.

---

## Sheet Structure

### 1. MASTER
Campaign execution summary. Core reference for all other sheets.

| Col | Header | Description |
|---|---|---|
| B | Tier | Site tier (1 = Tier 1, 2 = Tier 2) |
| C | Region | Regional HQ (e.g. MENA, EU, LATAM) |
| D | Subs | Subsidiary name |
| E | COUNTRY | Country name |
| F | RS | Reporting segment |
| G | Site Code | Lowercase site code (e.g. `us`, `de`) |
| H | Start Date | TY campaign start date |
| I | End Date | TY campaign end date |
| J | api대상 여부 | Whether site is included in api extract |
| K | Start Date | LY start date |
| L | End Date | LY end date |
| M | Note | LY notes |
| N | URL | TY campaign URL (formula from RAW_URL,PageName) |
| O | Page Name | TY pagename (formula from RAW_URL,PageName) |

- Data starts from **row 5**
- Right-side block (col W~) is a secondary reference list from 탑캠사이트_code

---

### 2. Appendix_URL
Delivery-ready URL appendix per site. Mirrors official campaign appendix format.

| Col | Header | Source |
|---|---|---|
| B | Tier | Formula from MASTER |
| C | Region | Formula from MASTER |
| D | Subs | Formula from MASTER |
| E | Country | Formula from MASTER |
| F | Site Code | Formula from MASTER (uppercase) |
| G | This Year | XLOOKUP from RAW_URL,PageName (TY) |
| H | Last Year | XLOOKUP from RAW_URL,PageName (LY) |

- Header row: **row 9**
- Data starts from **row 10**
- C7: Last Updated (manual or `=TODAY()`)

**URL formula (G col, This Year):**
```excel
=IFERROR(XLOOKUP(1,
  (LOWER('RAW_URL,PageName'!$C$2:$C$500)=LOWER(F10))
  *('RAW_URL,PageName'!$B$2:$B$500="TY"),
  'RAW_URL,PageName'!$D$2:$D$500,""),"")
```

---

### 3. RAW_tier
Auto-generated tier lookup table per site_code.

| Col | Header |
|---|---|
| A | # (auto index) |
| B | site_code |
| C | Tier (formula from MASTER) |

---

### 4. RAW_URL,PageName
Core input sheet. One row = one URL/PageName entry.

| Col | Header | Notes |
|---|---|---|
| A | # (auto index) | |
| B | This/Last Year | `TY` or `LY` |
| C | site_code | Lowercase |
| D | url | Must start with `https://` |
| E | pagename | Analytics page name |
| F | note | Optional memo |
| G | TY ref | Master site ref for TY (56 sites) |
| H | LY ref | Master site ref for LY (73 sites) |
| I | TY | TY site_code cross-ref |
| J | LY | LY site_code cross-ref |

- Data starts from **row 2**, up to row 500 (formula range)

---

### 5. api대상csv
API extract date range by site. Date cutoff controls all date calculations.

| Section | Cols | Description |
|---|---|---|
| This Year | B–E | Site / Start / End / Days |
| Prior | F–I | Previous period |
| Last Year | J–M | YoY comparison |

- **Data Cutoff** cell: `D2` — change this to update all date ranges automatically.
- Data starts from **row 5**

---

### 6. 고객법인일정파일
Customer/client schedule reference sheet.
Auto-updated by `update_schedule.py` from the latest client-shared Excel file.

| Range | Content |
|---|---|
| D1 | Source filename (auto-written by script) |
| B2:K999 | Schedule data (cleared and rewritten on each update) |

---

### 7. 탑캠사이트_code
Master site code reference list.

| Col | Header |
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

- Header row: **row 3**
- Data starts from **row 4**

---

## Automation

| Script | Role |
|---|---|
| `update_schedule.py` | Reads latest client schedule xlsx → updates 고객법인일정파일 sheet |
| `check_mail_attachment.py` | Checks Outlook inbox → saves new attachments to local folder |

**Scheduled task:** both scripts run every 20 minutes via Windows Task Scheduler (`/it` — only when logged on).

See `README.md` for script details.
