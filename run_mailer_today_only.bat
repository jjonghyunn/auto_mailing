@echo off
:: ════════════════════════════════════════════════════════════════
::  run_mailer_today_only.bat
::  역할: 오늘 날짜(_YYMMDD)와 일치하는 xlsx 파일만 발송
::        오늘 날짜 파일이 없으면 발송하지 않고 종료
:: ════════════════════════════════════════════════════════════════

:: ── 설정 ──────────────────────────────────────────────────────
set PYTHON=python
set SCRIPT_DIR=%~dp0
set SCRIPT=%SCRIPT_DIR%run_mailer.py
set LOG=%SCRIPT_DIR%auto_mailer_bat.log

:: ── 실행 ──────────────────────────────────────────────────────
echo [%DATE% %TIME%] run_mailer_today_only.bat 시작 >> "%LOG%"
%PYTHON% "%SCRIPT%" --today-only
if %ERRORLEVEL% neq 0 (
    echo [%DATE% %TIME%] ERROR 또는 오늘 날짜 파일 없음 (exit code %ERRORLEVEL%) >> "%LOG%"
    exit /b %ERRORLEVEL%
)
echo [%DATE% %TIME%] 완료 >> "%LOG%"

:: ════════════════════════════════════════════════════════════════
::  작업 스케줄러 등록 방법 (관리자 권한 CMD에서 실행)
::
::  매일 오전 10시 발송:
::  schtasks /create /tn "AutoMailer_TodayOnly" /tr "\"%~f0\"" /sc daily /st 10:00 /f
::
::  등록 확인:
::  schtasks /query /tn "AutoMailer_TodayOnly"
::
::  삭제:
::  schtasks /delete /tn "AutoMailer_TodayOnly" /f
:: ════════════════════════════════════════════════════════════════
