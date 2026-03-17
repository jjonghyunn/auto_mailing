@echo off
:: ════════════════════════════════════════════════════════════════
::  run_mailer_latest.bat
::  역할: 가장 최신 날짜 xlsx 파일을 자동 발송
::  스케줄러 등록 예시 → 아래 REGISTER 섹션 참고
:: ════════════════════════════════════════════════════════════════

:: ── 설정 ──────────────────────────────────────────────────────
set PYTHON=python
set SCRIPT_DIR=%~dp0
set SCRIPT=%SCRIPT_DIR%run_mailer.py
set LOG=%SCRIPT_DIR%auto_mailer_bat.log

:: ── 실행 ──────────────────────────────────────────────────────
echo [%DATE% %TIME%] run_mailer_latest.bat 시작 >> "%LOG%"
%PYTHON% "%SCRIPT%"
if %ERRORLEVEL% neq 0 (
    echo [%DATE% %TIME%] ERROR: 발송 실패 (exit code %ERRORLEVEL%) >> "%LOG%"
    exit /b %ERRORLEVEL%
)
echo [%DATE% %TIME%] 완료 >> "%LOG%"

:: ════════════════════════════════════════════════════════════════
::  작업 스케줄러 등록 방법 (관리자 권한 CMD에서 실행)
::
::  매일 오전 9시 발송:
::  schtasks /create /tn "AutoMailer_Latest" /tr "\"%~f0\"" /sc daily /st 09:00 /f
::
::  등록 확인:
::  schtasks /query /tn "AutoMailer_Latest"
::
::  삭제:
::  schtasks /delete /tn "AutoMailer_Latest" /f
:: ════════════════════════════════════════════════════════════════
