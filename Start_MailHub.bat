@echo off
REM ==========================================
REM RogoAI Mail Hub 起動スクリプト
REM ==========================================

echo RogoAI Mail Hub を起動しています...
echo.

REM Pythonの存在確認
python --version >nul 2>&1
if errorlevel 1 (
    echo [エラー] Pythonがインストールされていません
    echo Python 3.8以上をインストールしてください
    pause
    exit /b 1
)

REM main.pyの存在確認
if not exist "%~dp0main.py" (
    echo [エラー] main.pyが見つかりません
    pause
    exit /b 1
)

REM 起動
python "%~dp0main.py"

REM エラーがあった場合は一時停止
if errorlevel 1 (
    echo.
    echo [エラー] Mail Hubの起動に失敗しました
    pause
)