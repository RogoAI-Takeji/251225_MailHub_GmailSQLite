@echo off
REM ==========================================
REM RogoAI Mail Hub カスタム起動スクリプト（上級者向け）
REM ==========================================
REM 
REM このファイルは、設定ファイルとデータベースの保存場所を
REM カスタマイズしたい上級者向けのサンプルです。
REM 
REM 通常のユーザーは Start_MailHub.bat を使用してください。
REM ==========================================

echo RogoAI Mail Hub（カスタム設定）を起動しています...
echo.

REM ==========================================
REM カスタム保存場所の設定（必要に応じて変更してください）
REM ==========================================
REM 
REM 例1: Dドライブの専用フォルダに保存
REM set MAILHUB_CONFIG_DIR=D:\MySettings\MailHub
REM set MAILHUB_STORAGE_DIR=D:\MyData\MailHub
REM 
REM 例2: ドキュメントフォルダ内に保存
REM set MAILHUB_CONFIG_DIR=%USERPROFILE%\Documents\MailHub\config
REM set MAILHUB_STORAGE_DIR=%USERPROFILE%\Documents\MailHub\storage
REM 
REM 以下の行のコメント（REM）を外して、パスを設定してください
REM ==========================================

REM set MAILHUB_CONFIG_DIR=
REM set MAILHUB_STORAGE_DIR=

REM ==========================================
REM 注意事項
REM ==========================================
REM ⚠️ 以下の場所は避けてください：
REM   - ネットワークドライブ（\\server\share など）
REM   - 日本語を含むパス
REM   - 外付けHDD/USBメモリ
REM   - システムフォルダ（C:\Windows など）
REM ==========================================

REM Pythonの存在確認
python --version >nul 2>&1
if errorlevel 1 (
    echo [エラー] Pythonがインストールされていません
    pause
    exit /b 1
)

REM 環境変数が設定されているか確認
if defined MAILHUB_CONFIG_DIR (
    echo [カスタム設定] CONFIG: %MAILHUB_CONFIG_DIR%
)
if defined MAILHUB_STORAGE_DIR (
    echo [カスタム設定] STORAGE: %MAILHUB_STORAGE_DIR%
)
echo.

REM 起動
python "%~dp0main.py"

REM エラーがあった場合は一時停止
if errorlevel 1 (
    echo.
    echo [エラー] Mail Hubの起動に失敗しました
    pause
)