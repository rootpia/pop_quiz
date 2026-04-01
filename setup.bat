@echo off
chcp 65001 >nul
echo =======================================
echo Pop Quiz (英語問題ツール) セットアップ
echo =======================================

:: 管理者権限チェック
openfiles >nul 2>&1
if '%errorlevel%' NEQ '0' (
    echo [!] タスクスケジューラに登録するため、管理者権限が必要です。
    echo 右クリックから「管理者として実行」してください。
    pause
    exit /b
)

set SCRIPT_DIR=%~dp0
set HTA_PATH=%SCRIPT_DIR%pop_quiz.hta

echo 現在のディレクトリからツールをタスクに登録します:
echo %HTA_PATH%
echo.
echo [1] 登録する (ログイン時 および ロック解除時 に起動)
echo [2] 登録を解除する
echo [3] キャンセル
set /p choice="選択してください (1-3): "

if "%choice%"=="1" goto INSTALL
if "%choice%"=="2" goto UNINSTALL
goto END

:INSTALL
echo 登録中...
schtasks /create /tn "PopQuizUnlock" /tr "mshta.exe \"%HTA_PATH%\"" /sc onunlock /rl highest /f >nul 2>&1
schtasks /create /tn "PopQuizLogon" /tr "mshta.exe \"%HTA_PATH%\"" /sc onlogon /rl highest /f >nul 2>&1
if '%errorlevel%' EQU '0' (
    echo 登録が完了しました！次回のロック解除時やログイン時に起動します。
) else (
    echo [!] 登録に失敗しました。
)
pause
exit /b

:UNINSTALL
echo 解除中...
schtasks /delete /tn "PopQuizUnlock" /f >nul 2>&1
schtasks /delete /tn "PopQuizLogon" /f >nul 2>&1
echo 登録を解除しました。
pause
exit /b

:END
exit /b
