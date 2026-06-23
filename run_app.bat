@echo off
title 栄養食事指導支援システム - ローカルサーバー
echo ===================================================
echo   栄養食事指導支援システム (Local Web Server)
echo ===================================================
echo.
echo ローカルサーバーを起動しています...
echo.
echo [接続URL] http://localhost:8000
echo.
echo ※このウィンドウを閉じるとサーバーが停止します。
echo ===================================================
echo.

:: ブラウザでアプリを開く
start "" "http://localhost:8000"

:: Pythonの簡易HTTPサーバーを起動
python -m http.server 8000
