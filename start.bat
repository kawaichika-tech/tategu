@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
echo 建具図面チェッカー 起動
echo ブラウザで http://localhost:5000 を開いてください
python -X utf8 app.py
pause
