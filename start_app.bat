@echo off
echo ====================================
echo 啟動語音轉文字程式 (含AI摘要功能)
echo ====================================
echo.

rem 使用Python 3.10環境 (OpenAI已安裝在此環境)
set PYTHON_PATH="C:/Users/蔡昌諺/AppData/Local/Programs/Python/Python310/python.exe"
set SCRIPT_PATH="d:/PYfile/speechtotext/speech_to_text.py"

echo 正在啟動程式...
echo 程式路徑: %SCRIPT_PATH%
echo Python路徑: %PYTHON_PATH% (使用Python 3.10環境)
echo.

cd /d "d:\PYfile\speechtotext"
%PYTHON_PATH% %SCRIPT_PATH%

echo.
echo 程式已結束。
pause
