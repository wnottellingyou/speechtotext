@echo off
echo 正在安裝語音轉文字程式所需套件...
echo.

REM 設定Python路徑
set PYTHON_PATH="C:/Users/蔡昌諺/AppData/Local/Programs/Python/Python310/python.exe"

REM 升級pip
%PYTHON_PATH% -m pip install --upgrade pip

REM 安裝基本套件
%PYTHON_PATH% -m pip install SpeechRecognition
%PYTHON_PATH% -m pip install pyaudio
%PYTHON_PATH% -m pip install python-docx
%PYTHON_PATH% -m pip install pydub

REM 嘗試安裝Whisper（可選）
echo 正在安裝Whisper（可能需要較長時間）...
%PYTHON_PATH% -m pip install openai-whisper

echo.
echo 安裝完成！
echo 您現在可以執行以下程式：
echo 1. speech_to_text.py （完整版，支援Whisper）
echo 2. simple_speech_to_text.py （簡化版，僅使用Google API）
echo.
pause
