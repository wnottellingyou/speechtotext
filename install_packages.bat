@echo off
echo ====================================
echo 語音轉文字程式 - 套件安裝腳本
echo ====================================
echo.

rem 使用Python 3.10環境
set PYTHON_PATH="C:/Users/蔡昌諺/AppData/Local/Programs/Python/Python310/python.exe"

echo 使用Python環境: %PYTHON_PATH%
echo.

echo 正在升級 pip...
%PYTHON_PATH% -m pip install --upgrade pip
echo.

echo 正在安裝必要套件...
echo.

echo [1/7] 安裝 speech_recognition...
%PYTHON_PATH% -m pip install SpeechRecognition
echo.

echo [2/7] 安裝 pyaudio...
%PYTHON_PATH% -m pip install pyaudio
echo.

echo [3/7] 安裝 python-docx...
%PYTHON_PATH% -m pip install python-docx
echo.

echo [4/7] 安裝 pydub...
%PYTHON_PATH% -m pip install pydub
echo.

echo [5/7] 安裝 openai-whisper...
%PYTHON_PATH% -m pip install openai-whisper
echo.

echo [6/7] 確認 openai 已安裝...
%PYTHON_PATH% -m pip install openai
echo.

echo [7/7] 安裝 requests...
%PYTHON_PATH% -m pip install requests
echo.

echo ====================================
echo 安裝完成！
echo ====================================
echo.
echo 現在可以運行語音轉文字程式了：
echo 請使用 start_app.bat 啟動程式
echo.
pause
