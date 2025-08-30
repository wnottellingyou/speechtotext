@echo off
echo 正在安裝語音轉文字程式所需的套件...
echo.

echo 安裝語音識別相關套件...
pip install SpeechRecognition
pip install pyaudio
pip install pydub
pip install python-docx
pip install openai-whisper

echo.
echo 安裝完成！
echo 現在可以執行 python speech_to_text.py 來啟動程式
pause
