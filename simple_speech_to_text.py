#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
簡化版語音轉文字程式
使用基本的speech_recognition庫
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import speech_recognition as sr
import threading
import os
from datetime import datetime

class SimpleSpeechToTextApp:
    def __init__(self, root):
        self.root = root
        self.root.title("簡化版語音轉文字程式")
        self.root.geometry("700x500")
        
        # 初始化變數
        self.recognizer = sr.Recognizer()
        self.audio_data = None
        
        # 建立UI
        self.create_widgets()
        
        # 嘗試初始化麥克風
        try:
            self.microphone = sr.Microphone()
            with self.microphone as source:
                self.recognizer.adjust_for_ambient_noise(source)
            self.mic_available = True
        except Exception as e:
            print(f"麥克風初始化失敗: {e}")
            self.mic_available = False
    
    def create_widgets(self):
        """建立UI元件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 標題
        title_label = ttk.Label(main_frame, text="簡化版語音轉文字程式", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 錄音區域
        if self.mic_available:
            record_frame = ttk.LabelFrame(main_frame, text="語音錄製", padding="10")
            record_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
            
            self.record_button = ttk.Button(record_frame, text="錄製語音", command=self.record_audio)
            self.record_button.grid(row=0, column=0, padx=(0, 10))
            
            self.record_status = ttk.Label(record_frame, text="準備錄音")
            self.record_status.grid(row=0, column=1)
        
        # 檔案上傳區域
        upload_frame = ttk.LabelFrame(main_frame, text="音訊檔案上傳", padding="10")
        upload_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.upload_button = ttk.Button(upload_frame, text="選擇音訊檔案", command=self.upload_audio_file)
        self.upload_button.grid(row=0, column=0, padx=(0, 10))
        
        self.file_label = ttk.Label(upload_frame, text="未選擇檔案")
        self.file_label.grid(row=0, column=1)
        
        # 語言選擇
        lang_frame = ttk.LabelFrame(main_frame, text="語言設定", padding="10")
        lang_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.language_var = tk.StringVar(value="zh-TW")
        ttk.Radiobutton(lang_frame, text="中文", variable=self.language_var, value="zh-TW").grid(row=0, column=0, padx=(0, 10))
        ttk.Radiobutton(lang_frame, text="英文", variable=self.language_var, value="en-US").grid(row=0, column=1)
        
        # 轉換按鈕
        convert_button = ttk.Button(main_frame, text="開始轉換", command=self.convert_speech_to_text)
        convert_button.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
        # 結果顯示區域
        result_frame = ttk.LabelFrame(main_frame, text="轉換結果", padding="10")
        result_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        
        self.text_result = scrolledtext.ScrolledText(result_frame, width=60, height=12)
        self.text_result.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 導出按鈕
        export_frame = ttk.Frame(main_frame)
        export_frame.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(export_frame, text="導出為TXT", command=self.export_txt).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(export_frame, text="清除內容", command=self.clear_text).grid(row=0, column=1)
        
        # 配置權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
    
    def record_audio(self):
        """錄製音訊"""
        if not self.mic_available:
            messagebox.showerror("錯誤", "麥克風不可用")
            return
        
        self.record_status.config(text="請說話...")
        self.record_button.config(state="disabled")
        
        def record():
            try:
                with self.microphone as source:
                    self.recognizer.adjust_for_ambient_noise(source)
                    audio = self.recognizer.listen(source, timeout=10, phrase_time_limit=10)
                    self.audio_data = audio
                    self.root.after(0, lambda: self.record_status.config(text="錄音完成"))
                    self.root.after(0, lambda: self.record_button.config(state="normal"))
            except sr.WaitTimeoutError:
                self.root.after(0, lambda: self.record_status.config(text="錄音超時"))
                self.root.after(0, lambda: self.record_button.config(state="normal"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"錄音失敗: {e}"))
                self.root.after(0, lambda: self.record_button.config(state="normal"))
        
        threading.Thread(target=record, daemon=True).start()
    
    def upload_audio_file(self):
        """上傳音訊檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇音訊檔案",
            filetypes=[
                ("WAV檔案", "*.wav"),
                ("所有檔案", "*.*")
            ]
        )
        
        if file_path:
            self.file_label.config(text=os.path.basename(file_path))
            
            try:
                with sr.AudioFile(file_path) as source:
                    self.audio_data = self.recognizer.record(source)
                messagebox.showinfo("成功", "音訊檔案載入成功")
            except Exception as e:
                messagebox.showerror("錯誤", f"載入音訊檔案失敗: {e}")
    
    def convert_speech_to_text(self):
        """轉換語音為文字"""
        if not self.audio_data:
            messagebox.showwarning("警告", "請先錄音或上傳音訊檔案")
            return
        
        self.text_result.delete(1.0, tk.END)
        self.text_result.insert(tk.END, "正在轉換中，請稍候...\n")
        self.root.update()
        
        def convert():
            try:
                language = self.language_var.get()
                text = self.recognizer.recognize_google(self.audio_data, language=language)
                self.root.after(0, lambda: self.update_result(text))
            except sr.UnknownValueError:
                self.root.after(0, lambda: self.update_result("無法識別語音內容"))
            except sr.RequestError as e:
                self.root.after(0, lambda: self.update_result(f"請求失敗: {e}"))
            except Exception as e:
                self.root.after(0, lambda: self.update_result(f"轉換失敗: {e}"))
        
        threading.Thread(target=convert, daemon=True).start()
    
    def update_result(self, text):
        """更新結果顯示"""
        self.text_result.delete(1.0, tk.END)
        self.text_result.insert(tk.END, text)
    
    def clear_text(self):
        """清除文字內容"""
        self.text_result.delete(1.0, tk.END)
    
    def export_txt(self):
        """導出為TXT檔案"""
        content = self.text_result.get(1.0, tk.END).strip()
        if not content:
            messagebox.showwarning("警告", "沒有內容可以導出")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="儲存TXT檔案",
            defaultextension=".txt",
            filetypes=[("文字檔案", "*.txt"), ("所有檔案", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"語音轉文字結果\n")
                    f.write(f"轉換時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write("-" * 50 + "\n")
                    f.write(content)
                messagebox.showinfo("成功", f"已成功導出至: {file_path}")
            except Exception as e:
                messagebox.showerror("錯誤", f"導出失敗: {e}")

def main():
    """主函數"""
    root = tk.Tk()
    app = SimpleSpeechToTextApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
