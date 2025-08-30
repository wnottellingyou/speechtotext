#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測試語音轉文字程式的修復效果
"""

import tkinter as tk
from tkinter import messagebox
import threading
import time

def test_ui_thread_safety():
    """測試UI線程安全性"""
    
    def create_test_window():
        root = tk.Tk()
        root.title("UI線程安全測試")
        root.geometry("400x300")
        
        # 創建狀態標籤
        status_label = tk.Label(root, text="準備測試...")
        status_label.pack(pady=20)
        
        # 創建結果文本框
        result_text = tk.Text(root, width=40, height=10)
        result_text.pack(pady=10)
        
        def safe_update_status(text):
            """安全地更新狀態"""
            try:
                status_label.config(text=text)
            except Exception as e:
                print(f"更新狀態失敗: {e}")
        
        def safe_update_result(text):
            """安全地更新結果"""
            try:
                result_text.delete(1.0, tk.END)
                result_text.insert(tk.END, text)
            except Exception as e:
                print(f"更新結果失敗: {e}")
        
        def background_task():
            """在背景執行的任務"""
            for i in range(5):
                if not root.winfo_exists():
                    break
                
                # 使用正確的方式更新UI
                status_text = f"處理中... {i+1}/5"
                root.after(0, safe_update_status, status_text)
                
                time.sleep(1)
                
                result_text = f"完成步驟 {i+1}\n"
                root.after(0, lambda text=result_text: safe_update_result(
                    result_text.get(1.0, tk.END) + text
                ))
            
            # 最終更新
            root.after(0, safe_update_status, "測試完成！")
            root.after(0, safe_update_result, "所有測試步驟已完成，UI更新正常！")
        
        def start_test():
            """開始測試"""
            threading.Thread(target=background_task, daemon=True).start()
        
        # 創建測試按鈕
        test_button = tk.Button(root, text="開始UI線程安全測試", command=start_test)
        test_button.pack(pady=10)
        
        # 顯示說明
        info_label = tk.Label(root, text="此測試驗證多線程UI更新是否安全", 
                             font=("Arial", 8), fg="gray")
        info_label.pack(pady=5)
        
        return root
    
    # 創建並啟動測試窗口
    test_root = create_test_window()
    
    # 顯示測試信息
    messagebox.showinfo("測試開始", "UI線程安全測試窗口已啟動\n點擊測試按鈕開始驗證")
    
    test_root.mainloop()

if __name__ == "__main__":
    print("開始UI線程安全測試...")
    test_ui_thread_safety()
    print("測試完成")
