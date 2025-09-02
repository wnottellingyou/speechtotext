#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
安裝OpenAI套件腳本
"""

import subprocess
import sys

def install_openai():
    """安裝OpenAI套件"""
    try:
        print("正在安裝OpenAI套件...")
        result = subprocess.run([sys.executable, "-m", "pip", "install", "openai"], 
                              capture_output=True, text=True, check=True)
        print("OpenAI套件安裝成功！")
        print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"安裝失敗: {e}")
        print(f"錯誤輸出: {e.stderr}")
        return False
    except Exception as e:
        print(f"未知錯誤: {e}")
        return False

if __name__ == "__main__":
    success = install_openai()
    if success:
        print("\n安裝完成！現在可以使用AI摘要功能了。")
    else:
        print("\n安裝失敗，請手動執行: pip install openai")
    
    input("按Enter鍵退出...")
