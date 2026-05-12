import subprocess
import tkinter as tk
from tkinter import messagebox
import time
import os
import sys
import threading

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

DEBUG_PORT = 9222

URLS = [
    ("01:宇都宮",   "https://www.hug-globalkidsmethod.link/hug/wm/"),
    ("02:埼玉群馬", "https://www.hug-globalkidsmethod02.link/hug/wm/"),
    ("03:茨城千葉", "https://www.hug-globalkidsmethod03.link/hug/wm/"),
    ("04:その他",   "https://www.hug-globalkidsmethod04.link/hug/wm/"),
    ("05:栃木1",    "https://www.hug-globalkidsmethod05.link/hug/wm/"),
    ("06:栃木2",    "https://www.hug-globalkidsmethod06.link/hug/wm/"),
]

def find_chrome():
    for path in CHROME_PATHS:
        if os.path.exists(path):
            return path
    return None

def launch_chrome(url):
    chrome_path = find_chrome()
    if not chrome_path:
        messagebox.showerror("エラー", "Chromeが見つかりませんでした。")
        return

    user_data_dir = os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data")

    try:
        subprocess.Popen([
            chrome_path,
            f"--remote-debugging-port={DEBUG_PORT}",
            f"--user-data-dir={user_data_dir}",
            url
        ])
    except Exception as e:
        messagebox.showerror("エラー", f"Chrome起動に失敗しました:\n{e}")
        return

    if not SELENIUM_OK:
        return

    def auto_login():
        try:
            time.sleep(5)
            options = Options()
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
            driver = webdriver.Chrome(options=options)
            wait = WebDriverWait(driver, 10)
            try:
                login_btn = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[name="submit"]'))
                )
                login_btn.click()
            except Exception:
                pass
        except Exception:
            pass

    threading.Thread(target=auto_login, daemon=True).start()

def main():
    root = tk.Tk()
    root.title("HUG サイト選択")
    root.geometry("320x380")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    tk.Label(root, text="開くサイトを選択してください", font=("", 11, "bold")).pack(pady=15)

    for label, url in URLS:
        btn = tk.Button(
            root,
            text=label,
            width=25,
            height=2,
            font=("", 10),
            command=lambda u=url: launch_chrome(u)
        )
        btn.pack(pady=4)

    root.mainloop()

if __name__ == "__main__":
    main()
