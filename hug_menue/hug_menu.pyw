# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import messagebox
import datetime
import threading
import time
import psutil

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

CHROME_EXE     = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_PROFILE = "Profile 2"
USER_DATA_DIR  = r"C:\Users\owner\AppData\Local\Google\Chrome\User Data"

# ログインボタンとして検出する XPath（広め）
LOGIN_XPATH = (
    "//button[@type='submit'] | //input[@type='submit'] | "
    "//button[contains(normalize-space(.),'ログイン')] | "
    "//button[contains(normalize-space(.),'ろぐいん')] | "
    "//button[contains(translate(normalize-space(.),'login LOGIN Login','LOGINLOGINLOGIN'),'LOGIN')] | "
    "//a[contains(normalize-space(.),'ログイン')] | "
    "//input[@value='ログイン'] | "
    "//input[contains(translate(@value,'login LOGIN','LOGINLOGIN'),'LOGIN')] | "
    "//*[@id='login-btn' or @id='loginBtn' or @id='submit' or @id='login_button'] | "
    "//*[contains(@class,'login-btn') or contains(@class,'loginBtn') or contains(@class,'btn-login')]"
)

_driver = None
_driver_lock = threading.Lock()


def _get_driver(first_url=None):
    global _driver
    if _driver is not None:
        try:
            _ = _driver.window_handles
            return _driver, False
        except Exception:
            _driver = None

    options = Options()
    options.binary_location = CHROME_EXE
    options.add_argument(f"--user-data-dir={USER_DATA_DIR}")
    options.add_argument(f"--profile-directory={CHROME_PROFILE}")
    options.add_argument("--no-restore-last-session")
    options.add_argument("--disable-session-crashed-bubble")
    options.add_argument("--hide-crash-restore-bubble")
    options.add_argument("--new-window")

    service = Service(ChromeDriverManager().install())
    _driver = webdriver.Chrome(service=service, options=options)

    if first_url:
        _driver.get(first_url)

    return _driver, True


def _open_and_login(url):
    with _driver_lock:
        driver, is_first = _get_driver(first_url=url)
        if not is_first:
            driver.switch_to.new_window("tab")
            driver.get(url)
        # is_first の場合は Chrome 起動時に既に url を開いている
        # ページ読み込み完了を待つ
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(1.5)  # オートフィル完了を待つ
        try:
            btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, LOGIN_XPATH))
            )
            btn.click()
        except Exception:
            pass  # ログインボタンなし or 既にログイン済み


def open_url(url):
    threading.Thread(target=_open_and_login, args=(url,), daemon=True).start()


def open_all_urls(urls):
    def _task():
        for url in urls:
            _open_and_login(url)
    threading.Thread(target=_task, daemon=True).start()


class HugSiteMenu:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("HUG サイト選択")
        self.root.resizable(False, False)

        now = datetime.date.today()
        self.year_var = tk.IntVar(value=now.year)
        self.month_var = tk.IntVar(value=now.month)
        self.test_mode = tk.BooleanVar(value=False)

        self.URLS = {
            "01": "https://www.hug-globalkidsmethod.link/hug/wm/",
            "02": "https://www.hug-globalkidsmethod02.link/hug/wm/",
            "03": "https://www.hug-globalkidsmethod03.link/hug/wm/",
            "04": "https://www.hug-globalkidsmethod04.link/hug/wm/",
            "05": "https://www.hug-globalkidsmethod05.link/hug/wm/",
            "06": "https://www.hug-globalkidsmethod06.link/hug/wm/",
        }

        self._build_ui()
        self._center_window(300, 500)

    def _center_window(self, w, h):
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        root = self.root

        billing_frame = tk.LabelFrame(root, text="請求年月", padx=8, pady=4)
        billing_frame.pack(fill="x", padx=10, pady=(10, 4))

        tk.Label(billing_frame, text="年:").pack(side="left")
        tk.Spinbox(
            billing_frame, from_=2000, to=2099,
            textvariable=self.year_var, width=6, format="%4.0f"
        ).pack(side="left")

        tk.Label(billing_frame, text=" 月:").pack(side="left")
        tk.Spinbox(
            billing_frame, from_=1, to=12,
            textvariable=self.month_var, width=3, format="%2.0f"
        ).pack(side="left")

        tk.Label(
            root, text="開くサイトを選択してください",
            font=("", 11, "bold")
        ).pack(pady=(6, 8))

        btn_frame = tk.Frame(root)
        btn_frame.pack(fill="x", padx=10)

        sites = [
            ("07:全て（01～06を順番に処理）", "07", True),
            ("01:宇都宮",   "01", False),
            ("02:埼玉群馬", "02", False),
            ("03:茨城千葉", "03", False),
            ("04:その他",   "04", False),
            ("05:栃木1",    "05", False),
            ("06:栃木2",    "06", False),
        ]

        for label, code, is_all in sites:
            bg = "#c0392b" if is_all else "#f0f0f0"
            fg = "white"   if is_all else "black"
            tk.Button(
                btn_frame, text=label,
                bg=bg, fg=fg,
                activebackground="#e74c3c" if is_all else "#d0d0d0",
                activeforeground=fg,
                relief="raised", width=30, height=1,
                font=("", 10),
                command=lambda c=code: self._on_site_select(c)
            ).pack(fill="x", pady=2)

        self.test_btn = tk.Button(
            root, text="テストモード:OFF",
            relief="sunken", anchor="w", font=("", 9),
            command=self._toggle_test_mode
        )
        self.test_btn.pack(side="bottom", fill="x")

    def _on_site_select(self, code):
        if code == "07":
            open_all_urls([self.URLS[c] for c in ["01", "02", "03", "04", "05", "06"]])
        else:
            open_url(self.URLS[code])

    def _toggle_test_mode(self):
        self.test_mode.set(not self.test_mode.get())
        state = "ON" if self.test_mode.get() else "OFF"
        self.test_btn.config(text=f"テストモード:{state}")

    def _on_close(self):
        global _driver
        if _driver is not None:
            try:
                _driver.quit()
            except Exception:
                pass
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()


def chrome_is_running():
    return any(p.name().lower() == "chrome.exe" for p in psutil.process_iter(["name"]))


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    while chrome_is_running():
        ok = messagebox.askretrycancel(
            "Chrome が起動中",
            "Chrome が起動しています。\nChromeをすべて閉じてから「再試行」を押してください。",
        )
        if not ok:
            root.destroy()
            raise SystemExit
    root.destroy()
    HugSiteMenu().run()
