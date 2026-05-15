import subprocess
import tkinter as tk
from tkinter import messagebox
import time
import os
import threading
import ctypes
import pyperclip
from datetime import datetime
from PIL import ImageGrab

try:
    import pyautogui
    PYAUTOGUI_OK = True
except ImportError:
    PYAUTOGUI_OK = False

import threading as _threading
_stop_event = _threading.Event()

def _esc_monitor():
    try:
        import keyboard
        keyboard.wait('esc')
        _stop_event.set()
        messagebox.showinfo("強制終了", "ESCキーが押されました。\nプログラムを終了します。")
        os._exit(0)
    except ImportError:
        pass

_esc_thread = _threading.Thread(target=_esc_monitor, daemon=True)
_esc_thread.start()

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

URLS = [
    ("01:宇都宮",   "https://www.hug-globalkidsmethod.link/hug/wm/"),
    ("02:埼玉群馬", "https://www.hug-globalkidsmethod02.link/hug/wm/"),
    ("03:茨城千葉", "https://www.hug-globalkidsmethod03.link/hug/wm/"),
    ("04:その他",   "https://www.hug-globalkidsmethod04.link/hug/wm/"),
    ("05:栃木1",    "https://www.hug-globalkidsmethod05.link/hug/wm/"),
    ("06:栃木2",    "https://www.hug-globalkidsmethod06.link/hug/wm/"),
]

WORK_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "算定区分ダウンロード作業")

LOGIN_BTN_IMG    = os.path.join(WORK_DIR, "login_btn.png")
SEIKYU_IMG       = os.path.join(WORK_DIR, "seikyu_menu.png")
KOKUHOREN_IMG    = os.path.join(WORK_DIR, "kokuhoren.png")
SHISETSU_IMG     = os.path.join(WORK_DIR, "shisetsu_dd.png")
HYOJI_IMG        = os.path.join(WORK_DIR, "hyoji_btn.png")
SKIP_UPDATE_IMG  = os.path.join(WORK_DIR, "skip_update.png")
SHUKEI_IMG       = os.path.join(WORK_DIR, "shukei_link.png")
CHK_ZERO_IMG     = os.path.join(WORK_DIR, "chk_zero.png")
CHK_UCHIWAKE_IMG = os.path.join(WORK_DIR, "chk_uchiwake.png")
CHK_FURIGANA_IMG = os.path.join(WORK_DIR, "chk_furigana.png")
CHK_SERVICE_IMG  = os.path.join(WORK_DIR, "chk_service.png")
HYOJI_KIRIKAE_IMG= os.path.join(WORK_DIR, "hyoji_kirikae.png")
CSV_BTN_IMG      = os.path.join(WORK_DIR, "csv_btn.png")
CANCEL_BTN_IMG   = os.path.join(WORK_DIR, "cancel_btn.png")

LOG_DIR  = WORK_DIR
LOG_FILE = os.path.join(WORK_DIR, "download_result.txt")

# テストモード状態
test_mode = False

def find_chrome():
    for path in CHROME_PATHS:
        if os.path.exists(path):
            return path
    return None

def click_image(img_path, confidence=0.8, timeout=15):
    start = time.time()
    while time.time() - start < timeout:
        for conf in [confidence, confidence-0.1, confidence-0.2]:
            try:
                loc = pyautogui.locateCenterOnScreen(img_path, confidence=max(conf, 0.5), grayscale=False)
                if loc:
                    pyautogui.click(loc.x, loc.y)
                    return True
            except Exception:
                pass
        time.sleep(0.5)
    return False

def find_image(img_path, timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        for conf in [0.8, 0.7, 0.6, 0.5]:
            try:
                loc = pyautogui.locateCenterOnScreen(img_path, confidence=conf, grayscale=False)
                if loc:
                    return loc
            except Exception:
                pass
        time.sleep(0.5)
    return None

def focus_chrome():
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_size_t, ctypes.c_size_t)
    all_windows = []
    def enum_cb(hwnd, _):
        if ctypes.windll.user32.IsWindowVisible(hwnd):
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buf, length + 1)
                all_windows.append((hwnd, buf.value))
        return True
    cb = EnumWindowsProc(enum_cb)
    ctypes.windll.user32.EnumWindows(cb, 0)
    hwnd = None
    for h, t in all_windows:
        if any(kw in t for kw in ["ログイン", "ホーム | 成長", "HUG", "hug-global", "国保連"]):
            hwnd = h
            break
    if not hwnd:
        for h, t in all_windows:
            if "Chrome" in t or "Google" in t:
                hwnd = h
                break
    if hwnd:
        ctypes.windll.user32.ShowWindow(hwnd, 3)
        time.sleep(0.3)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.3)
        return True
    return False

def minimize_gui():
    try:
        root_hwnd = ctypes.windll.user32.FindWindowW(None, "HUG サイト選択")
        if root_hwnd:
            ctypes.windll.user32.ShowWindow(root_hwnd, 6)
    except Exception:
        pass

def restore_gui():
    try:
        root_hwnd = ctypes.windll.user32.FindWindowW(None, "HUG サイト選択")
        if root_hwnd:
            ctypes.windll.user32.ShowWindow(root_hwnd, 9)
            ctypes.windll.user32.SetForegroundWindow(root_hwnd)
    except Exception:
        pass
def is_shukei_grayed_out(loc):
    region = (loc.x - 30, loc.y - 10, loc.x + 30, loc.y + 10)
    img = ImageGrab.grab(bbox=region)
    import numpy as np
    arr = np.array(img)
    r, g, b = arr[:,:,0].mean(), arr[:,:,1].mean(), arr[:,:,2].mean()
    diff = max(abs(r-g), abs(g-b), abs(r-b))
    return diff < 20 and 130 < r < 220

def find_shukei():
    import numpy as np
    import traceback

    _dbg("find_shukei 開始")

    try:
        focus_chrome()
        time.sleep(0.5)

        # テンプレートマッチングで集計一覧の印刷リンクを直接クリック
        # pngが広く切り取られているため中央ではなく左上付近をクリック
        box = None
        start_t = time.time()
        while time.time() - start_t < 10:
            for conf in [0.8, 0.7, 0.6, 0.5]:
                try:
                    box = pyautogui.locateOnScreen(SHUKEI_IMG, confidence=conf, grayscale=False)
                    if box:
                        break
                except Exception:
                    pass
            if box:
                break
            time.sleep(0.5)
        if box:
            cx = box.left + box.width // 2   # ← ここ
            cy = box.top + box.height // 2    # ← ここ
            _dbg(f"→ SHUKEI_IMG 検出 box=({box.left},{box.top},{box.width},{box.height}) クリック({cx},{cy})")
            pyautogui.click(cx, cy)
            time.sleep(1.0)
            return True

        _dbg("SHUKEI_IMG 未検出 → Ctrl+F差分検出にフォールバック")

        # Chromeウィンドウ中央をクリックしてフォーカスを確実に取得
        sw, sh = pyautogui.size()
        pyautogui.click(sw // 2, sh // 2)
        time.sleep(0.5)

        before = np.array(ImageGrab.grab())
        _dbg(f"before shape={before.shape}")

        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1.5)
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.3)
        pyperclip.copy('集計一覧の印刷')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2.0)

        with_hl = np.array(ImageGrab.grab())

        pyautogui.press('escape')
        time.sleep(0.5)

        diff = np.abs(with_hl.astype(int) - before.astype(int)).sum(axis=2)
        diff[:int(sh * 0.15)] = 0   # 上部15%除外（ブラウザUI）

        row_sums = diff.sum(axis=1)
        peak_row = int(np.argmax(row_sums))
        peak_val = float(row_sums[peak_row])

        band_top = max(0, peak_row - 15)
        band_bot = min(sh, peak_row + 15)
        col_sums = diff[band_top:band_bot, :].sum(axis=0)
        peak_col = int(np.argmax(col_sums))
        cx, cy = peak_col, peak_row

        _dbg(f"peak_val={peak_val:.0f} row={peak_row} cx={cx} cy={cy} sw={sw} sh={sh}")

        from PIL import Image as _PILImg
        _PILImg.fromarray(before).save(os.path.join(WORK_DIR, "debug_before.png"))
        _PILImg.fromarray(with_hl).save(os.path.join(WORK_DIR, "debug_with_hl.png"))

        if peak_val < 50:
            _dbg("→ peak_val 閾値未満 スキップ")
            return False

        _dbg(f"→ クリック cx={cx} cy={cy}")
        pyautogui.click(cx, cy)
        time.sleep(1.0)
        return True

    except Exception:
        _dbg(f"例外発生:\n{traceback.format_exc()}")
        return False

def write_log(log_lines):
    os.makedirs(LOG_DIR, exist_ok=True)
    now = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    with open(LOG_FILE, 'w', encoding='utf-8') as f:
        f.write(f"ダウンロード結果　{now}\n")
        f.write("=" * 40 + "\n")
        for line in log_lines:
            f.write(line + "\n")

_DBG_FILE = os.path.join(WORK_DIR, "debug_shukei.txt")

def _dbg(msg):
    try:
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        with open(_DBG_FILE, "a", encoding="utf-8") as f:
            f.write(line)
        print(line, end="", flush=True)
    except Exception as e:
        print(f"_dbg error: {e} | {msg}", flush=True)

def do_one_shisetsu():
    """現在選択されている施設の処理（表示変更→集計→CSV→キャンセル）。成功=True"""
    import traceback
    def click_hyoji():
        loc = find_image(HYOJI_IMG, timeout=5)
        if loc:
            pyautogui.click(loc.x, loc.y)
            return True
        return False

    _dbg("do_one_shisetsu 開始")
    try:
        if not click_hyoji():
            _dbg("click_hyoji 失敗（HYOJI_IMG 未検出）")
            return False
        _dbg("click_hyoji 成功")
        time.sleep(5)

        pyautogui.hotkey('ctrl', 'end')
        time.sleep(1)
        _dbg("ctrl+end 完了")

        # SKIP_UPDATE: 信頼度0.92以上のみ（低信頼度での誤マッチ防止）
        skip_loc = None
        try:
            skip_loc = pyautogui.locateCenterOnScreen(SKIP_UPDATE_IMG, confidence=0.92, grayscale=False)
        except Exception:
            pass
        if skip_loc:
            _dbg(f"SKIP_UPDATE クリック ({skip_loc.x},{skip_loc.y})")
            pyautogui.click(skip_loc.x, skip_loc.y)
            time.sleep(2)

        focus_chrome()
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'home')
        time.sleep(1.0)
        _dbg("ctrl+home 完了 → find_shukei 呼び出し直前")

        if not find_shukei():
            _dbg("find_shukei False → return False")
            return False
    except Exception:
        _dbg(f"例外発生:\n{traceback.format_exc()}")
        return False

    # find_shukei内でクリック済みなので直接チェック画面へ
    time.sleep(2)

    first_chk = find_image(CHK_ZERO_IMG, timeout=3)
    if not first_chk:
        focus_chrome()
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'home')
        time.sleep(0.8)
        return False

    for img in [CHK_ZERO_IMG, CHK_UCHIWAKE_IMG, CHK_FURIGANA_IMG, CHK_SERVICE_IMG]:
        loc = find_image(img, timeout=5)
        if loc:
            pyautogui.click(loc.x, loc.y)
            time.sleep(0.3)

    if not click_image(HYOJI_KIRIKAE_IMG, timeout=10):
        return False
    time.sleep(3)

    if not click_image(CSV_BTN_IMG, timeout=10):
        return False

    dl_folder = os.path.expandvars(r"%USERPROFILE%\Downloads")
    before = set(os.listdir(dl_folder))
    start_dl = time.time()
    while time.time() - start_dl < 30:
        after = set(os.listdir(dl_folder))
        done = [f for f in (after - before) if not f.endswith('.crdownload')]
        if done:
            break
        time.sleep(1)
    time.sleep(1)

    if not click_image(CANCEL_BTN_IMG, timeout=10):
        return False
    time.sleep(2)
    return True


def focus_shisetsu_dd():
    pyautogui.hotkey('ctrl', 'home')
    time.sleep(0.5)
    cur = find_image(SHISETSU_IMG, timeout=5)
    if not cur:
        return False
    pyautogui.click(cur.x, cur.y)
    time.sleep(0.4)
    pyautogui.press('tab')
    time.sleep(0.3)
    return True


def _find_chrome_hwnd():
    result = []
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_size_t, ctypes.c_size_t)
    def _cb(hwnd, _):
        if ctypes.windll.user32.IsWindowVisible(hwnd):
            length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, buf, length + 1)
                if any(kw in buf.value for kw in ["Chrome", "Google", "hug-global"]):
                    result.append(hwnd)
        return True
    ctypes.windll.user32.EnumWindows(EnumWindowsProc(_cb), 0)
    return result[0] if result else None


def process_site(url, label, year, month):
    """1サイト分の処理。処理件数を返す。エラー時は-1。"""
    _dbg(f"process_site 開始: {label}")
    chrome_path = find_chrome()
    if not chrome_path:
        return -1

    try:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

        subprocess.Popen([chrome_path, "--start-maximized", url])
        time.sleep(5)

        minimize_gui()
        time.sleep(0.3)
        focus_chrome()
        time.sleep(0.5)

        # Chromeウィンドウを確実に最大化
        chrome_hwnd = _find_chrome_hwnd()
        if chrome_hwnd:
            ctypes.windll.user32.ShowWindow(chrome_hwnd, 3)  # 3 = SW_MAXIMIZE
            time.sleep(0.5)
            ctypes.windll.user32.SetForegroundWindow(chrome_hwnd)
            time.sleep(0.3)

        if click_image(LOGIN_BTN_IMG, timeout=6):
            time.sleep(4)
            minimize_gui()
            focus_chrome()
            time.sleep(0.3)
            time.sleep(0.5)

        focus_chrome()
        time.sleep(0.3)
        time.sleep(0.5)

        seikyu_pos = find_image(SEIKYU_IMG, timeout=10)
        if not seikyu_pos:
            return -1
        pyautogui.moveTo(seikyu_pos.x, seikyu_pos.y, duration=0.3)
        time.sleep(1.5)

        if not click_image(KOKUHOREN_IMG, timeout=10):
            return -1
        time.sleep(3)

        s_pos = find_image(SHISETSU_IMG, timeout=10)
        if not s_pos:
            return -1

        pyautogui.click(s_pos.x, s_pos.y)
        time.sleep(0.4)
        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('tab')
        time.sleep(0.4)
        pyautogui.press('home')
        time.sleep(0.3)
        for _ in range(int(year) - 2015):
            pyautogui.press('down')
            time.sleep(0.1)
        time.sleep(0.3)

        pyautogui.press('tab')
        time.sleep(0.4)
        pyautogui.press('home')
        time.sleep(0.3)
        for _ in range(int(month) - 1):
            pyautogui.press('down')
            time.sleep(0.1)
        time.sleep(0.3)

        if not focus_shisetsu_dd():
            return -1

        total = 0

        if test_mode:
            # テストモード：最初の施設と最後の施設だけ処理
            # 最後の施設（end）
            pyautogui.press('end')
            time.sleep(0.3)
            if do_one_shisetsu():
                total += 1

            # 最初の施設（home）
            if not focus_shisetsu_dd():
                return total
            pyautogui.press('home')
            time.sleep(0.3)
            if do_one_shisetsu():
                total += 1

        else:
            # 通常モード：末尾から全施設ループ
            pyautogui.press('end')
            time.sleep(0.3)

            while True:
                if do_one_shisetsu():
                    total += 1

                # 次の施設へ
                focus_chrome()
                time.sleep(0.3)
                if not focus_shisetsu_dd():
                    break

                pyautogui.press('end')
                time.sleep(0.2)

                import numpy as np
                before_ss = pyautogui.screenshot()
                pyautogui.press('up')
                time.sleep(0.2)
                after_ss = pyautogui.screenshot()
                b = np.array(before_ss)
                a = np.array(after_ss)
                if np.array_equal(b, a):
                    break
                else:
                    pyautogui.press('down')
                    time.sleep(0.2)

        return total

    except Exception:
        return -1


def launch_and_login(url, year, month, label=""):
    chrome_path = find_chrome()
    if not chrome_path:
        messagebox.showerror("エラー", "Chromeが見つかりませんでした。")
        return
    if not PYAUTOGUI_OK:
        messagebox.showerror("エラー", "pip install pyautogui pillow opencv-python")
        return

    def run():
        total = process_site(url, label, year, month)
        mode_str = "【テスト】" if test_mode else ""
        log_lines = [f"{mode_str}{label}　：{total if total >= 0 else 'エラー'}件　{datetime.now().strftime('%Y/%m/%d')}"]
        write_log(log_lines)
        restore_gui()
        completion_win = tk.Toplevel()
        completion_win.title("完了")
        completion_win.attributes("-topmost", True)
        completion_win.geometry("300x100")
        tk.Label(completion_win, text=f"{mode_str}処理完了\n{label}：{total}件", font=("", 11), pady=20).pack()
        completion_win.after(5000, completion_win.destroy)

    threading.Thread(target=run, daemon=True).start()


def launch_all(year, month):
    chrome_path = find_chrome()
    if not chrome_path:
        messagebox.showerror("エラー", "Chromeが見つかりませんでした。")
        return
    if not PYAUTOGUI_OK:
        messagebox.showerror("エラー", "pip install pyautogui pillow opencv-python")
        return

    def run():
        log_lines = []
        date_str = datetime.now().strftime("%Y/%m/%d")
        mode_str = "【テスト】" if test_mode else ""
        for label, url in URLS:
            total = process_site(url, label, year, month)
            count_str = f"{total}件" if total >= 0 else "エラー"
            log_lines.append(f"{mode_str}{label}　：{count_str}　{date_str}")
        write_log(log_lines)
        restore_gui()
        messagebox.showinfo("正常終了", "正常に終了しました。")

    threading.Thread(target=run, daemon=True).start()


def main():
    global test_mode
    now = datetime.now()
    root = tk.Tk()
    root.title("HUG サイト選択")
    root.geometry("320x640")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    # 年月入力
    frame_ym = tk.LabelFrame(root, text="請求年月", font=("", 9), padx=8, pady=6)
    frame_ym.pack(fill="x", padx=12, pady=8)
    tk.Label(frame_ym, text="年:").grid(row=0, column=0, sticky="e")
    year_var = tk.StringVar(value=str(now.year))
    tk.Spinbox(frame_ym, from_=2015, to=2099, textvariable=year_var, width=6, font=("", 10)).grid(row=0, column=1, padx=4)
    tk.Label(frame_ym, text="月:").grid(row=0, column=2, sticky="e")
    month_var = tk.StringVar(value=str(now.month))
    tk.Spinbox(frame_ym, from_=1, to=12, textvariable=month_var, width=4, font=("", 10)).grid(row=0, column=3, padx=4)

    tk.Label(root, text="開くサイトを選択してください", font=("", 11, "bold")).pack(pady=6)

    tk.Button(
        root,
        text="07:全て（01～06を順番に処理）",
        width=25,
        height=2,
        font=("", 10, "bold"),
        bg="#d9534f",
        fg="white",
        command=lambda: launch_all(year_var.get(), month_var.get())
    ).pack(pady=3)

    for label, url in URLS:
        tk.Button(
            root,
            text=label,
            width=25,
            height=2,
            font=("", 10),
            command=lambda u=url, l=label: launch_and_login(u, year_var.get(), month_var.get(), l)
        ).pack(pady=3)

    if not PYAUTOGUI_OK:
        tk.Label(root, text="⚠ pip install pyautogui pillow opencv-python", fg="red", font=("", 9)).pack(pady=5)

    # テストモードON/OFFボタン（左下）
    test_btn_var = tk.StringVar(value="🧪 テストモード: OFF")
    def toggle_test():
        global test_mode
        test_mode = not test_mode
        if test_mode:
            test_btn_var.set("🧪 テストモード: ON")
            test_btn.config(bg="#f0ad4e", fg="white")
        else:
            test_btn_var.set("🧪 テストモード: OFF")
            test_btn.config(bg="#e0e0e0", fg="black")

    test_btn = tk.Button(
        root,
        textvariable=test_btn_var,
        font=("", 9, "bold"),
        bg="#e0e0e0",
        fg="black",
        command=toggle_test
    )
    test_btn.pack(side="left", anchor="sw", padx=10, pady=8)

    root.mainloop()

if __name__ == "__main__":
    main()
