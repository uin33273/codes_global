import json
import logging
import os
import time
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date
from pathlib import Path
from playwright.sync_api import sync_playwright, Page, TimeoutError as PWTimeoutError

HUG_JSON_PATH = Path(r'C:\Users\owner\Desktop\HUG.json')
DOWNLOAD_DIR = Path(r'C:\Users\owner\Downloads\算定区分データ')

SITES = [
    {'id': '01', 'name': '宇都宮', 'url': 'https://www.hug-globalkidsmethod.link/hug/wm/'},
    {'id': '02', 'name': '埼玉群馬', 'url': 'https://www.hug-globalkidsmethod02.link/hug/wm/'},
    {'id': '03', 'name': '茨城千葉', 'url': 'https://www.hug-globalkidsmethod03.link/hug/wm/'},
    {'id': '04', 'name': 'その他', 'url': 'https://www.hug-globalkidsmethod04.link/hug/wm/'},
    {'id': '05', 'name': '栃木1', 'url': 'https://www.hug-globalkidsmethod05.link/hug/wm/'},
    {'id': '06', 'name': '栃木2', 'url': 'https://www.hug-globalkidsmethod06.link/hug/wm/'},
]

SELECTORS = {
    'login_id': "input[name='username']",
    'login_password': "input[name='password']",
    'login_button': 'input.btn-login',
    'facility_select': "select[name='facility']",
    'year_select': "select[name='s_year']",
    'month_select': "select[name='s_month']",
    'display_change_btn': "button.btn.btn-sm[type='submit']:not(.search)",
    'skip_update_dialog': 'a.btn.btn-md.check.clearDisabled',
    'print_list_btn': "#payment06 a[href*='payment_sum.php']",
    'check_shienshihi': 'input#selenium_sum_select',
    'check_service_code': "input[name='decision_code']",
    'display_toggle_btn': "button[name='mode'][value='print_line']",
    'csv_output_btn': 'button.btn-sm.csv',
    'cancel_btn': 'button.btn-sm.cancel',
}

MAX_RETRIES = 3


class SkipShop(Exception):
    """この店舗をスキップすべき場合に送出する"""


def ask_year_month() -> tuple[str, str]:
    result = {}
    today = date.today()
    if today.month == 1:
        default_year = str(today.year - 1)
        default_month = '12'
    else:
        default_year = str(today.year)
        default_month = str(today.month - 1).zfill(2)

    root = tk.Tk()
    root.title('対象年月の入力')
    root.resizable(False, False)

    font = ('Yu Gothic UI', 13)
    font_small = ('Yu Gothic UI', 10)

    frame = tk.Frame(root, padx=24, pady=20)
    frame.pack()

    tk.Label(frame, text='対象年月を入力してください', font=('Yu Gothic UI', 13, 'bold')).grid(
        row=0, column=0, columnspan=3, pady=(0, 16))

    tk.Label(frame, text='年', font=font).grid(row=1, column=0, sticky='e', padx=(0, 6))
    year_var = tk.StringVar(value=default_year)
    year_entry = tk.Entry(frame, textvariable=year_var, width=7, font=font, justify='center')
    year_entry.grid(row=1, column=1, sticky='w')
    tk.Label(frame, text='例: 2026', font=font_small, fg='#888').grid(row=1, column=2, sticky='w', padx=(6, 0))

    tk.Label(frame, text='月', font=font).grid(row=2, column=0, sticky='e', padx=(0, 6), pady=(10, 0))
    month_var = tk.StringVar(value=default_month)
    month_entry = tk.Entry(frame, textvariable=month_var, width=4, font=font, justify='center')
    month_entry.grid(row=2, column=1, sticky='w', pady=(10, 0))
    tk.Label(frame, text='例: 05（2桁）', font=font_small, fg='#888').grid(
        row=2, column=2, sticky='w', padx=(6, 0), pady=(10, 0))

    ok_btn = tk.Button(frame, text='次へ →', width=10, font=font, command=lambda: on_ok())
    ok_btn.grid(row=3, column=0, columnspan=3, pady=(20, 0))

    def on_ok():
        y = year_var.get().strip()
        m = month_var.get().strip().zfill(2)
        if len(y) != 4 or not y.isdigit():
            year_entry.config(bg='#ffe0e0')
            year_entry.focus_set()
            return
        if not (m.isdigit() and 1 <= int(m) <= 12):
            month_entry.config(bg='#ffe0e0')
            month_entry.focus_set()
            return
        result['year'] = y
        result['month'] = m
        root.destroy()

    def on_year_change(*_):
        year_entry.config(bg='white')
        if len(year_var.get()) == 4:
            month_entry.focus_set()

    def on_month_change(*_):
        month_entry.config(bg='white')
        if len(month_var.get()) >= 2:
            ok_btn.focus_set()

    year_var.trace_add('write', on_year_change)
    month_var.trace_add('write', on_month_change)
    root.bind('<Return>', lambda e: on_ok())
    year_entry.focus_set()
    root.mainloop()

    if 'year' not in result:
        raise SystemExit('入力がキャンセルされました')

    return (result['year'], result['month'])


def ask_sites() -> list[dict]:
    selected = {}
    root = tk.Tk()
    root.title('処理するサイトの選択')
    root.resizable(False, False)

    font = ('Yu Gothic UI', 13)
    font_bold = ('Yu Gothic UI', 13, 'bold')

    frame = tk.Frame(root, padx=24, pady=20)
    frame.pack()

    tk.Label(frame, text='処理するサイトを選択してください', font=font_bold).grid(
        row=0, column=0, columnspan=2, pady=(0, 12))

    vars_ = []
    for i, site in enumerate(SITES):
        v = tk.BooleanVar(value=True)
        vars_.append(v)
        cb = tk.Checkbutton(
            frame,
            text=f"{site['id']}  {site['name']}",
            variable=v,
            font=font,
            anchor='w',
        )
        cb.grid(row=i + 1, column=0, columnspan=2, sticky='w', pady=2)

    tk.Frame(frame, height=1, bg='#ccc').grid(
        row=len(SITES) + 1, column=0, columnspan=2, sticky='ew', pady=8)

    def select_all():
        for v in vars_:
            v.set(True)

    tk.Button(
        frame, text='07  01〜06を連続で処理', font=font, width=24, command=select_all
    ).grid(row=len(SITES) + 2, column=0, columnspan=2, pady=(0, 12))

    def on_ok():
        chosen = [SITES[i] for i, v in enumerate(vars_) if v.get()]
        if not chosen:
            return
        selected['sites'] = chosen
        root.destroy()

    tk.Button(
        frame, text='OK  開始', font=font_bold, width=24,
        bg='#4a90d9', fg='white', command=on_ok,
    ).grid(row=len(SITES) + 3, column=0, columnspan=2)

    root.bind('<Return>', lambda e: on_ok())
    root.mainloop()

    if 'sites' not in selected:
        raise SystemExit('選択がキャンセルされました')

    return selected['sites']


class ProgressWindow:
    """画面左下に常駐する進捗バー。中断ボタン／Escキーで停止要求。"""

    def __init__(self):
        self.stop_requested = False
        self.root = tk.Toplevel()
        self.root.title('ダウンロード進捗')
        self.root.resizable(False, False)
        self.root.attributes('-topmost', True)

        w, h = 400, 130
        sh = self.root.winfo_screenheight()
        self.root.geometry(f'{w}x{h}+0+{sh - h - 60}')

        font = ('Yu Gothic UI', 10)
        font_b = ('Yu Gothic UI', 11, 'bold')

        self._site_var = tk.StringVar(value='準備中...')
        self._shop_var = tk.StringVar(value='')
        self._pct_var = tk.StringVar(value='')

        tk.Label(self.root, textvariable=self._site_var, font=font_b, anchor='w').pack(
            fill='x', padx=10, pady=(6, 2))

        bar_frame = tk.Frame(self.root)
        bar_frame.pack(fill='x', padx=10)

        self._bar = ttk.Progressbar(bar_frame, length=260, mode='determinate')
        self._bar.pack(side='left')

        tk.Label(bar_frame, textvariable=self._pct_var, font=font, width=8).pack(
            side='left', padx=(6, 0))

        tk.Label(self.root, textvariable=self._shop_var, font=font, anchor='w', fg='#555').pack(
            fill='x', padx=10, pady=(2, 4))

        self._stop_btn = tk.Button(
            self.root,
            text='⛔ 中断（現在の店舗完了後に停止）',
            font=font,
            fg='white',
            bg='#c0392b',
            activebackground='#e74c3c',
            relief='flat',
            command=self._request_stop,
        )
        self._stop_btn.pack(fill='x', padx=10, pady=(0, 6))

        self.root.bind('<Escape>', lambda e: self._request_stop())
        self.root.update()

    def _request_stop(self):
        self.stop_requested = True
        self._site_var.set('⛔ 中断要求中… 現在の店舗が終わり次第停止')
        self._stop_btn.config(text='中断待機中…', state='disabled', bg='#888')
        self.root.update()

    def update(self, site_label: str, shop: str, current: int, total: int):
        if not self.stop_requested:
            self._site_var.set(site_label)
        self._shop_var.set(shop)
        self._bar['maximum'] = total
        self._bar['value'] = current
        self._pct_var.set(f'{current}/{total}')
        self.root.update()

    def close(self):
        try:
            self.root.destroy()
        except Exception:
            pass


def setup_log(year: str, month: str) -> tuple[logging.Logger, Path]:
    log_dir = DOWNLOAD_DIR / 'logs'
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"{year}{month}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger('download')
    logger.setLevel(logging.INFO)

    fmt = logging.Formatter('%(asctime)s %(message)s', datefmt='%H:%M:%S')

    fh = logging.FileHandler(log_path, encoding='utf-8')
    fh.setFormatter(fmt)

    sh = logging.StreamHandler()
    sh.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(sh)

    return logger, log_path


def main():
    year, month = ask_year_month()
    sites = ask_sites()
    logger, log_path = setup_log(year, month)

    logger.info(f'{year}年{month}月  対象: {", ".join(s["id"] + " " + s["name"] for s in sites)}')

    creds = json.loads(HUG_JSON_PATH.read_text(encoding='utf-8'))

    all_success = 0
    all_failure = []

    _root = tk.Tk()
    _root.withdraw()

    progress = ProgressWindow()

    with sync_playwright() as p:
        for site in sites:
            logger.info(f"[{site['id']}] {site['name']} ─ 処理開始")

            save_dir = DOWNLOAD_DIR / f'{year}{month}' / site['id']
            save_dir.mkdir(parents=True, exist_ok=True)

            existing = list(save_dir.glob('*.csv'))
            if existing:
                ans = messagebox.askyesno(
                    '上書き確認',
                    f"[{site['id']}] {site['name']} の保存先に\n既に {len(existing)} 件のCSVファイルがあります。\n\n"
                    f'{save_dir}\n\nこのまま続けますか？（既存ファイルは上書きされます）',
                )
                if not ans:
                    logger.info(f"  [{site['id']}] {site['name']} ─ スキップ（ユーザーが中止）")
                    continue

            csv_page_url = site['url'] + 'payment/payment.php?date_change_flg=1'

            browser = p.chromium.launch(headless=False)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            success_list = []
            failure_list = []

            try:
                page.goto(site['url'])
                page.wait_for_load_state('domcontentloaded')
                page.fill(SELECTORS['login_id'], creds['id'])
                page.fill(SELECTORS['login_password'], creds['password'])
                with page.expect_navigation(timeout=15000):
                    page.click(SELECTORS['login_button'])
                page.wait_for_load_state('domcontentloaded')

                navigate_to_csv_page(page, csv_page_url)
                shops = get_shop_list(page)
                logger.info(f'  店舗数: {len(shops)}')

                for i, shop in enumerate(reversed(shops)):
                    if progress.stop_requested:
                        logger.info('  ⛔ 中断要求により処理を停止しました')
                        break
                    progress.update(f"[{site['id']}] {site['name']}", shop, i, len(shops))
                    logger.info(f'  [{i + 1}/{len(shops)}] {shop}')
                    ok = download_with_retry(page, shop, year, month, save_dir, csv_page_url)
                    (success_list if ok else failure_list).append(shop)

                progress.update(f"[{site['id']}] {site['name']}  完了", '', len(shops), len(shops))
            except Exception as e:
                logger.error(f'  サイト処理中にエラー: {e}')
            finally:
                browser.close()

            total = len(success_list) + len(failure_list)
            logger.info(f"  {site['name']}：{len(success_list)}/{total}")
            if failure_list:
                for s in failure_list:
                    logger.info(f'    失敗: {s}')

            all_success += len(success_list)
            all_failure += [f"[{site['id']}]{s}" for s in failure_list]

            if progress.stop_requested:
                logger.info('⛔ 中断のため残りのサイトをスキップします')
                break

    was_stopped = progress.stop_requested
    progress.close()

    logger.info(f'{"=" * 40}')
    logger.info(f'全サイト完了  成功合計: {all_success} 件 / 失敗合計: {len(all_failure)} 件')
    if all_failure:
        for s in all_failure:
            logger.info(f'  失敗: {s}')

    _root.deiconify()
    _root.attributes('-topmost', True)

    msg = '中断して終了しました' if was_stopped else '正常に終了しました'
    messagebox.showinfo('完了', msg, parent=_root)
    _root.destroy()

    os.startfile(log_path)


def navigate_to_csv_page(page: Page, csv_page_url: str):
    try:
        page.goto(csv_page_url)
        page.wait_for_load_state('domcontentloaded')
    except Exception as e:
        raise RuntimeError(f'CSVページへの移動に失敗: {e}')


def get_shop_list(page: Page) -> list[str]:
    options = page.locator(f"{SELECTORS['facility_select']} option").all()
    return [opt.inner_text().strip() for opt in options if opt.inner_text().strip()]


def download_with_retry(
    page: Page, shop_name: str, year: str, month: str, save_dir: Path, csv_page_url: str
) -> bool:
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            download_csv_for_shop(page, shop_name, year, month, save_dir, csv_page_url)
            return True
        except SkipShop as e:
            print(f'  スキップ: {e}')
            return True
        except Exception as e:
            print(f'  試行 {attempt}/{MAX_RETRIES} 失敗: {e}')
            if attempt < MAX_RETRIES:
                try:
                    navigate_to_csv_page(page, csv_page_url)
                except Exception:
                    pass
                time.sleep(2)

    print(f'  {shop_name} のダウンロードに失敗しました')
    return False


def download_csv_for_shop(
    page: Page, shop_name: str, year: str, month: str, save_dir: Path, csv_page_url: str
):
    navigate_to_csv_page(page, csv_page_url)

    page.select_option(SELECTORS['facility_select'], label=shop_name)
    page.select_option(SELECTORS['year_select'], value=year)

    month_no_zero = str(int(month))
    try:
        page.select_option(SELECTORS['month_select'], value=month_no_zero, timeout=5000)
    except PWTimeoutError:
        page.select_option(SELECTORS['month_select'], value=month, timeout=5000)

    page.click(SELECTORS['display_change_btn'])
    page.wait_for_load_state('domcontentloaded')

    try:
        page.wait_for_selector(SELECTORS['skip_update_dialog'], timeout=5000)
        page.click(SELECTORS['skip_update_dialog'])
        page.wait_for_load_state('domcontentloaded')
    except PWTimeoutError:
        pass

    print_btn = page.locator(SELECTORS['print_list_btn'])
    try:
        print_btn.wait_for(state='attached', timeout=5000)
    except PWTimeoutError:
        raise SkipShop('集計一覧の印刷ボタンが見つからない')

    if not print_btn.is_enabled():
        raise SkipShop('集計一覧の印刷がグレーアウト')

    page.click(SELECTORS['print_list_btn'])
    page.wait_for_load_state('domcontentloaded')

    check_if_unchecked(page, SELECTORS['check_shienshihi'])
    check_if_unchecked(page, SELECTORS['check_service_code'])

    page.click(SELECTORS['display_toggle_btn'])
    page.wait_for_load_state('domcontentloaded')

    with page.expect_download(timeout=30000) as dl_info:
        page.click(SELECTORS['csv_output_btn'])

    download = dl_info.value
    filename = download.suggested_filename or f'{year}{month}_{shop_name}.csv'
    save_path = save_dir / filename
    download.save_as(save_path)
    print(f'  保存: {save_path.name}')

    try:
        page.click(SELECTORS['cancel_btn'], timeout=3000)
    except PWTimeoutError:
        pass

    time.sleep(1)


def check_if_unchecked(page: Page, selector: str):
    cb = page.locator(selector)
    if not cb.is_checked():
        cb.click()


if __name__ == '__main__':
    main()

