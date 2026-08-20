#メインループ
#このコードは、ユーザーに対して一連の処理を実行する前に確認を求め、複数のPythonスクリプトを順番に実行し、完了後にメッセージを表示するためのものです。escapeキーで終了することもできます。
import ctypes
import tkinter as tk
from tkinter import messagebox, ttk
import re
import sys
import os
import shutil
import webbrowser
from pathlib import Path


def _minimize_console_window():
    """python.exe実行時に自動で開く黒いコンソール画面を、起動直後に最小化する。"""
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 6)  # SW_MINIMIZE
    except Exception:
        pass


_minimize_console_window()

# 1. 各スクリプトをモジュールとしてインポート
import 算定区分01
import 算定区分02
import 算定区分03
import 算定区分04
import 算定区分05

# ============================== 仕事の進め方 / 取説 ==============================

INSTRUCTIONS_FILENAME = '手順.md'

DEFAULT_INSTRUCTIONS = """\
※この手順は変更されることがあります。内容を直接書き換えたい場合は
　右上の「✎ 取説を編集」ボタンを押してください(メモ帳でこのファイルが開きます)。
※URL(https://...)やファイルのパス(例: C:\\folder\\file.xlsx)をそのまま1行に
　書いておくと、「仕事の進め方」画面でクリックできるリンクになります。
"""

_LINK_CHAR = r"[^\s。、，」』（）\"']"
LINK_PATTERN = re.compile(rf"(https?://{_LINK_CHAR}+|[A-Za-z]:\\{_LINK_CHAR}+|\\\\{_LINK_CHAR}+)")
LINK_TRAILING_CHARS = "。、,.)）」』\"'"


def _instructions_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_instructions_path() -> Path:
    return _instructions_dir() / INSTRUCTIONS_FILENAME


def ensure_instructions_file() -> Path:
    path = get_instructions_path()
    if not path.exists():
        path.write_text(DEFAULT_INSTRUCTIONS, encoding='utf-8')
    return path


def read_instructions() -> str:
    path = ensure_instructions_file()
    try:
        return path.read_text(encoding='utf-8')
    except OSError as e:
        return f'手順ファイルの読み込みに失敗しました。\n{path}\n{e}'


def open_instructions_editor():
    path = ensure_instructions_file()
    import subprocess
    try:
        subprocess.Popen(['notepad.exe', path])
    except OSError as e:
        messagebox.showerror('編集に失敗しました', f'{path}\n{e}')


def open_link(href):
    if href.startswith('http://') or href.startswith('https://'):
        webbrowser.open(href)
        return
    try:
        os.startfile(href)
    except OSError as e:
        messagebox.showerror('開けませんでした', f'{href}\n{e}')


_MD_INLINE_PATTERN = re.compile(
    r"(?P<bold>\*\*(?P<boldtext>.+?)\*\*)"
    rf"|(?P<code>`(?P<codetext>[^`\n]+)`)"
    rf"|(?P<link>https?://{_LINK_CHAR}+|[A-Za-z]:\\{_LINK_CHAR}+|\\\\{_LINK_CHAR}+)"
)


def insert_instructions_with_links(text_widget, content):
    """Markdownの簡易記法(見出し・引用・コードブロック・太字・区切り線)をそれらしく
    整形しつつ、URL・ファイルパスをクリック可能なリンクとして挿入する。"""
    from tkinter import font as tkfont
    try:
        base_font = tkfont.nametofont(text_widget.cget('font'))
        family, size = base_font.actual('family'), base_font.actual('size')
    except Exception:
        family, size = 'TkDefaultFont', 10

    text_widget.tag_configure('hyperlink', foreground='#2f6f63', underline=True)
    text_widget.tag_configure('md_h1', font=(family, size + 5, 'bold'), spacing1=10, spacing3=6)
    text_widget.tag_configure('md_h2', font=(family, size + 3, 'bold'), spacing1=8, spacing3=4)
    text_widget.tag_configure('md_h3', font=(family, size + 1, 'bold'), spacing1=6, spacing3=3)
    text_widget.tag_configure('md_bold', font=(family, size, 'bold'))
    text_widget.tag_configure('md_code', font=('Consolas', size), background='#eeeeee')
    text_widget.tag_configure('md_code_block', font=('Consolas', size), background='#f2f2f2', lmargin1=14, lmargin2=14)
    text_widget.tag_configure('md_quote', foreground='#4a4a4a', background='#f0efe6', lmargin1=20, lmargin2=20, spacing1=3, spacing3=3)
    text_widget.tag_configure('md_hr', foreground='#aaaaaa')

    orig_cursor = text_widget.cget('cursor')
    link_index = 0

    def insert_inline(text_line, block_tags):
        nonlocal link_index
        pos = 0
        for m in _MD_INLINE_PATTERN.finditer(text_line):
            start, end = m.span()
            if start > pos:
                text_widget.insert('end', text_line[pos:start], block_tags)
            if m.group('bold'):
                text_widget.insert('end', m.group('boldtext'), block_tags + ('md_bold',))
            elif m.group('code'):
                text_widget.insert('end', m.group('codetext'), block_tags + ('md_code',))
            else:
                raw = m.group('link')
                href = raw.rstrip(LINK_TRAILING_CHARS)
                trailing = raw[len(href):]
                tag_name = f'link_{link_index}'
                link_index += 1
                text_widget.insert('end', href, block_tags + ('hyperlink', tag_name))
                text_widget.tag_bind(tag_name, '<Button-1>', lambda e, href=href: open_link(href))
                text_widget.tag_bind(tag_name, '<Enter>', lambda e: text_widget.config(cursor='hand2'))
                text_widget.tag_bind(tag_name, '<Leave>', lambda e: text_widget.config(cursor=orig_cursor))
                if trailing:
                    text_widget.insert('end', trailing, block_tags)
            pos = end
        text_widget.insert('end', text_line[pos:], block_tags)

    in_code_block = False
    for raw_line in content.splitlines():
        stripped = raw_line.strip()

        if stripped.startswith('```'):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            text_widget.insert('end', raw_line + '\n', ('md_code_block',))
            continue

        if re.fullmatch(r'[-_*]{3,}', stripped):
            text_widget.insert('end', '─' * 42 + '\n', ('md_hr',))
            continue

        m = re.match(r'^(#{1,6})\s+(.*)$', raw_line)
        if m:
            level = len(m.group(1))
            tag = 'md_h1' if level == 1 else 'md_h2' if level == 2 else 'md_h3'
            insert_inline(m.group(2), (tag,))
            text_widget.insert('end', '\n')
            continue

        m = re.match(r'^>\s?(.*)$', raw_line)
        if m:
            inner = m.group(1)
            hm = re.match(r'^(#{1,6})\s+(.*)$', inner)
            if hm:
                level = len(hm.group(1))
                tag = 'md_h1' if level == 1 else 'md_h2' if level == 2 else 'md_h3'
                insert_inline(hm.group(2), ('md_quote', tag))
            else:
                insert_inline(inner, ('md_quote',))
            text_widget.insert('end', '\n')
            continue

        insert_inline(raw_line, ())
        text_widget.insert('end', '\n')


def show_instructions_dialog():
    """起動直後に「仕事の進め方」を表示し、閉じるまで先へ進ませない。"""
    root = tk.Tk()
    root.withdraw()

    win = tk.Toplevel(root)
    win.title(f"仕事の進め方({_instructions_dir().name})")
    win.geometry('560x520')

    text_frame = ttk.Frame(win)
    text_frame.pack(fill='both', expand=True, padx=14, pady=(14, 6))

    text_widget = tk.Text(text_frame, wrap='word')
    scroll = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
    text_widget.config(yscrollcommand=scroll.set)
    scroll.pack(side='right', fill='y')
    text_widget.pack(side='left', fill='both', expand=True)

    insert_instructions_with_links(text_widget, read_instructions())
    text_widget.config(state='disabled')

    button_row = ttk.Frame(win)
    button_row.pack(side='bottom', fill='x', padx=14, pady=(0, 14))
    ttk.Button(button_row, text='✎ 取説を編集', command=open_instructions_editor).pack(side='left')
    ttk.Button(button_row, text='閉じる(作業を始める)', command=win.destroy).pack(side='right')

    win.protocol('WM_DELETE_WINDOW', win.destroy)
    win.attributes('-topmost', True)
    win.lift()
    win.focus_force()
    root.wait_window(win)
    root.destroy()


def exit_program(event=None):
    """Escキーが押された時の処理"""
    if messagebox.askyesno("中断", "プログラム全体を終了しますか？"):
        sys.exit()

def run():
    show_instructions_dialog()

    root = tk.Tk()
    root.withdraw()

    # Escキーを監視
    # root.bind('<Escape>', exit_program)
    
    #from tkinter import messagebox
    confirm_msg = (
        "【Excel版】算定区分・加算等集計表（運営用）は準備済ですか？\n\n"
        "「はい」を選択すると処理を開始します。\n"
        "「いいえ」を選択すると終了します。"
    )
    
    if not messagebox.askyesno("確認", confirm_msg):
        messagebox.showwarning("強制終了", "キャンセルが選択されたため、強制終了しました。")
        return

    root.cancelled = False

    try:
        # 2. 各ファイルの main() 関数を順番に実行
        # これにより PyInstaller は自動的にこれらのファイルを EXE に含めます
        # 途中のファイル選択ダイアログでキャンセルされた場合は root.cancelled が True になり、
        # そこで処理を打ち切って強制終了メッセージを表示する

        #print("--- 01開始 ---")
        算定区分01.main(root)
        # 栃木/栃木以外への振り分け(コピー)が終わったので、ダウンロード直後の
        # 番号フォルダ(01~06)を「{年月}生データ」フォルダへ退避する
        source_root = 算定区分01.find_source_root(Path.home() / "Downloads")
        算定区分01.archive_raw_folders(source_root)
        #print("--- 01終了 / 02を呼び出します ---") #
        算定区分02.main(root)
        #print("--- 02終了 / 03を呼び出します ---")
        算定区分03.main(root)
        if root.cancelled:
            messagebox.showwarning("強制終了", "キャンセルが選択されたため、強制終了しました。")
            return

        算定区分04.main(root)
        if root.cancelled:
            messagebox.showwarning("強制終了", "キャンセルが選択されたため、強制終了しました。")
            return

        算定区分05.main(root)
        if root.cancelled:
            messagebox.showwarning("強制終了", "キャンセルが選択されたため、強制終了しました。")
            return

        messagebox.showinfo("完了", "すべての処理が正常に終了しました。")

        # --- フォルダ削除機能の追加 ---
        downloads = Path.home() / "Downloads"
        # 削除対象のリスト（Downloads配下と、実行ファイルと同階層の両方をチェック）
        cleanup_targets = [
            downloads / "excel_results",
            downloads / "downroads_temp",
            Path("excel_results"),
            Path("downroads_temp")
        ]
        for target in cleanup_targets:
            if target.exists() and target.is_dir():
                try:
                    shutil.rmtree(target)
                    print(f"削除成功: {target}")
                except Exception as e:
                    print(f"削除失敗: {target} - {e}")

        # --- 全作業終了。最後に確認用ファイルを開いて終わる ---
        final_file = downloads / "3-spreadsheetへコピペ用データ.xlsx"
        try:
            os.startfile(final_file)
        except OSError as e:
            messagebox.showwarning("ファイルを開けませんでした", f"{final_file}\n{e}")

    except Exception:
            # エラーの詳細（どこが間違っているか）を画面に出す
            import traceback
            error_details = traceback.format_exc()
            print(error_details) # コンソールに表示
            messagebox.showerror("エラー詳細", error_details) # 画面に表示

            # except Exception as e:
            #     # 各スクリプト内で root.destroy() された際のエラーや予期せぬエラーをキャッチ
            #     messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
            
    #root.destroy()

if __name__ == "__main__":
    run()
    #pyinstaller 算定区分処理実行.py --onefile --noconsole
    #cd C:\Users\owner\Documents\GitHub\global_python