# カウネットの実績結合と集計を一括で実行するためのGUIアプリ
import os
import re
import webbrowser
import tkinter as tk
from tkinter import messagebox, ttk
import カウネット実績結合
import カウネット集計 as カウネット集計

# ============================== 仕事の進め方 / 取説 ==============================

INSTRUCTIONS_FILENAME = "手順.md"

DEFAULT_INSTRUCTIONS = """\
※この手順は変更されることがあります。内容を直接書き換えたい場合は
　右上の「✎ 取説を編集」ボタンを押してください(メモ帳でこのファイルが開きます)。
※URL(https://...)やファイルのパス(例: C:\\folder\\file.xlsx)をそのまま1行に
　書いておくと、「仕事の進め方」画面でクリックできるリンクになります。
"""

_LINK_CHAR = r"[^\s。、，」』（）\"']"
LINK_PATTERN = re.compile(rf"(https?://{_LINK_CHAR}+|[A-Za-z]:\\{_LINK_CHAR}+|\\\\{_LINK_CHAR}+)")
LINK_TRAILING_CHARS = "。、,.)）」』\"'"


def get_instructions_path():
    app_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(app_dir, INSTRUCTIONS_FILENAME)


def ensure_instructions_file():
    path = get_instructions_path()
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            f.write(DEFAULT_INSTRUCTIONS)
    return path


def read_instructions():
    path = ensure_instructions_file()
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except OSError as e:
        return f"手順ファイルの読み込みに失敗しました。\n{path}\n{e}"


def open_instructions_editor():
    path = ensure_instructions_file()
    import subprocess
    try:
        subprocess.Popen(["notepad.exe", path])
    except OSError as e:
        messagebox.showerror("編集に失敗しました", f"{path}\n{e}")


def open_link(href):
    if href.startswith("http://") or href.startswith("https://"):
        webbrowser.open(href)
        return
    try:
        os.startfile(href)
    except OSError as e:
        messagebox.showerror("開けませんでした", f"{href}\n{e}")


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
        base_font = tkfont.nametofont(text_widget.cget("font"))
        family, size = base_font.actual("family"), base_font.actual("size")
    except Exception:
        family, size = "TkDefaultFont", 10

    text_widget.tag_configure("hyperlink", foreground="#2f6f63", underline=True)
    text_widget.tag_configure("md_h1", font=(family, size + 5, "bold"), spacing1=10, spacing3=6)
    text_widget.tag_configure("md_h2", font=(family, size + 3, "bold"), spacing1=8, spacing3=4)
    text_widget.tag_configure("md_h3", font=(family, size + 1, "bold"), spacing1=6, spacing3=3)
    text_widget.tag_configure("md_bold", font=(family, size, "bold"))
    text_widget.tag_configure("md_code", font=("Consolas", size), background="#eeeeee")
    text_widget.tag_configure("md_code_block", font=("Consolas", size), background="#f2f2f2", lmargin1=14, lmargin2=14)
    text_widget.tag_configure("md_quote", foreground="#4a4a4a", background="#f0efe6", lmargin1=20, lmargin2=20, spacing1=3, spacing3=3)
    text_widget.tag_configure("md_hr", foreground="#aaaaaa")

    orig_cursor = text_widget.cget("cursor")
    link_index = 0

    def insert_inline(text_line, block_tags):
        nonlocal link_index
        pos = 0
        for m in _MD_INLINE_PATTERN.finditer(text_line):
            start, end = m.span()
            if start > pos:
                text_widget.insert("end", text_line[pos:start], block_tags)
            if m.group("bold"):
                text_widget.insert("end", m.group("boldtext"), block_tags + ("md_bold",))
            elif m.group("code"):
                text_widget.insert("end", m.group("codetext"), block_tags + ("md_code",))
            else:
                raw = m.group("link")
                href = raw.rstrip(LINK_TRAILING_CHARS)
                trailing = raw[len(href):]
                tag_name = f"link_{link_index}"
                link_index += 1
                text_widget.insert("end", href, block_tags + ("hyperlink", tag_name))
                text_widget.tag_bind(tag_name, "<Button-1>", lambda e, href=href: open_link(href))
                text_widget.tag_bind(tag_name, "<Enter>", lambda e: text_widget.config(cursor="hand2"))
                text_widget.tag_bind(tag_name, "<Leave>", lambda e: text_widget.config(cursor=orig_cursor))
                if trailing:
                    text_widget.insert("end", trailing, block_tags)
            pos = end
        text_widget.insert("end", text_line[pos:], block_tags)

    in_code_block = False
    for raw_line in content.splitlines():
        stripped = raw_line.strip()

        if stripped.startswith("```"):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            text_widget.insert("end", raw_line + "\n", ("md_code_block",))
            continue

        if re.fullmatch(r"[-_*]{3,}", stripped):
            text_widget.insert("end", "─" * 42 + "\n", ("md_hr",))
            continue

        m = re.match(r"^(#{1,6})\s+(.*)$", raw_line)
        if m:
            level = len(m.group(1))
            tag = "md_h1" if level == 1 else "md_h2" if level == 2 else "md_h3"
            insert_inline(m.group(2), (tag,))
            text_widget.insert("end", "\n")
            continue

        m = re.match(r"^>\s?(.*)$", raw_line)
        if m:
            inner = m.group(1)
            hm = re.match(r"^(#{1,6})\s+(.*)$", inner)
            if hm:
                level = len(hm.group(1))
                tag = "md_h1" if level == 1 else "md_h2" if level == 2 else "md_h3"
                insert_inline(hm.group(2), ("md_quote", tag))
            else:
                insert_inline(inner, ("md_quote",))
            text_widget.insert("end", "\n")
            continue

        insert_inline(raw_line, ())
        text_widget.insert("end", "\n")


_instructions_win = None


def show_instructions_dialog(root, auto=False):
    global _instructions_win
    if _instructions_win is not None and _instructions_win.winfo_exists():
        _instructions_win.lift()
        _instructions_win.focus_force()
        return

    win = tk.Toplevel(root)
    _instructions_win = win
    win.title(f"仕事の進め方({os.path.basename(os.path.dirname(os.path.abspath(__file__)))})")
    win.geometry("560x520")
    win.transient(root)

    text_frame = ttk.Frame(win)
    text_frame.pack(fill="both", expand=True, padx=14, pady=(14, 6))

    text_widget = tk.Text(text_frame, wrap="word")
    scroll = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
    text_widget.config(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    text_widget.pack(side="left", fill="both", expand=True)

    insert_instructions_with_links(text_widget, read_instructions())
    text_widget.config(state="disabled")

    button_row = ttk.Frame(win)
    button_row.pack(side="bottom", fill="x", padx=14, pady=(0, 14))
    ttk.Button(button_row, text="✎ 取説を編集", command=open_instructions_editor).pack(side="left")
    ttk.Button(button_row, text="閉じる", command=win.destroy).pack(side="right")

    if not auto:
        win.lift()
        win.focus_force()


def run_kekko():
    try:
        カウネット実績結合.main()
    except Exception as e:
        messagebox.showerror("エラー", f"実績結合でエラーが発生しました:\n{e}")

def run_shukei():
    try:
        カウネット集計.main()
    except Exception as e:
        messagebox.showerror("エラー", f"集計処理でエラーが発生しました:\n{e}")

def main():
    root = tk.Tk()
    root.title("カウネット処理ツール一式")
    root.geometry("400x320")

    # --- 最前面に表示する設定 ---
    root.attributes("-topmost", True)
    # --------------------------

    header_row = tk.Frame(root)
    header_row.pack(side="top", fill="x", padx=10, pady=(8, 0))
    tk.Button(header_row, text="📋 仕事の進め方", command=lambda: show_instructions_dialog(root)).pack(side="right")
    tk.Button(header_row, text="✎ 取説を編集", command=open_instructions_editor).pack(side="right", padx=(0, 6))

    label = tk.Label(root, text="実行したいメニューを選択してください", font=("MS Gothic", 11))
    label.pack(pady=20)

    btn1 = tk.Button(root, text="1. 実績データのダウンロードと結合を実行", 
                     command=run_kekko, width=35, height=2, bg="#f0f0f0")
    btn1.pack(pady=5)

    btn2 = tk.Button(root, text="2. コピペ用データの集計を実行", 
                     command=run_shukei, width=35, height=2, bg="#f0f0f0")
    btn2.pack(pady=5)

    btn3 = tk.Button(root, text="3. アプリ終了", 
                     command=root.destroy, width=35, height=2, bg="#ffcccc")
    btn3.pack(pady=10)

    # メイン画面が表示された直後に、仕事の進め方ガイドを自動表示する
    root.after(200, lambda: show_instructions_dialog(root, auto=True))

    root.mainloop()

if __name__ == "__main__":
    main()



