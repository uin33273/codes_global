r"""
Auto-crop an image's whitespace margins to the limit and save the result.

- Dropping a single *.jpg file saves the cropped copy into the same
  folder as the source file, named "<元のファイル名>_トリミング済.<ext>".
- Dropping a folder processes every .jpg/.jpeg file inside it
  (non-recursive) and saves the cropped copies into a folder created
  next to it (i.e. in its parent folder), named "<フォルダ名>_トリミング済"
  (if that name already exists, "(1)", "(2)", ... is appended).
  e.g. dropping "...\6.3(5_9〆)\image" saves into
       "...\6.3(5_9〆)\image_トリミング済".

Two ways to use:

1) Drag & drop mode (no arguments): double-click this script (or run
   `python イメージトリミング.py`), leave the console window open, then
   drag a *.jpg file (or a folder containing *.jpg files) from Explorer
   onto the console window and press Enter. It crops + saves, then waits
   for the next file/folder. Type nothing and press Enter (or type exit)
   to quit.

2) One-shot mode: python イメージトリミング.py "path\to\image_or_folder" [threshold]

threshold (0-255, default 245): pixels lighter than this (per channel,
near-white) are treated as margin/background to be cropped away.
"""
import sys
import os
import re
import webbrowser
import tkinter as tk
from tkinter import messagebox, ttk
import numpy as np
from PIL import Image

DEFAULT_THRESHOLD = 245

# ============================== 仕事の進め方 / 取説 ==============================

INSTRUCTIONS_FILENAME = "手順_イメージトリミング.md"

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


def show_instructions_dialog():
    """起動直後に「仕事の進め方」を表示し、閉じるまで先へ進ませない。"""
    root = tk.Tk()
    root.withdraw()

    win = tk.Toplevel(root)
    win.title(f"仕事の進め方({os.path.basename(os.path.dirname(os.path.abspath(__file__)))})")
    win.geometry("560x520")

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
    ttk.Button(button_row, text="閉じる(作業を始める)", command=win.destroy).pack(side="right")

    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.attributes("-topmost", True)
    win.lift()
    win.focus_force()
    root.wait_window(win)
    root.destroy()


def autocrop(im, threshold=DEFAULT_THRESHOLD):
    rgb = im.convert("RGB")
    arr = np.array(rgb)
    # a pixel counts as "content" if any channel is darker than threshold
    mask = np.any(arr < threshold, axis=2)
    rows = np.where(mask.any(axis=1))[0]
    cols = np.where(mask.any(axis=0))[0]
    if rows.size == 0 or cols.size == 0:
        return im  # nothing but background found; return unchanged
    top, bottom = rows[0], rows[-1]
    left, right = cols[0], cols[-1]
    return im.crop((left, top, right + 1, bottom + 1))


def split_paths(line):
    """Split a line of one or more (optionally quoted) Windows paths
    without mangling backslashes the way shlex.split() would."""
    line = line.strip()
    # PowerShell prefixes a dragged file/folder with the call operator
    # "&" and single-quotes the path (e.g. & 'C:\...\file.jpg'); strip
    # that operator so the quoted path below is parsed normally
    if line.startswith("&"):
        line = line[1:].strip()
    return [
        dq or sq or bare
        for dq, sq, bare in re.findall(r'"([^"]+)"|\'([^\']+)\'|(\S+)', line)
    ]


def process_one(path, threshold=DEFAULT_THRESHOLD, out_dir=None):
    path = path.strip().strip('"').strip("'")
    if not path:
        return
    if not os.path.isfile(path):
        print(f"  !! ファイルが見つかりません: {path}")
        return
    if not path.lower().endswith((".jpg", ".jpeg")):
        print(f"  !! .jpg / .jpeg ではありません: {path}")
        return

    im = Image.open(path)
    print(f"  元サイズ: {im.size}")

    cropped = autocrop(im, threshold)
    print(f"  トリミング後: {cropped.size}")

    if out_dir is None:
        # top-level single-file drop: save next to the source file, suffixed
        out_dir = os.path.dirname(os.path.abspath(path))
        base, ext = os.path.splitext(os.path.basename(path))
        out_name = f"{base}_トリミング済{ext}"
    else:
        # called while processing a folder: folder itself is already suffixed
        out_name = os.path.basename(path)

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, out_name)
    cropped.convert("RGB").save(out_path, quality=95)
    print(f"  保存先: {out_path}")


def unique_dir(base_dir):
    """base_dir if it doesn't exist yet, otherwise base_dir(1), base_dir(2), ..."""
    if not os.path.exists(base_dir):
        return base_dir
    n = 1
    while True:
        candidate = f"{base_dir}({n})"
        if not os.path.exists(candidate):
            return candidate
        n += 1


def process_folder(folder, threshold=DEFAULT_THRESHOLD):
    names = sorted(
        name for name in os.listdir(folder)
        if name.lower().endswith((".jpg", ".jpeg"))
    )
    if not names:
        print(f"  !! フォルダ内に.jpg/.jpegが見つかりません: {folder}")
        return None
    folder = os.path.normpath(folder)
    parent_dir = os.path.dirname(folder)
    out_dir = unique_dir(os.path.join(parent_dir, f"{os.path.basename(folder)}_トリミング済"))
    print(f"  フォルダ内の画像 {len(names)} 件を処理します: {folder}")
    print(f"  保存先フォルダ: {out_dir}")
    for name in names:
        print(f"[{name}]")
        process_one(os.path.join(folder, name), threshold, out_dir=out_dir)
    return out_dir


def process_path(path, threshold=DEFAULT_THRESHOLD):
    path = path.strip().strip('"').strip("'")
    if not path:
        return None
    if os.path.isdir(path):
        return process_folder(path, threshold)
    elif os.path.isfile(path):
        process_one(path, threshold)
        return None
    else:
        print(f"  !! ファイル/フォルダが見つかりません: {path}")
        return None


def interactive_loop():
    print("=== JPEGトリミング (ドラッグ&ドロップ待機モード) ===")
    print("保存先: ファイルの場合は同じフォルダへ「ファイル名_トリミング済」として保存")
    print("        フォルダの場合はその親フォルダに「フォルダ名_トリミング済」フォルダを作成して保存(重複時は連番)")
    print(".jpg ファイル、またはそれらを含むフォルダをこのウィンドウにドラッグ&ドロップして Enter を押してください。")
    print("終了するには何も入力せず Enter、または exit と入力してください。")
    print()
    while True:
        try:
            line = input("> ").strip()
        except EOFError:
            break
        if not line or line.lower() in ("exit", "quit"):
            print("終了します。")
            break
        # Explorer can drop multiple files/folders on one line, each quoted separately
        folder_outputs = [d for d in (process_path(path) for path in split_paths(line)) if d]
        print("作業が終了しました。")
        if folder_outputs:
            for out_dir in folder_outputs:
                os.startfile(out_dir)
            return
        print()


def main():
    show_instructions_dialog()
    if len(sys.argv) > 1:
        threshold = int(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_THRESHOLD
        process_path(sys.argv[1], threshold)
        print()
        print("作業が終了しました。")
    else:
        interactive_loop()


if __name__ == "__main__":
    main()
