import os
import re
import shutil
import tkinter as tk
import unicodedata
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import openpyxl

STORE_LIST_XLSM = r"C:\Users\owner\Desktop\works\保管書類\店舗リスト\table_HHHL共通店舗一覧m.xlsm"
LIST_SHEET_NAME = "リスト"

# リネーム対象とする拡張子 (通常のPDFとzipの両方に対応)
TARGET_EXTENSIONS = {".pdf", ".zip"}

# ============================== 仕事の進め方 / 取説 ==============================

INSTRUCTIONS_FILENAME = "手順_欠席加算rename.txt"

DEFAULT_INSTRUCTIONS = """\
※この手順は変更されることがあります。内容を直接書き換えたい場合は
　右上の「✎ 取説を編集」ボタンを押してください(メモ帳でこのファイルが開きます)。
※URL(https://...)やファイルのパス(例: C:\\folder\\file.xlsx)をそのまま1行に
　書いておくと、「仕事の進め方」画面でクリックできるリンクになります。
"""

_LINK_CHAR = r"[^\s。、，」』（）\"']"
LINK_PATTERN = re.compile(rf"(https?://{_LINK_CHAR}+|[A-Za-z]:\\{_LINK_CHAR}+|\\\\{_LINK_CHAR}+)")
LINK_TRAILING_CHARS = "。、,.)）」』\"'"


def get_instructions_path() -> Path:
    return Path(__file__).resolve().parent / INSTRUCTIONS_FILENAME


def ensure_instructions_file() -> Path:
    path = get_instructions_path()
    if not path.exists():
        path.write_text(DEFAULT_INSTRUCTIONS, encoding="utf-8")
    return path


def read_instructions() -> str:
    path = ensure_instructions_file()
    try:
        return path.read_text(encoding="utf-8")
    except OSError as e:
        return f"手順ファイルの読み込みに失敗しました。\n{path}\n{e}"


def open_instructions_editor():
    path = ensure_instructions_file()
    try:
        os.startfile(path)
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


def insert_instructions_with_links(text_widget, content):
    text_widget.tag_configure("hyperlink", foreground="#2f6f63", underline=True)
    orig_cursor = text_widget.cget("cursor")
    link_index = 0
    for line in content.splitlines(keepends=True):
        pos = 0
        for m in LINK_PATTERN.finditer(line):
            start, end = m.span()
            if start > pos:
                text_widget.insert("end", line[pos:start])
            raw = m.group(1)
            href = raw.rstrip(LINK_TRAILING_CHARS)
            trailing = raw[len(href):]
            tag_name = f"link_{link_index}"
            link_index += 1
            text_widget.insert("end", href, ("hyperlink", tag_name))
            text_widget.tag_bind(tag_name, "<Button-1>", lambda e, href=href: open_link(href))
            text_widget.tag_bind(tag_name, "<Enter>", lambda e: text_widget.config(cursor="hand2"))
            text_widget.tag_bind(tag_name, "<Leave>", lambda e: text_widget.config(cursor=orig_cursor))
            if trailing:
                text_widget.insert("end", trailing)
            pos = end
        text_widget.insert("end", line[pos:])


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


def select_target_dir() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askdirectory(title="リネームする日報PDF(欠席時対応記録票)のフォルダを選択してください")
    root.destroy()
    if not selected:
        return None
    return Path(selected)

# 店舗名側の店舗コードパターン。
# KM/KP: "151KM今市" のように 数字+K+M(またはP)
# SF   : "001SFサカフル宇都宮" のように 数字+SF
# TH   : "001THザスパキッズ" のように 数字+TH
STORE_CODE_RE = re.compile(r"^(\d+)(K[MP]|SF|TH)")

# 店舗コード種別 -> ファイル名側のキー生成方法
# KM/KP はファイル名側で "数字が先" (例: 151M)、SF/TH は "英字が先" (例: S001, T001)
def store_code_to_file_key(digits: str, code: str) -> str:
    if code == "KM":
        return f"{digits}M"
    if code == "KP":
        return f"{digits}P"
    if code == "SF":
        return f"S{digits}"
    if code == "TH":
        return f"T{digits}"
    raise ValueError(code)


# ファイル名側: "数字+M/P" (例: 151M) または "S/T+数字" (例: S001, T001) を拾う
FILE_CODE_RE_DIGIT_FIRST = re.compile(r"^(\d+)([MP])")
FILE_CODE_RE_LETTER_FIRST = re.compile(r"^([ST])(\d+)")


def extract_file_key(name_norm: str) -> str | None:
    m = FILE_CODE_RE_LETTER_FIRST.match(name_norm)
    if m:
        return f"{m.group(1)}{m.group(2)}"
    m = FILE_CODE_RE_DIGIT_FIRST.match(name_norm)
    if m:
        return f"{m.group(1)}{m.group(2)}"
    return None


def normalize(text: str) -> str:
    if text is None:
        return ""
    return unicodedata.normalize("NFKC", text)


def load_code_area_map(xlsm_path: str, sheet_name: str) -> dict[str, str]:
    """"151M" のようなキー -> 区分(エリア) の辞書を作る"""
    wb = openpyxl.load_workbook(xlsm_path, data_only=True, keep_vba=False)
    ws = wb[sheet_name]

    code_map: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        store_name, kubun = row[0], row[1]
        if not store_name or not kubun:
            continue
        store_name = normalize(str(store_name).strip())
        kubun = str(kubun).strip()

        m = STORE_CODE_RE.match(store_name)
        if not m:
            continue
        digits, code = m.group(1), m.group(2)
        key = store_code_to_file_key(digits, code)
        code_map[key] = kubun

    return code_map


def build_new_name(original_name: str, area: str) -> str:
    prefix = f"{area}_"
    if original_name.startswith(prefix):
        return original_name
    return prefix + original_name


def rename_files_in_dir(
    target_dir: Path, output_dir: Path, code_map: dict[str, str], dry_run: bool = True
) -> None:
    for path in sorted(target_dir.iterdir()):
        if not path.is_file():
            continue
        if path.suffix.lower() not in TARGET_EXTENSIONS:
            continue

        name_norm = normalize(path.name)
        key = extract_file_key(name_norm)
        area = code_map.get(key) if key is not None else None

        if key is None:
            print(f"[未マッチ] コードが見つかりません: {path.name}")
            new_name = path.name
        elif area is None:
            print(f"[未マッチ] エクセルに該当コードなし ({key}): {path.name}")
            new_name = path.name
        else:
            new_name = build_new_name(path.name, area)
            print(f"{path.name}")
            print(f"  コード: {key} -> 区分(エリア): {area}")
            print(f"  新しい名前: {new_name}")

        new_path = output_dir / new_name
        print(f"  移動先: {new_path}")

        if dry_run:
            continue

        if new_path.exists():
            print(f"  [エラー] 移動先が既に存在します: {new_path}")
            continue

        shutil.move(str(path), str(new_path))
        print("  -> 移動完了")


def main():
    show_instructions_dialog()

    target_dir = select_target_dir()
    if target_dir is None:
        print("フォルダが選択されませんでした。処理を中止します。")
        return
    if not target_dir.exists():
        print(f"[エラー] フォルダが見つかりません: {target_dir}")
        return

    output_dir = Path.home() / "Desktop" / f"{target_dir.name}_追記済"
    print(f"出力先フォルダ: {output_dir}")

    code_map = load_code_area_map(STORE_LIST_XLSM, LIST_SHEET_NAME)

    print("\n--- 確認のため、まず移動予定の内容を表示します ---")
    rename_files_in_dir(target_dir, output_dir, code_map, dry_run=True)

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    apply = messagebox.askyesno(
        "確認",
        f"上記の内容でファイルを移動しますか？\n\n移動先: {output_dir}",
    )
    root.destroy()

    if not apply:
        print("\n(処理を中止しました。ファイルは移動されていません。)")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    print("\n--- 移動を実行します ---")
    rename_files_in_dir(target_dir, output_dir, code_map, dry_run=False)


if __name__ == "__main__":
    main()
