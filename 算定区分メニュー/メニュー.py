import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox


def app_dir() -> Path:
    """このメニューexe（または.py）が置かれているフォルダ"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


# 呼び出す2つのexe（このメニューexeと同じフォルダに配置しておくこと）
EXES = [
    {'label': '① 算定区分CSVダウンロード', 'name': '算定区分CSVダウンロード.exe'},
    {'label': '② 算定区分後工程一括', 'name': '算定区分後工程一括.exe'},
]


def run_exe(name: str):
    path = app_dir() / name
    if not path.exists():
        messagebox.showerror(
            'エラー',
            f'ファイルが見つかりません:\n{path}\n\n'
            'このメニューexeと同じフォルダに配置してください。',
        )
        return
    try:
        subprocess.Popen([str(path)], cwd=str(path.parent))
    except Exception as e:
        messagebox.showerror('エラー', f'起動に失敗しました:\n{e}')


def main():
    root = tk.Tk()
    root.title('算定区分メニュー')
    root.resizable(False, False)

    font_title = ('Yu Gothic UI', 14, 'bold')
    font_btn = ('Yu Gothic UI', 13)

    frame = tk.Frame(root, padx=30, pady=24)
    frame.pack()

    tk.Label(frame, text='実行する処理を選択してください', font=font_title).pack(pady=(0, 16))

    for item in EXES:
        tk.Button(
            frame,
            text=item['label'],
            font=font_btn,
            width=32,
            height=2,
            command=lambda n=item['name']: run_exe(n),
        ).pack(pady=6)

    tk.Frame(frame, height=1, bg='#ccc').pack(fill='x', pady=(14, 10))

    tk.Button(frame, text='終了', font=font_btn, width=32, command=root.destroy).pack()

    root.mainloop()


if __name__ == '__main__':
    main()
