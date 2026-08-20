# -*- coding: utf-8 -*-
"""
店舗検索ミニウィンドウ
- 画面右下に常時表示・最前面に表示され続ける小さな検索ボックス
- キーワードを入力すると、エクセル「リスト」シートの店舗名にマッチする
  行を検索し、対応するエリア（区分）を一覧表示する
- Excelファイルは起動時に読み込み。ファイルが更新された場合は
  検索欄の右にある「更新」ボタン、または Ctrl+R で再読込できる
"""

import ctypes
from ctypes import wintypes
import tkinter as tk
from tkinter import font as tkfont
import openpyxl
import sys
import os

EXCEL_PATH = r"C:\Users\owner\Desktop\works\保管書類\店舗リスト\table_HHHL共通店舗一覧m.xlsm"
SHEET_NAME = "リスト"
NAME_COL = "店舗名"
AREA_COL = "区分"

WIN_WIDTH = 224   # 300px - 2cm(約76px, 96dpi換算)
WIN_HEIGHT = 246  # 260px - 1cm(約38px, 96dpi換算) + 最前面スイッチ分
MARGIN_RIGHT = 12
MARGIN_BOTTOM = 8  # タスクバーはワークエリア計算で自動的に避けるため、ここは隙間程度でよい

SPI_GETWORKAREA = 0x0030


class StoreData:
    def __init__(self, path, sheet_name, name_col, area_col):
        self.path = path
        self.sheet_name = sheet_name
        self.name_col = name_col
        self.area_col = area_col
        self.rows = []  # [(店舗名, エリア), ...]
        self.error = None
        self.reload()

    def reload(self):
        try:
            wb = openpyxl.load_workbook(
                self.path, read_only=True, data_only=True, keep_vba=True
            )
            ws = wb[self.sheet_name]
            header = None
            rows = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    header = row
                    try:
                        name_idx = header.index(self.name_col)
                        area_idx = header.index(self.area_col)
                    except ValueError:
                        self.error = "見出し行に「{}」または「{}」が見つかりません".format(
                            self.name_col, self.area_col
                        )
                        self.rows = []
                        wb.close()
                        return
                    continue
                name = row[name_idx] if name_idx < len(row) else None
                area = row[area_idx] if area_idx < len(row) else None
                if name:
                    rows.append((str(name), str(area) if area else ""))
            wb.close()
            self.rows = rows
            self.error = None
        except FileNotFoundError:
            self.error = "エクセルファイルが見つかりません:\n{}".format(self.path)
            self.rows = []
        except PermissionError:
            self.error = "エクセルファイルを開けません（他で使用中の可能性があります）"
            self.rows = []
        except Exception as e:
            self.error = "読み込みエラー: {}".format(e)
            self.rows = []

    def search(self, keyword):
        if not keyword:
            return []
        kw = keyword.casefold()
        return [
            (name, area)
            for name, area in self.rows
            if kw in name.casefold()
        ]


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("店舗検索")
        self.root.geometry("{}x{}".format(WIN_WIDTH, WIN_HEIGHT))
        self.root.resizable(False, False)
        self.topmost_var = tk.BooleanVar(value=True)
        self.root.attributes("-topmost", self.topmost_var.get())

        self._position_bottom_right()

        self.data = StoreData(EXCEL_PATH, SHEET_NAME, NAME_COL, AREA_COL)

        self._build_ui()
        self._refresh_status()

        self.root.bind("<Control-r>", lambda e: self._reload())
        self.root.bind("<Escape>", lambda e: self.entry.delete(0, tk.END))

    def _get_work_area(self):
        rect = wintypes.RECT()
        if ctypes.windll.user32.SystemParametersInfoW(
            SPI_GETWORKAREA, 0, ctypes.byref(rect), 0
        ):
            return rect.left, rect.top, rect.right, rect.bottom
        return 0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight()

    def _position_bottom_right(self):
        self.root.update_idletasks()
        work_left, work_top, work_right, work_bottom = self._get_work_area()

        x = work_right - WIN_WIDTH - MARGIN_RIGHT
        y = work_bottom - WIN_HEIGHT - MARGIN_BOTTOM
        self.root.geometry("+{}+{}".format(max(x, work_left), max(y, work_top)))
        self.root.update_idletasks()

        # タイトルバー分の実際の高さを測り、ウィンドウ全体がワークエリア内に
        # 収まるように補正する（タイトルバーの高さはOS/テーマ依存のため）
        try:
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            outer = wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(outer))
            overflow = outer.bottom - work_bottom
            if overflow > 0:
                y = max(y - overflow, work_top)
                self.root.geometry("+{}+{}".format(max(x, work_left), y))
        except Exception:
            pass

    def _build_ui(self):
        base_font = tkfont.nametofont("TkDefaultFont")
        base_font.configure(size=10)

        top_frame = tk.Frame(self.root)
        top_frame.pack(fill="x", padx=6, pady=(6, 2))

        tk.Label(top_frame, text="店舗名:").pack(side="left")

        self.entry = tk.Entry(top_frame, font=base_font)
        self.entry.pack(side="left", fill="x", expand=True, padx=(4, 4))
        self.entry.bind("<KeyRelease>", self._on_key)
        self.entry.focus_set()

        reload_btn = tk.Button(top_frame, text="更新", width=4, command=self._reload)
        reload_btn.pack(side="left")

        option_frame = tk.Frame(self.root)
        option_frame.pack(fill="x", padx=6, pady=(0, 2))

        topmost_check = tk.Checkbutton(
            option_frame,
            text="常に最前面",
            variable=self.topmost_var,
            command=self._toggle_topmost,
        )
        topmost_check.pack(side="left")

        self.status_var = tk.StringVar()
        status_label = tk.Label(
            self.root, textvariable=self.status_var, fg="#666666", anchor="w"
        )
        status_label.pack(fill="x", padx=8)

        list_frame = tk.Frame(self.root)
        list_frame.pack(fill="both", expand=True, padx=6, pady=(2, 6))

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        # Listboxはテキストのコピーができないため、選択・Ctrl+Cでの
        # コピーに対応したTextウィジェット（読み取り専用）を使用する。
        # 店舗名とエリアを別々の行にして、行単位で選択・コピーしやすくする。
        self.result_text = tk.Text(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=base_font,
            wrap="none",
            cursor="arrow",
        )
        self.result_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.result_text.yview)

        self.result_text.bind("<Key>", self._block_edit_keys)
        self.result_text.bind("<<Paste>>", lambda e: "break")
        self.result_text.bind("<<Cut>>", lambda e: "break")
        self.result_text.bind("<Button-2>", lambda e: "break")

        self.context_menu = tk.Menu(self.result_text, tearoff=0)
        self.context_menu.add_command(label="コピー", command=self._copy_selection)
        self.context_menu.add_command(label="すべて選択してコピー", command=self._copy_all)
        self.result_text.bind("<Button-3>", self._show_context_menu)

    def _show_context_menu(self, event):
        # 選択範囲がなければ、右クリックした行を自動選択してからコピーできるようにする
        if not self.result_text.tag_ranges("sel"):
            line_start = "{} linestart".format(self.result_text.index("@{},{}".format(event.x, event.y)))
            line_end = "{} lineend".format(line_start)
            self.result_text.tag_add("sel", line_start, line_end)
        self.context_menu.tk_popup(event.x_root, event.y_root)
        return "break"

    def _copy_selection(self):
        if self.result_text.tag_ranges("sel"):
            self.result_text.event_generate("<<Copy>>")

    def _copy_all(self):
        self.result_text.tag_add("sel", "1.0", "end")
        self.result_text.event_generate("<<Copy>>")

    def _toggle_topmost(self):
        self.root.attributes("-topmost", self.topmost_var.get())

    def _refresh_status(self):
        if self.data.error:
            self.status_var.set(self.data.error)
        else:
            self.status_var.set("{} 件の店舗データを読み込み済み".format(len(self.data.rows)))

    def _reload(self):
        self.data.reload()
        self._refresh_status()
        self._on_key(None)

    def _block_edit_keys(self, event):
        # ナビゲーション・選択・コピー（Ctrl+C）以外のキー入力は無効化し、
        # 結果表示欄を実質的に読み取り専用にする。
        allowed_keysyms = {
            "Left", "Right", "Up", "Down", "Home", "End",
            "Prior", "Next", "Shift_L", "Shift_R", "Control_L", "Control_R",
        }
        ctrl_pressed = (event.state & 0x4) != 0
        if ctrl_pressed and event.keysym.lower() in ("c", "a"):
            return None
        if event.keysym in allowed_keysyms:
            return None
        return "break"

    def _on_key(self, event):
        keyword = self.entry.get().strip()
        self.result_text.delete("1.0", tk.END)

        if self.data.error:
            return

        if not keyword:
            return

        results = self.data.search(keyword)
        if not results:
            self.result_text.insert(tk.END, "該当なし")
            return

        lines = [
            "{}  →  {}".format(name, area) for name, area in results[:200]
        ]
        self.result_text.insert(tk.END, "\n".join(lines))

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
