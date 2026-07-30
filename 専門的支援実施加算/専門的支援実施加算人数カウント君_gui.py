# -*- coding: utf-8 -*-
"""
専門的支援実施加算 人数カウント君 (GUI版)

やること:
  1. カレンダー画面からコピーしたテキストを、下のテキスト欄に貼り付け(Ctrl+V)
  2. 「集計する」ボタンを押すと、日ごとに「専門的支援実施加算」に記載されている
     人数をウィンドウ内に一覧表示する。結果はCSVにも保存できる。

■事前準備
  追加のライブラリは不要(標準のtkinterのみ使用)。

■起動方法
    同じフォルダの「専門的支援実施加算人数カウント君.lnk」をダブルクリック
    (pythonw.exe経由で起動するため、黒いコンソール画面は出ません)
    ※ python 専門的支援実施加算人数カウント君_gui.py で直接実行することも可能です。
"""

import csv
import os
import re
import unicodedata
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# URL、またはWindowsのファイルパス(ドライブレター指定 / \\サーバー共有)にマッチさせ、
# 「仕事の進め方」欄でクリック可能なリンクとして表示するための正規表現。
# 日本語の句読点・括弧はリンクの地の文で直後に続くことが多いため、リンク本体には含めない。
_LINK_CHAR = r"[^\s。、，」』（）\"']"
LINK_PATTERN = re.compile(
    rf"(https?://{_LINK_CHAR}+|[A-Za-z]:\\{_LINK_CHAR}+|\\\\{_LINK_CHAR}+)"
)
LINK_TRAILING_CHARS = "。、,.)）」』\"'"


TARGET_LABEL = "専門的支援実施加算"

INSTRUCTIONS_FILENAME = "手順.txt"

DEFAULT_INSTRUCTIONS = """\
①ブラウザでエリアを選択してHUGを開く
②HUGメニューの「請求管理」→「加算項目管理」を開く
③HUGで対象の児童を選択
④HUGで施設から対象店舗を選択
⑤「HUG表示切替」を押す
⑥更新された画面上でCtrl+Aを押し、画面上の全テキストを選択
⑦Ctrl+Cで選択テキストをコピー
⑧この「専門的支援実施加算人数カウント君」の貼り付け欄にCtrl+Vで貼り付け
⑨「集計する」ボタンを押す
⑩集計された日にちと人数を、RPA業務のkidsyyyymm.xlsxの中身と比較する
⑪齟齬があれば、スプレッドシート「yyyy年m月_欠席・送迎・上限・専門的実施チェックシート」へ指摘内容を入力する

※この手順は変更されることがあります。内容を直接書き換えたい場合は
　右上の「✎ 取説を編集」ボタンを押してください(メモ帳でこのファイルが開きます)。
※URL(https://...)やファイルのパス(例: C:\\folder\\file.xlsx)をそのまま1行に
　書いておくと、「仕事の進め方」画面でクリックできるリンクになります。
"""


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


def is_day_header(line: str):
    """カレンダーの日付見出し行かどうかを判定して日にちを返す。
    (例:'1'→1、祝日名が同じ行に付く'海の日 20'→20)"""
    m = re.fullmatch(r"(?:\S+\s+)?(\d{1,2})", line.strip())
    if m:
        n = int(m.group(1))
        if 1 <= n <= 31:
            return n
    return None


def is_weekday_header_row(line: str) -> bool:
    """「日 月 火 水 木 金 土」の曜日見出し行かどうか(区切り文字の違いに影響されないよう単語単位で比較)"""
    parts = line.split()
    return parts == ["日", "月", "火", "水", "木", "金", "土"]


def parse_target_counts(text: str):
    """
    カレンダー画面のコピペテキストから、日ごとの「専門的支援実施加算」の人数と氏名一覧を抽出する。
    戻り値: {day(int): [name, ...], ...}  (該当行が無い日も空リストで含まれる)

    見出し行(例:' 専門的支援実施加算' '延長支援加算' '個別サポート加算（Ⅰ）①')は
    いずれも「加算」という文字を含むため、それを区切りとして
    「専門的支援実施加算」の範囲内にある行だけを氏名として数える。
    """
    results = {}
    order = []
    current_day = None
    in_target = False

    for raw_line in text.splitlines():
        s = unicodedata.normalize("NFKC", raw_line).strip()
        if not s:
            continue
        if is_weekday_header_row(s):
            continue

        dnum = is_day_header(s)
        if dnum is not None:
            current_day = dnum
            in_target = False
            if current_day not in results:
                results[current_day] = []
                order.append(current_day)
            continue

        if current_day is None:
            continue

        if TARGET_LABEL in s:
            in_target = True
            continue

        if "加算" in s:
            in_target = False
            continue

        if in_target:
            results[current_day].append(s)

    return results, order


# ============================== GUI ==============================


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("専門的支援実施加算 人数カウント君")
        self.geometry("820x720")

        self.day_counts = []  # 保存用: [(day, count), ...]
        self._instructions_win = None

        self._build_widgets()

        # メイン画面が表示された直後に、仕事の進め方ガイドを自動表示する
        self.after(200, lambda: self._show_instructions_dialog(auto=True))

    def _build_widgets(self):
        pad = {"padx": 10, "pady": 6}

        header_row = ttk.Frame(self)
        header_row.pack(side="top", fill="x", padx=10, pady=(8, 0))
        ttk.Button(header_row, text="📋 仕事の進め方", command=self._show_instructions_dialog).pack(side="right")
        ttk.Button(header_row, text="✎ 取説を編集", command=self._open_instructions_editor).pack(side="right", padx=(0, 6))

        main_paned = tk.PanedWindow(self, orient="vertical", sashrelief="raised", sashwidth=6)
        main_paned.pack(fill="both", expand=True)

        top_frame = ttk.Frame(main_paned)
        main_paned.add(top_frame, height=320, minsize=150, stretch="always")

        bottom_frame = ttk.Frame(main_paned)
        main_paned.add(bottom_frame, minsize=200, stretch="always")

        # --- カレンダーテキストエリア ---
        frame_cal = ttk.LabelFrame(top_frame, text="① カレンダー画面のテキストを貼り付け(全選択コピペでOK)")
        frame_cal.pack(fill="both", expand=True, **pad)

        self.cal_text = tk.Text(frame_cal, height=13, wrap="none")
        self.cal_text.pack(fill="both", expand=True, padx=10, pady=10)
        self._add_edit_context_menu(self.cal_text)

        button_row = ttk.Frame(top_frame)
        button_row.pack(side="bottom", pady=6)
        ttk.Button(button_row, text="② 集計する", command=self._run_count).pack(side="left", padx=(0, 6))
        ttk.Button(button_row, text="全クリア", command=self._clear_all).pack(side="left")

        # --- 結果表示エリア ---
        frame_result = ttk.LabelFrame(bottom_frame, text="結果")
        frame_result.pack(fill="both", expand=True, **pad)

        result_button_row = ttk.Frame(frame_result)
        result_button_row.pack(side="bottom", pady=(0, 10))
        ttk.Button(result_button_row, text="日別人数をコピー(Excel貼り付け用)", command=self._copy_day_counts).pack(side="left", padx=(0, 6))
        ttk.Button(result_button_row, text="日別人数をCSVで保存...", command=self._save_csv).pack(side="left")

        paned = tk.PanedWindow(frame_result, orient="horizontal", sashrelief="raised", sashwidth=6)
        paned.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        left_frame = ttk.Frame(paned)
        ttk.Label(left_frame, text="日別人数(専門的支援実施加算)").pack(anchor="w")
        left_text_frame = ttk.Frame(left_frame)
        left_text_frame.pack(fill="both", expand=True)
        self.summary_text = tk.Text(left_text_frame, height=14, wrap="word", state="disabled")
        left_scroll = ttk.Scrollbar(left_text_frame, orient="vertical", command=self.summary_text.yview)
        self.summary_text.config(yscrollcommand=left_scroll.set)
        left_scroll.pack(side="right", fill="y")
        self.summary_text.pack(side="left", fill="both", expand=True)
        self._add_copy_context_menu(self.summary_text)
        paned.add(left_frame, stretch="always")

        right_frame = ttk.Frame(paned)
        ttk.Label(right_frame, text="内訳(該当者氏名・確認用)").pack(anchor="w")
        right_text_frame = ttk.Frame(right_frame)
        right_text_frame.pack(fill="both", expand=True)
        self.detail_text = tk.Text(right_text_frame, height=14, wrap="word", state="disabled")
        right_scroll = ttk.Scrollbar(right_text_frame, orient="vertical", command=self.detail_text.yview)
        self.detail_text.config(yscrollcommand=right_scroll.set)
        right_scroll.pack(side="right", fill="y")
        self.detail_text.pack(side="left", fill="both", expand=True)
        self._add_copy_context_menu(self.detail_text)
        paned.add(right_frame, stretch="always")

    # ---------- 右クリックメニュー ----------

    def _add_copy_context_menu(self, text_widget):
        menu = tk.Menu(text_widget, tearoff=0)
        menu.add_command(label="コピー", command=lambda: self._copy_selection(text_widget))
        menu.add_command(label="すべて選択", command=lambda: self._select_all_text(text_widget))

        def show_menu(event):
            menu.tk_popup(event.x_root, event.y_root)

        text_widget.bind("<Button-3>", show_menu)

    def _copy_selection(self, text_widget):
        try:
            selected = text_widget.get("sel.first", "sel.last")
        except tk.TclError:
            return
        # Excelは行区切りに\r\nを必要とするため、\nのみだとセル内改行として扱われてしまう
        selected = selected.replace("\r\n", "\n").replace("\n", "\r\n")
        self.clipboard_clear()
        self.clipboard_append(selected)

    def _select_all_text(self, text_widget):
        text_widget.tag_add("sel", "1.0", "end")

    def _add_edit_context_menu(self, text_widget):
        menu = tk.Menu(text_widget, tearoff=0)
        menu.add_command(label="切り取り", command=lambda: text_widget.event_generate("<<Cut>>"))
        menu.add_command(label="コピー", command=lambda: text_widget.event_generate("<<Copy>>"))
        menu.add_command(label="貼り付け", command=lambda: text_widget.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="すべて選択", command=lambda: self._select_all_text(text_widget))

        def show_menu(event):
            menu.tk_popup(event.x_root, event.y_root)

        text_widget.bind("<Button-3>", show_menu)

    # ---------- 仕事の進め方ガイド ----------

    def _show_instructions_dialog(self, auto=False):
        if self._instructions_win is not None and self._instructions_win.winfo_exists():
            self._instructions_win.lift()
            self._instructions_win.focus_force()
            return

        win = tk.Toplevel(self)
        self._instructions_win = win
        win.title(f"仕事の進め方({os.path.basename(os.path.dirname(os.path.abspath(__file__)))})")
        win.geometry("560x520")
        win.transient(self)

        text_frame = ttk.Frame(win)
        text_frame.pack(fill="both", expand=True, padx=14, pady=(14, 6))

        text_widget = tk.Text(text_frame, wrap="word", state="normal")
        scroll = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.config(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text_widget.pack(side="left", fill="both", expand=True)

        self._insert_with_links(text_widget, read_instructions())
        text_widget.config(state="disabled")
        self._add_copy_context_menu(text_widget)

        button_row = ttk.Frame(win)
        button_row.pack(side="bottom", fill="x", padx=14, pady=(0, 14))
        ttk.Button(button_row, text="✎ 取説を編集", command=self._open_instructions_editor).pack(side="left")
        ttk.Button(button_row, text="閉じる", command=win.destroy).pack(side="right")

        if not auto:
            win.lift()
            win.focus_force()

    def _open_instructions_editor(self):
        path = ensure_instructions_file()
        try:
            os.startfile(path)
        except OSError as e:
            messagebox.showerror("編集に失敗しました", f"{path}\n{e}")

    def _insert_with_links(self, text_widget, content):
        """本文中のURL・ファイルパスをクリック可能なリンクとして挿入する。"""
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
                text_widget.tag_bind(tag_name, "<Button-1>", lambda e, href=href: self._open_link(href))
                text_widget.tag_bind(tag_name, "<Enter>", lambda e: text_widget.config(cursor="hand2"))
                text_widget.tag_bind(tag_name, "<Leave>", lambda e: text_widget.config(cursor=orig_cursor))

                if trailing:
                    text_widget.insert("end", trailing)
                pos = end
            text_widget.insert("end", line[pos:])

    def _open_link(self, href):
        if href.startswith("http://") or href.startswith("https://"):
            webbrowser.open(href)
            return
        try:
            os.startfile(href)
        except OSError as e:
            messagebox.showerror("開けませんでした", f"{href}\n{e}")

    # ---------- イベント ----------

    def _run_count(self):
        cal_text = self.cal_text.get("1.0", "end").strip()
        if not cal_text:
            messagebox.showwarning("確認", "カレンダーのテキストが貼り付けられていません。")
            return

        results, order = parse_target_counts(cal_text)
        if not results:
            messagebox.showwarning("確認", "貼り付けたテキストから日付が見つかりませんでした。コピー範囲・書式をご確認ください。")
            return

        # 1日〜31日を必ず連番で出力する(見出し検出に失敗した日があっても
        # Excel貼り付け時に行がズレないよう、抜けている日は0人として埋める)
        days_sorted = list(range(1, 32))
        self.day_counts = [(day, len(results.get(day, []))) for day in days_sorted]

        self._show_result(days_sorted, results)

    def _clear_all(self):
        self.cal_text.delete("1.0", "end")
        self.day_counts = []

        self.summary_text.config(state="normal")
        self.summary_text.delete("1.0", "end")
        self.summary_text.config(state="disabled")

        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.config(state="disabled")

    def _show_result(self, days_sorted, results):
        total = sum(len(results.get(day, [])) for day in days_sorted)

        self.summary_text.config(state="normal")
        self.summary_text.delete("1.0", "end")
        lines = [f"■ 合計: {total}人", ""]
        for day in days_sorted:
            lines.append(f"{day}日\t{len(results.get(day, []))}\t人")
        self.summary_text.insert("1.0", "\n".join(lines))
        self.summary_text.config(state="disabled")

        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", "end")
        detail_lines = []
        for day in days_sorted:
            names = results.get(day, [])
            detail_lines.append(f"{day}日({len(names)}人)")
            for name in names:
                detail_lines.append(f"　{name}")
            detail_lines.append("")
        self.detail_text.insert("1.0", "\n".join(detail_lines))
        self.detail_text.config(state="disabled")

    def _copy_day_counts(self):
        if not self.day_counts:
            messagebox.showinfo("確認", "コピーする結果がありません(先に集計を実行してください)。")
            return
        # \r\nで行を区切ることで、Excelへ貼り付けた際に日付・人数・「人」が
        # それぞれ別セル、日ごとに別の行として分かれるようにする
        rows = "\r\n".join(f"{day}日\t{count}\t人" for day, count in self.day_counts)
        self.clipboard_clear()
        self.clipboard_append(rows)
        messagebox.showinfo("コピー完了", "日別人数をコピーしました。Excelのセルを選んで貼り付けてください。")

    def _save_csv(self):
        if not self.day_counts:
            messagebox.showinfo("確認", "保存する結果がありません(先に集計を実行してください)。")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile="専門的支援実施加算_日別人数.csv"
        )
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["日", "人数"])
            writer.writerows(self.day_counts)
        messagebox.showinfo("保存完了", f"{path} に保存しました。")


if __name__ == "__main__":
    app = App()
    app.mainloop()
