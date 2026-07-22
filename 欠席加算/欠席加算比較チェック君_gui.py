# -*- coding: utf-8 -*-
"""
欠席チェック君 (GUI版)

やること:
  1. PDF「欠席時対応記録票」をウィンドウにドラッグ&ドロップ(またはボタンで選択)
  2. カレンダー画面からコピーしたテキストを、下のテキスト欄に貼り付け(Ctrl+V)
  3. 「比較する」ボタンを押すと、一致件数・不一致(片方にしかない組み合わせ)を
     ウィンドウ内に一覧表示する。不一致リストはCSVにも保存できる。

■事前準備 (初回のみ、コマンドプロンプトで実行)
    pip install pdfplumber tkinterdnd2

■起動方法
    python compare_absences_gui.py

  ドラッグ&ドロップができない場合は自動的に「PDFを選択」ボタンのみのモードで動作します。
"""

import csv
import re
import sys
import unicodedata
import calendar as calendar_mod
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

# tkinterdnd2があればドラッグ&ドロップ対応、無ければ通常のtkinterにフォールバック
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


# ============ 突き合わせロジック(compare_absences.pyと同じ考え方) ============

RELATION_WORDS = ["祖父母", "祖父", "祖母", "本人", "その他", "父", "母", "兄", "姉", "弟", "妹"]


def normalize_name(name: str) -> str:
    if name is None:
        return ""
    name = unicodedata.normalize("NFKC", name)
    name = name.upper()
    name = re.sub(r"\s+", "", name)
    return name


def extract_day(date_str: str):
    """日付文字列から末尾の"日"の数字だけを取り出す('7/16'→16, '16'→16)"""
    if date_str is None:
        return None
    s = unicodedata.normalize("NFKC", date_str).strip()
    m = re.search(r"(\d{1,2})\s*$", s)
    return int(m.group(1)) if m else None


def format_mmdd(date_str: str, month=None) -> str:
    """
    表示用に日付を 'MM/DD'(ゼロ埋め2桁) 形式へ整える。
    date_strに月が含まれていれば(例:'7/1')それを使い、
    月が無ければ(例:'1'だけ)引数のmonthを補って使う。
    """
    if date_str is None:
        return ""
    s = unicodedata.normalize("NFKC", date_str).strip()
    m = re.search(r"(?:(\d{1,2})[/\-])?(\d{1,2})\s*$", s)
    if not m:
        return s
    mm = m.group(1)
    dd = m.group(2)
    if mm is None:
        mm = month
    if mm is None:
        return f"?/{int(dd):02d}"
    return f"{int(mm):02d}/{int(dd):02d}"


def normalize_date(date_str: str) -> str:
    """
    同じ月のPDFとカレンダーを比較する前提のため、月の有無によらず
    末尾の"日"の数字だけを取り出して正規化する('7/1'も'1'も同じキーになる)。
    """
    if date_str is None:
        return ""
    date_str = unicodedata.normalize("NFKC", date_str).strip()
    m = re.search(r"(\d{1,2})\s*$", date_str)
    if not m:
        return date_str
    return str(int(m.group(1)))


def _extract_pdf_table_rows(pdf_path: str):
    """
    PDFの表を1行ずつパースして、
    [{'row_no':..., 'flag':..., 'name':..., 'date':(欠席日), 'uketsuke':(受付日),
      'has_blank':(記録票の必須項目に空欄があるか)}, ...]
    を返す(○が付いているかどうかに関わらず全行)。
    """
    if pdfplumber is None:
        raise RuntimeError("pdfplumberがインストールされていません。\npip install pdfplumber を実行してください。")

    rows_out = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                header = table[0]

                def find_col(keyword):
                    for i, h in enumerate(header):
                        if h and keyword in h.replace("\n", ""):
                            return i
                    return None

                def find_col_exact(keyword):
                    # "次回利用日"と"次回利用日の確認"のように、部分一致だと
                    # 別の列を誤って拾ってしまう場合があるための完全一致版
                    for i, h in enumerate(header):
                        if h and h.replace("\n", "") == keyword:
                            return i
                    return None

                idx_flag = find_col("算定")
                idx_name = find_col("児童氏名")
                idx_rel = find_col("続柄")
                idx_date = find_col("欠席日")
                idx_uketsuke = find_col("受付日")
                if idx_flag is None or idx_name is None or idx_date is None:
                    continue

                # 欠席加算を算定する場合に記入必須となる項目(ルール③)
                idx_reason = find_col("欠席理由")
                idx_child_state = find_col("児童の様子")
                idx_support = find_col("援助内容")
                idx_next_confirm = find_col_exact("次回利用日の確認")
                idx_next_date = find_col_exact("次回利用日")
                idx_recorder = find_col("記入者")

                # 「欠席加算」列の左隣にある通し番号列(ヘッダーは空欄)を行番号として使う
                idx_rowno = idx_flag - 1 if idx_flag is not None and idx_flag > 0 else None

                for row in table[1:]:
                    if not row or len(row) <= max(idx_flag, idx_name, idx_date):
                        continue
                    flag = (row[idx_flag] or "").strip()
                    name = (row[idx_name] or "").strip()
                    date = (row[idx_date] or "").strip()
                    uketsuke = (row[idx_uketsuke] or "").strip() if idx_uketsuke is not None else ""
                    rel_cell = (row[idx_rel] or "").strip() if idx_rel is not None else ""
                    row_no = (row[idx_rowno] or "").strip() if idx_rowno is not None else ""

                    final_relation = rel_cell
                    if rel_cell and rel_cell not in RELATION_WORDS:
                        # 氏名が長くて続柄列にはみ出しているケースを補正
                        combined = name + rel_cell
                        final_relation = ""
                        for w in RELATION_WORDS:
                            if combined.endswith(w):
                                name = combined[: -len(w)]
                                final_relation = w
                                break

                    def cell(idx):
                        if idx is None or idx >= len(row):
                            return ""
                        return (row[idx] or "").strip()

                    required_values = [
                        uketsuke,
                        final_relation,
                        cell(idx_reason),
                        cell(idx_child_state),
                        cell(idx_support),
                        cell(idx_next_confirm),
                        cell(idx_next_date),
                        cell(idx_recorder),
                    ]
                    has_blank = any(v == "" for v in required_values)

                    if name and date:
                        rows_out.append(
                            {
                                "row_no": row_no,
                                "flag": flag,
                                "name": name,
                                "date": date,
                                "uketsuke": uketsuke,
                                "has_blank": has_blank,
                            }
                        )
    return rows_out


def is_maru(flag: str) -> bool:
    """
    欠席加算(算定)欄の判定。「○」はもちろん、空欄(未記入)の場合も
    ○が入力されているものとして扱う。
    """
    return (flag or "").strip() in ("○", "")


def extract_pdf_rows(pdf_path: str):
    """PDFから「欠席加算=○(空欄含む)」の行を [{'date':..., 'name':..., 'row_no':..., 'has_blank':...}, ...] で返す"""
    all_rows = _extract_pdf_table_rows(pdf_path)
    return [
        {"date": r["date"], "name": r["name"], "row_no": r["row_no"], "has_blank": r["has_blank"]}
        for r in all_rows
        if is_maru(r["flag"])
    ]


def find_pdf_blank_flag_rows(pdf_path: str):
    """PDFの「欠席加算」欄が空欄(未記入)になっている行を返す。
    空欄は計算上○として扱われるが、記入漏れ確認用に一覧化する。
    [{'name':..., 'date':..., 'row_no':...}, ...]"""
    all_rows = _extract_pdf_table_rows(pdf_path)
    return [r for r in all_rows if (r["flag"] or "").strip() == ""]


def get_pdf_flag_map(pdf_path: str):
    """(欠席日, 氏名)ごとの「欠席加算」欄の記載内容とPDF行番号を返す(○以外も含む全行対象)。
    only_cal(HUGにはあるが日報側に○が無い)の表示で、本当に記載が無いのか
    「×」等が明示的に書いてあるのかを区別し、行番号も表示するために使う。
    戻り値: {(date_key, name_key): {'flag':..., 'row_no':...}, ...}"""
    all_rows = _extract_pdf_table_rows(pdf_path)
    result = {}
    for r in all_rows:
        key = (normalize_date(r["date"]), normalize_name(r["name"]))
        result[key] = {"flag": r["flag"], "row_no": r["row_no"]}
    return result


def find_pdf_duplicates(pdf_path: str):
    """PDF内で「氏名+欠席日」が完全一致する行(=重複入力の疑い)をまとめて返す。
    ○以外の行も含めて全行を対象にする。
    [{'name':..., 'date':..., 'row_nos': [...]}, ...]"""
    all_rows = _extract_pdf_table_rows(pdf_path)
    groups = {}
    for r in all_rows:
        key = (normalize_date(r["date"]), normalize_name(r["name"]))
        groups.setdefault(key, []).append(r)

    duplicates = []
    for rows in groups.values():
        if len(rows) > 1:
            duplicates.append(
                {
                    "name": rows[0]["name"],
                    "date": rows[0]["date"],
                    "row_nos": [r["row_no"] for r in rows],
                }
            )
    return duplicates


def parse_month_day(date_str: str):
    """'7/1'のような文字列を(月,日)のタプルにする。読み取れなければNone"""
    if not date_str:
        return None
    s = unicodedata.normalize("NFKC", date_str).strip()
    m = re.search(r"(\d{1,2})[/\-](\d{1,2})", s)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2))


def find_pdf_date_errors(pdf_path: str):
    """
    PDFの「欠席加算=○(空欄含む)」の行だけを対象に、受付日が欠席日より後になっている
    (＝欠席日 < 受付日、日付の誤記の疑いがある)行を
    [{'row_no':..., 'name':..., 'date':..., 'uketsuke':...}, ...] で返す。
    """
    all_rows = _extract_pdf_table_rows(pdf_path)
    errors = []
    for r in all_rows:
        if not is_maru(r["flag"]):
            continue
        d1 = parse_month_day(r["date"])       # 欠席日
        d2 = parse_month_day(r["uketsuke"])    # 受付日
        if d1 is not None and d2 is not None and d1 < d2:
            errors.append(r)
    return errors


def parse_calendar_text(text: str, month: int = None):
    """カレンダー画面のコピペテキストから「欠席Ⅰ」(加算なし除く)を抽出"""
    lines = text.splitlines()

    def is_day_header(line):
        m = re.fullmatch(r"(?:\S+\s+)?(\d{1,2})", line.strip())
        if m:
            n = int(m.group(1))
            if 1 <= n <= 31:
                return n
        return None

    results = []
    current_day = None
    state = None

    for line in lines:
        s = line.strip()
        if not s:
            continue
        if s.startswith("日\t月") or s.startswith("日 月"):
            continue

        dnum = is_day_header(s)
        if dnum is not None:
            current_day = dnum
            state = None
            continue

        if s == "休み":
            state = "skip"
            continue
        if re.fullmatch(r"出席(\d+)人", s):
            state = "attend"
            continue
        if re.fullmatch(r"欠席Ⅰ（加算なし）(\d+)人", s):
            state = "absent1_nokasan"
            continue
        if re.fullmatch(r"欠席Ⅰ(\d+)人", s):
            state = "absent1"
            continue
        if re.match(r"欠席Ⅱ", s):
            state = "other"
            continue

        if state == "absent1" and current_day is not None:
            date_label = f"{month}/{current_day}" if month else str(current_day)
            results.append({"date": date_label, "name": s})

    return results


def compare_rows(pdf_rows, cal_rows):
    def key(r):
        return (normalize_date(r["date"]), normalize_name(r["name"]))

    pdf_map = {key(r): r for r in pdf_rows}
    cal_map = {key(r): r for r in cal_rows}

    pdf_keys, cal_keys = set(pdf_map), set(cal_map)
    matched = pdf_keys & cal_keys
    only_pdf = sorted(pdf_keys - cal_keys)
    only_cal = sorted(cal_keys - pdf_keys)

    matched_pdf_rows = [pdf_map[k] for k in sorted(matched)]

    return matched_pdf_rows, [pdf_map[k] for k in only_pdf], [cal_map[k] for k in only_cal]


# ============================== GUI ==============================

BaseTk = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk


class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title("欠席チェック君")
        self.geometry("820x720")

        self.pdf_path = None
        self.mismatch_rows = []  # 保存用に最後の不一致リストを覚えておく

        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 10, "pady": 6}

        # --- PDFエリア ---
        frame_pdf = ttk.LabelFrame(self, text="① 欠席時対応記録票(日報PDF)")
        frame_pdf.pack(fill="x", **pad)

        drop_text = "ここにPDFをドラッグ&ドロップ" if DND_AVAILABLE else "PDFファイルを選択してください"
        self.pdf_label = tk.Label(
            frame_pdf, text=drop_text, bg="#f0f0f0", height=3, relief="groove"
        )
        self.pdf_label.pack(fill="x", padx=10, pady=(10, 4))

        if DND_AVAILABLE:
            self.pdf_label.drop_target_register(DND_FILES)
            self.pdf_label.dnd_bind("<<Drop>>", self._on_pdf_drop)

        ttk.Button(frame_pdf, text="PDFを選択...", command=self._choose_pdf).pack(pady=(0, 10))

        # --- カレンダーテキストエリア ---
        frame_cal = ttk.LabelFrame(self, text="② HUGカレンダー画面のテキストを貼り付け(全選択コピペでOK)")
        frame_cal.pack(fill="both", expand=True, **pad)

        top_row = ttk.Frame(frame_cal)
        top_row.pack(fill="x", padx=10, pady=(10, 0))
        tk.Label(top_row, text="今は何年何月ですか").pack(side="left")
        self.year_entry = ttk.Entry(top_row, width=6)
        self.year_entry.pack(side="left", padx=(6, 2))
        self.year_entry.insert(0, str(datetime.now().year))
        tk.Label(top_row, text="年").pack(side="left")
        self.month_entry = ttk.Entry(top_row, width=4)
        self.month_entry.pack(side="left", padx=(6, 2))
        self.month_entry.insert(0, str(datetime.now().month))
        tk.Label(top_row, text="月　(結果はMM/DD形式で表示されます)").pack(side="left")

        range_row = ttk.Frame(frame_cal)
        range_row.pack(fill="x", padx=10, pady=(6, 0))
        tk.Label(range_row, text="対象日付範囲:").pack(side="left")
        self.range_var = tk.StringVar(value="first_half")
        ttk.Radiobutton(
            range_row, text="1日〜15日", variable=self.range_var, value="first_half"
        ).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(
            range_row, text="16日〜月末", variable=self.range_var, value="second_half"
        ).pack(side="left", padx=(10, 0))

        self.cal_text = tk.Text(frame_cal, height=14, wrap="none")
        self.cal_text.pack(fill="both", expand=True, padx=10, pady=10)

        # --- 実行ボタン ---
        ttk.Button(self, text="③ 比較する", command=self._run_compare).pack(pady=6)

        # --- 結果表示エリア(左右にドラッグで幅調整できる分割) ---
        frame_result = ttk.LabelFrame(self, text="結果")
        frame_result.pack(fill="both", expand=True, **pad)

        paned = tk.PanedWindow(frame_result, orient="horizontal", sashrelief="raised", sashwidth=6)
        paned.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        left_frame = ttk.Frame(paned)
        ttk.Label(left_frame, text="不一致・日付誤記の疑い").pack(anchor="w")
        left_text_frame = ttk.Frame(left_frame)
        left_text_frame.pack(fill="both", expand=True)
        self.result_text = tk.Text(left_text_frame, height=14, wrap="word", state="disabled")
        left_scroll = ttk.Scrollbar(left_text_frame, orient="vertical", command=self.result_text.yview)
        self.result_text.config(yscrollcommand=left_scroll.set)
        left_scroll.pack(side="right", fill="y")
        self.result_text.pack(side="left", fill="both", expand=True)
        paned.add(left_frame, stretch="always")

        right_frame = ttk.Frame(paned)
        ttk.Label(right_frame, text="一致しているが〇の行に空欄がある").pack(anchor="w")
        right_text_frame = ttk.Frame(right_frame)
        right_text_frame.pack(fill="both", expand=True)
        self.blank_text = tk.Text(right_text_frame, height=14, wrap="word", state="disabled")
        right_scroll = ttk.Scrollbar(right_text_frame, orient="vertical", command=self.blank_text.yview)
        self.blank_text.config(yscrollcommand=right_scroll.set)
        right_scroll.pack(side="right", fill="y")
        self.blank_text.pack(side="left", fill="both", expand=True)
        paned.add(right_frame, stretch="always")

        ttk.Button(frame_result, text="不一致リストをCSVで保存...", command=self._save_csv).pack(pady=(0, 10))

    # ---------- イベント ----------

    def _on_pdf_drop(self, event):
        path = event.data.strip("{}")  # Windowsでスペースを含むパスは{}で囲まれる
        self._set_pdf(path)

    def _choose_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDFファイル", "*.pdf")])
        if path:
            self._set_pdf(path)

    def _set_pdf(self, path):
        if not path.lower().endswith(".pdf"):
            messagebox.showerror("エラー", "PDFファイルを指定してください。")
            return
        self.pdf_path = path
        self.pdf_label.config(text=f"選択済み: {path}")

        # 新しいPDFが来たら、前回分の貼り付けテキストと結果表示をクリアして
        # 新しい情報の受付待ち状態に戻す(年月欄は「今」で埋め直し、範囲は前半に戻す)
        self.cal_text.delete("1.0", "end")
        self.year_entry.delete(0, "end")
        self.year_entry.insert(0, str(datetime.now().year))
        self.month_entry.delete(0, "end")
        self.month_entry.insert(0, str(datetime.now().month))
        self.range_var.set("first_half")
        self.mismatch_rows = []
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.config(state="disabled")
        self.blank_text.config(state="normal")
        self.blank_text.delete("1.0", "end")
        self.blank_text.config(state="disabled")

    def _run_compare(self):
        if not self.pdf_path:
            messagebox.showwarning("確認", "PDFがまだ選択されていません。")
            return

        cal_text = self.cal_text.get("1.0", "end").strip()
        if not cal_text:
            messagebox.showwarning("確認", "カレンダーのテキストが貼り付けられていません。")
            return

        year_str = self.year_entry.get().strip()
        month_str = self.month_entry.get().strip()
        year = int(year_str) if year_str.isdigit() else None
        month = int(month_str) if month_str.isdigit() else None

        try:
            pdf_rows = extract_pdf_rows(self.pdf_path)
            date_errors = find_pdf_date_errors(self.pdf_path)
            pdf_flag_map = get_pdf_flag_map(self.pdf_path)
            pdf_duplicates = find_pdf_duplicates(self.pdf_path)
            pdf_blank_flag_rows = find_pdf_blank_flag_rows(self.pdf_path)
        except Exception as e:
            messagebox.showerror("PDF読み込みエラー", str(e))
            return

        cal_rows = parse_calendar_text(cal_text, month)

        if not pdf_rows:
            messagebox.showwarning("確認", "PDFから「欠席加算=○」の行が見つかりませんでした。フォーマットをご確認ください。")
        if not cal_rows:
            messagebox.showwarning("確認", "貼り付けたテキストから「欠席Ⅰ」が見つかりませんでした。コピー範囲・書式をご確認ください。")

        # 対象日付範囲(1日〜15日 / 16日〜月末)で絞り込む。月末は年月から算出する。
        is_first_half = self.range_var.get() == "first_half"
        if year and month:
            last_day = calendar_mod.monthrange(year, month)[1]
        else:
            last_day = 31  # 年月が未入力/不正な場合は安全側(31日まで)で判定

        def in_range(date_str):
            day = extract_day(date_str)
            if day is None:
                return True  # 日にちが読み取れない行はフィルタせずそのまま残す
            return (1 <= day <= 15) if is_first_half else (16 <= day <= last_day)

        pdf_rows = [r for r in pdf_rows if in_range(r["date"])]
        cal_rows = [r for r in cal_rows if in_range(r["date"])]
        date_errors = [r for r in date_errors if in_range(r["date"])]  # 欠席日を基準に範囲を揃える
        pdf_duplicates = [d for d in pdf_duplicates if in_range(d["date"])]
        pdf_blank_flag_rows = [r for r in pdf_blank_flag_rows if in_range(r["date"])]

        matched, only_pdf, only_cal = compare_rows(pdf_rows, cal_rows)

        # 「HUG(カレンダー)」「日報(PDF)」それぞれに欠席Ⅰ/欠席加算〇があるかを
        # date・氏名ごとにまとめ、日付順に並べる(日報側にPDFの行番号・空欄有無があれば付ける)
        merged = (
            [
                (
                    "欠席チェック不一致",
                    format_mmdd(r["date"], month),
                    r["name"],
                    "記載なし",
                    "〇",
                    r.get("row_no", ""),
                    "",
                    "あり" if r.get("has_blank") else "",
                )
                for r in only_pdf
            ]
            + [
                (
                    "欠席チェック不一致",
                    format_mmdd(r["date"], month),
                    r["name"],
                    "〇",
                    pdf_info["flag"] if pdf_info else "記載なし",
                    pdf_info["row_no"] if pdf_info else "",
                    "",
                    "",
                )
                for r in only_cal
                for pdf_info in [pdf_flag_map.get((normalize_date(r["date"]), normalize_name(r["name"])))]
            ]
        )
        merged.sort(key=lambda x: (x[1], x[2]))

        # 欠席日 < 受付日 になっている行(日付の誤記の疑い)
        date_error_rows = [
            (
                "日付誤記疑い",
                format_mmdd(r["date"], month),
                r["name"],
                "-",
                "-",
                r["row_no"],
                format_mmdd(r["uketsuke"], month),
                "",
            )
            for r in date_errors
        ]
        date_error_rows.sort(key=lambda x: (x[1], x[2]))

        # 一致しているが、日報側の〇の行に空欄がある(記入漏れの疑い)ものを別枠でまとめる
        matched_blank_rows = [
            (r["name"], r.get("row_no", ""), format_mmdd(r["date"], month))
            for r in matched
            if r.get("has_blank")
        ]
        matched_blank_rows.sort(key=lambda x: (x[2], x[0]))

        pdf_duplicates.sort(key=lambda d: (normalize_date(d["date"]), d["name"]))
        pdf_blank_flag_rows.sort(key=lambda r: (normalize_date(r["date"]), r["name"]))

        self.mismatch_rows = merged + date_error_rows
        range_label = "1日〜15日" if is_first_half else f"16日〜{last_day}日"
        self._show_result(
            len(matched), merged, date_error_rows, matched_blank_rows, pdf_duplicates, pdf_blank_flag_rows, month, range_label
        )

    def _show_result(
        self, matched_count, merged, date_error_rows, matched_blank_rows, pdf_duplicates, pdf_blank_flag_rows, month, range_label=""
    ):
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")

        lines = [f"■ 対象範囲: {range_label}", f"■ HUGと日報の一致: {matched_count}件", ""]
        lines.append(f"■ 不一致(HUGと日報で欠席Ⅰ/欠席加算〇が食い違っている): {len(merged)}件")
        for _, date, name, hug, nippo, row_no, _uke, has_blank in merged:
            row_no_disp = row_no if row_no else "-"
            blank_part = "（空欄あり）" if has_blank == "あり" else ""
            lines.append(f"{name}\t{row_no_disp}\t{date}\tHUG:{hug}、日報:{nippo}{blank_part}")

        lines.append("")
        lines.append(f"■ 日付の誤記の疑い(〇の行のみ対象・受付日が欠席日より後になっている): {len(date_error_rows)}件")
        for _, date, name, _hug, _nippo, row_no, uketsuke, _blank in date_error_rows:
            lines.append(f"{name}\t{row_no}\t{date}\t受付日:{uketsuke}　欠席日:{date}")

        lines.append("")
        lines.append(f"■ 重複の疑い(氏名+欠席日がPDF内で完全一致): {len(pdf_duplicates)}件")
        for d in pdf_duplicates:
            row_no_str = "と".join((f"{rn}番" if rn else "-") for rn in d["row_nos"])
            lines.append(f"{d['name']}\t{format_mmdd(d['date'], month)}\t行番号{row_no_str}\t重複")

        self.result_text.insert("1.0", "\n".join(lines))
        self.result_text.config(state="disabled")

        # 右側: 一致しているが〇の行に空欄がある(記入漏れの疑い)
        self.blank_text.config(state="normal")
        self.blank_text.delete("1.0", "end")
        blank_lines = [f"■ 該当: {len(matched_blank_rows)}件", ""]
        for name, row_no, date in matched_blank_rows:
            row_no_disp = row_no if row_no else "-"
            blank_lines.append(f"{name}\t{row_no_disp}\t{date}\t（空欄あり）")

        blank_lines.append("")
        blank_lines.append(f"■ 欠席加算欄が未記入(空欄): {len(pdf_blank_flag_rows)}件")
        for r in pdf_blank_flag_rows:
            row_no_disp = r["row_no"] if r["row_no"] else "-"
            blank_lines.append(f"{r['name']}\t{row_no_disp}\t{format_mmdd(r['date'], month)}")

        self.blank_text.insert("1.0", "\n".join(blank_lines))
        self.blank_text.config(state="disabled")

    def _save_csv(self):
        if not self.mismatch_rows:
            messagebox.showinfo("確認", "保存する不一致がありません(先に比較を実行してください)。")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile="mismatch.csv"
        )
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["区分", "date", "name", "HUG", "日報", "PDF行番号", "受付日(誤記疑いのみ)", "空欄(日報〇のみ)"])
            writer.writerows(self.mismatch_rows)
        messagebox.showinfo("保存完了", f"{path} に保存しました。")


if __name__ == "__main__":
    if not DND_AVAILABLE:
        print("※ tkinterdnd2が見つからないため、ドラッグ&ドロップは使えません。")
        print("  pip install tkinterdnd2 を実行すると有効になります。")
    app = App()
    app.mainloop()
