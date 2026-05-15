#店名をスプレッドシートのマスターと照合して紐づけるツール
import os
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import ctypes
import sys

# IME操作関数
def set_ime_on(widget):
    try:
        hwnd = widget.winfo_id()
        imm32 = ctypes.windll.imm32
        handle = imm32.ImmGetContext(hwnd)
        imm32.ImmSetOpenStatus(handle, 1)
        imm32.ImmReleaseContext(hwnd, handle)
    except: pass

def get_clean_place(text, is_csv=False):
    t = str(text)
    target = t.split('_')[0] if is_csv and '_' in t else t
    clean = re.sub(r'[0-9a-zA-Z!-~ 　\(\)（）/／]+', '', target)
    for word in ["店", "店舗", "支店", "メソッド"]:
        clean = clean.replace(word, "")
    return clean

class SearchBox(tk.Frame):
    def __init__(self, master, full_options, initial_val="", **kwargs):
        super().__init__(master, **kwargs)
        self.full_options = [str(opt) for opt in full_options]
        display_val = str(initial_val[0]) if isinstance(initial_val, list) and len(initial_val) > 0 else (str(initial_val) if initial_val else "")
        self.entry = tk.Entry(self, font=("Yu Gothic", 10))
        self.entry.insert(0, display_val)
        self.entry.pack(fill="x", side="top")
        
        self.entry.bind('<KeyRelease>', self.on_key)
        self.entry.bind('<FocusIn>', lambda e: set_ime_on(self.entry))
        self.entry.bind('<Return>', self.confirm_and_next)
        self.entry.bind('<Down>', lambda e: self.listbox.focus_set())

        self.listbox = tk.Listbox(self, height=4, font=("Yu Gothic", 10), exportselection=False)
        self.listbox.pack(fill="x", side="top")
        
        self.listbox.bind('<Double-Button-1>', self.confirm_selection)
        self.listbox.bind('<Return>', self.confirm_selection)
        
        self.update_list(display_val)

    def on_key(self, event):
        if event.keysym in ('Down', 'Up', 'Return'): return
        self.update_list(self.entry.get())

    def update_list(self, term):
        self.listbox.delete(0, tk.END)
        matches = [opt for opt in self.full_options if str(term).lower() in opt.lower()]
        for m in matches: self.listbox.insert(tk.END, m)
        if self.listbox.size() > 0: self.listbox.select_set(0)

    def confirm_selection(self, event=None):
        if self.listbox.curselection():
            selected = self.listbox.get(self.listbox.curselection())
            self.entry.delete(0, tk.END)
            self.entry.insert(0, selected)
            return True
        return False

    def confirm_and_next(self, event=None):
        self.confirm_selection()
        self.entry.event_generate("<<NextWidget>>")

    def get(self): return self.entry.get()

class ShopNameMatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("算定区分04: 店名紐づけ")
        self.root.geometry("850x850")
        self.root.attributes('-topmost', True)
       # self.root.withdraw() 
        self.results = {}
        self.root.bind("<Escape>", lambda e: self.root.destroy())

    def start_process(self):
        downloads = Path.home() / "Downloads"
        
        path1 = filedialog.askopenfilename(title="1つ目：anyfiles_to_1files.xlsxを選択", initialdir=str(downloads), parent=self.root)
        if not path1: self.root.destroy(); return
        df_csv = pd.read_excel(path1)

        path2 = filedialog.askopenfilename(title="2つ目：算定区分・加算等集計表（運営用）.xlsxを選択", initialdir=str(downloads), parent=self.root)
        if not path2: self.root.destroy(); return
        
        try:
            # header=2は3行目が列名という意味
            df_xlsx = pd.read_excel(path2, sheet_name='集計', header=2)
            xlsx_master = sorted(df_xlsx["店舗名"].dropna().astype(str).unique().tolist())
        except Exception as e:
            messagebox.showerror("エラー", f"Excelの読み込みに失敗しました。\n{e}", parent=self.root)
            self.root.destroy(); return

        while len(df_csv.columns) < 11: df_csv[f'Ex_{len(df_csv.columns)}'] = ""
        k_idx = 10 
        pending_list = []

        for idx in range(len(df_csv)):
            raw_name = str(df_csv.iloc[idx, 1])
            target_place = get_clean_place(raw_name, is_csv=True)
            place_matches = [m for m in xlsx_master if target_place == get_clean_place(m)]
            suffix = "P" if ("児" in raw_name or "P" in raw_name.upper()) else "M"
            final_m = [m for m in place_matches if suffix in str(m).upper()]

            if len(final_m) == 1: df_csv.iloc[idx, k_idx] = final_m[0]
            else: pending_list.append({'idx': idx, 'raw': raw_name, 'matches': final_m if final_m else place_matches, 'all_master': xlsx_master})

        if pending_list:
            self.root.deiconify()
            self.setup_ui(pending_list, df_csv, path1)
        else:
            self.save_and_exit(df_csv, path1)

    # --- インデントを修正したメソッド群 ---
    def setup_ui(self, pending_list, df_csv, path1):
        tk.Label(self.root, text="地名入力で絞り込み ➔ 矢印キーで選択 ➔ Enterで確定", font=("Yu Gothic", 12, "bold"), fg="#0056b3").pack(pady=15)
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw", width=800)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self.widgets = []
        for i, item in enumerate(pending_list):
            f = ttk.LabelFrame(scroll_frame, text=f" 元の店名: {item['raw']} ")
            f.pack(fill="x", padx=30, pady=5)
            sb = SearchBox(f, item['all_master'], initial_val=item['matches'])
            sb.pack(fill="x", padx=15, pady=5)
            sb.entry.bind("<<NextWidget>>", lambda e, idx=i: self.focus_next(idx))
            self.widgets.append((item['idx'], sb))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        tk.Button(self.root, text=" 修正内容をExcelに保存して次へ ", bg="#28a745", fg="white", 
                  font=("Yu Gothic", 13, "bold"), command=lambda: self.save_and_exit(df_csv, path1), pady=10).pack(fill="x", padx=60, pady=20)

    def focus_next(self, current_idx):
        if current_idx + 1 < len(self.widgets):
            w = self.widgets[current_idx + 1][1].entry
            w.focus_set(); set_ime_on(w)

    def save_and_exit(self, df_csv, path1):
        if hasattr(self, 'widgets'):
            for idx, sb in self.widgets: df_csv.iloc[idx, 10] = sb.get()
        
        df_csv.columns = list(df_csv.columns[:10]) + ["店舗名"] + list(df_csv.columns[11:])
        out = os.path.join(os.path.dirname(path1), "2-スプレッドシートの店名へ紐づけ.xlsx")
        try:
            df_csv.to_excel(out, index=False)
            messagebox.showinfo("完了", f"保存しました:\n{out}", parent=self.root)
        except Exception as e:
            messagebox.showerror("保存エラー", f"ファイルが保存できません。開いたままではありませんか？\n{e}", parent=self.root)
        
        # self.root.quit() # mainloopを抜ける
        self.root.destroy()

def main(root=None):
    if root is None:
        # 単体起動の場合
        root = tk.Tk()
        standalone = True
    else:
        # 親スクリプトから呼ばれた場合
        root = tk.Toplevel(root)
        standalone = False

    # ウィンドウ設定
    root.deiconify() 
    root.lift()
    root.focus_force()
    
    app = ShopNameMatcherApp(root)
    
    # 起動後に処理を開始
    root.after(100, app.start_process)
    
    if standalone:
        root.mainloop()
    else:
        # 紐づけ作業（ウィンドウの消滅）が終わるまで、親プログラムをここで待機させる
        root.wait_window()

if __name__ == "__main__":
    main()