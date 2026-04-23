import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from openpyxl import load_workbook

def convert_store_name(dept_name):
    if pd.isna(dept_name): return ""
    name = str(dept_name)
    for target, rep in [("グローバル", ""), ("キッズ", "K"), ("メソッド", "M"), ("パーク", "P"), ("サカフル", "SF"), ("店", "")]:
        name = name.replace(target, rep)
    return name.strip()

class ManualSelectionWindow:
    def __init__(self, target_stores, master_list):
        self.root = tk.Toplevel()
        self.root.title("店舗マスタ一括引き当て")
        self.root.geometry("600x550")
        self.root.attributes('-topmost', True)
        self.master_list = ["一致なし"] + master_list
        self.results = {}
        self.target_stores = target_stores
        self.current_idx = 0
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        self.label_info = tk.Label(main_frame, text="", font=("MS Gothic", 12, "bold"), fg="blue")
        self.label_info.pack(pady=5)
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.update_list)
        self.entry_search = tk.Entry(main_frame, textvariable=self.search_var, width=50)
        self.entry_search.pack(pady=5)
        self.entry_search.bind("<Return>", self.confirm_selection)
        self.entry_search.bind("<Down>", self.focus_list)
        self.listbox = tk.Listbox(main_frame, width=60, height=15, font=("MS Gothic", 10))
        self.listbox.pack(pady=5)
        self.listbox.bind("<Double-Button-1>", self.confirm_selection)
        self.listbox.bind("<Return>", self.confirm_selection)
        self.btn_ok = tk.Button(main_frame, text="確定して次へ (Enter)", command=self.confirm_selection, bg="lightblue", width=20)
        self.btn_ok.pack(pady=10)
        self.load_store()
        self.root.grab_set()
        self.root.wait_window()

    def focus_list(self, event):
        self.listbox.focus_set()
        if self.listbox.size() > 0: self.listbox.selection_set(0)

    def load_store(self):
        if self.current_idx < len(self.target_stores):
            store = self.target_stores[self.current_idx]
            self.label_info.config(text=f"【確認中】 {self.current_idx+1}/{len(self.target_stores)} ： {store}")
            self.search_var.set(store)
            self.update_list()
            self.entry_search.focus_set()
            self.entry_search.selection_range(0, tk.END)
        else: self.root.destroy()

    def update_list(self, *args):
        search_term = self.search_var.get()
        self.listbox.delete(0, tk.END)
        filtered_list = ["一致なし"] + [m for m in self.master_list[1:] if search_term in m]
        for item in filtered_list: self.listbox.insert(tk.END, item)
        if self.listbox.size() > 0: self.listbox.selection_set(0)

    def confirm_selection(self, event=None):
        selection = self.listbox.curselection()
        if selection:
            self.results[self.target_stores[self.current_idx]] = self.listbox.get(selection)
            self.current_idx += 1
            self.load_store()

def aggregate_orders_final(df):
    """注文番号に基づいて重複を排除し、金額合計と商品名加工を行う"""
    # 注文番号を文字列に統一
    df['注文番号'] = df['注文番号'].astype(str)
    
    # 1. 各注文番号の合計金額を計算
    amt_map = df.groupby('注文番号')['税込金額'].sum()
    
    # 2. 加工商品名（最大金額 + 「他」）のリスト作成
    def get_complex_name(group):
        max_row = group.loc[group['税込金額'].idxmax()]
        base_name = str(max_row['商品名'])
        return f"{base_name} 他" if len(group) > 1 else base_name

    name_map = df.groupby('注文番号').apply(get_complex_name)

    # 3. 重複を削除して1行目だけを残す
    df_unique = df.drop_duplicates(subset='注文番号', keep='first').copy()

    # 4. 集計した値を反映
    df_unique['税込金額'] = df_unique['注文番号'].map(amt_map)
    df_unique['商品名'] = df_unique['注文番号'].map(name_map)
    
    return df_unique

def main():
    root = tk.Tk()
    root.withdraw()
    csv_path = filedialog.askopenfilename(title="CSV選択", filetypes=[("CSV", "*.csv")])
    if not csv_path: return
    messagebox.showinfo("案内", "店番リスト(Excel)を選択してください。")
    master_path = filedialog.askopenfilename(title="マスタ選択", filetypes=[("Excel", "*.xlsx")])
    if not master_path: return

    try:
        master_df = pd.read_excel(master_path, sheet_name="リスト", usecols="C")
        master_choices = master_df.iloc[:, 0].dropna().astype(str).tolist()

        df = pd.read_csv(csv_path, skiprows=4, encoding='shift_jis')
        # A列のフッター削除
        df = df[~df[df.columns[0]].astype(str).str.contains("※税抜小計や税額小計", na=False)]

        # --- 【重要】ここが集約の核心部 ---
        if all(col in df.columns for col in ['注文番号', '税込金額', '商品名']):
            df = aggregate_orders_final(df)

        if "お届け先部署名" in df.columns:
            df.insert(df.columns.get_loc("お届け先部署名") + 1, "店舗名", df["お届け先部署名"].apply(convert_store_name))
            df.insert(df.columns.get_loc("店舗名") + 1, "店番付き店舗名", "")
            
            unmatched_unique = []
            for i, row in df.iterrows():
                s_name = str(row["店舗名"]).strip()
                if not s_name: continue
                match = [m for m in master_choices if s_name in m]
                if match: df.at[i, "店番付き店舗名"] = match[0]
                else: unmatched_unique.append(s_name)

            if unmatched_unique:
                selector = ManualSelectionWindow(sorted(list(set(unmatched_unique))), master_choices)
                for s, m in selector.results.items():
                    df.loc[df["店舗名"] == s, "店番付き店舗名"] = m

        excel_path = os.path.splitext(csv_path)[0] + ".xlsx"
        df.to_excel(excel_path, index=False)

        delete_list = ["伝票タイプ", "お支払い方法", "商品番号", "商品カテゴリ(大)", "商品カテゴリ(中)", "商品カテゴリ(小)", "行メモ", "税込単価", "税抜単価", "消費税額", "税抜小計", "税額小計", "税率", "軽減税率対象商品", "GPNDB掲載", "グリーン購入法", "ご登録電話番号", "担当販売店名", "ユーザーID", "ユーザー名", "備考", "エコマーク", "お届け先社名"]
        wb = load_workbook(excel_path)
        ws = wb.active
        for col_idx in range(ws.max_column, 0, -1):
            if str(ws.cell(row=1, column=col_idx).value).strip() in delete_list:
                ws.delete_cols(col_idx)
        
        wb.save(excel_path)
        messagebox.showinfo("完了", "注文を1行に集約し、整形を完了しました。")

    except Exception as e:
        messagebox.showerror("エラー", f"失敗しました:\n{e}")

if __name__ == "__main__":
    main()
