import pandas as pd
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import sys
import warnings
from datetime import datetime
from openpyxl.styles import PatternFill

warnings.simplefilter('ignore', UserWarning)

def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    target_ym = simpledialog.askstring("入力", "抽出する年月を入力してください\n（例: 202604）", parent=root)
    if not target_ym: return

    folder_path = filedialog.askdirectory(title="Excelファイルがあるフォルダを選択してください")
    if not folder_path: return

    target_files = [f for f in os.listdir(folder_path) 
                    if (f.endswith(".xlsx") or f.endswith(".xls")) and not f.startswith("~$")]

    combined_list = []
    
    progress_window = tk.Toplevel(root)
    progress_window.title("データ集計中...")
    progress_window.geometry("450x150")
    progress_window.attributes("-topmost", True)
    label = tk.Label(progress_window, text="準備中...", wraplength=400)
    label.pack(pady=10)
    pb = ttk.Progressbar(progress_window, orient="horizontal", length=350, mode="determinate")
    pb.pack(pady=5)
    pb["maximum"] = len(target_files)

    for i, filename in enumerate(target_files):
        pb["value"] = i + 1
        label.config(text=f"処理中 ({i+1}/{len(target_files)}):\n{filename}")
        progress_window.update()
        
        file_path = os.path.join(folder_path, filename)
        try:
            xl = pd.ExcelFile(file_path, engine='openpyxl')
            if target_ym not in xl.sheet_names: continue

            wb_in = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws_in = wb_in[target_ym]
            shop_name = ws_in["B3"].value if ws_in["B3"].value else "名称不明"
            wb_in.close()

            df_raw = pd.read_excel(file_path, sheet_name=target_ym, header=None, engine='openpyxl')
            
            target_row_idx, target_col_idx, date_col_idx = None, None, None
            for r_idx, row in df_raw.iterrows():
                row_str_list = [str(x).replace(" ", "").replace("　", "") for x in row]
                found_col = next((c_idx for c_idx, val in enumerate(row_str_list) if "弁当注文人数" in val), None)
                if found_col is not None:
                    target_row_idx, target_col_idx = r_idx, found_col
                    date_col_idx = next((c_idx for c_idx, val in enumerate(row_str_list) if "日付" in val), None)
                    break
            
            if target_row_idx is None: continue

            df = pd.read_excel(file_path, sheet_name=target_ym, header=target_row_idx, engine='openpyxl')
            col_name_bento = df.columns[target_col_idx]
            col_name_date = df.columns[date_col_idx] if date_col_idx is not None else "日付"

            df[col_name_bento] = pd.to_numeric(df[col_name_bento], errors='coerce')
            df_filtered = df[df[col_name_bento] > 0].dropna(subset=[col_name_bento]).copy()
            
            if not df_filtered.empty:
                df_filtered['ファイル名'] = filename
                df_filtered['店舗名'] = shop_name
                # ソート用に日付をdatetime型に変換（エラーはNaTにする）
                df_filtered['sort_date'] = pd.to_datetime(df_filtered[col_name_date], errors='coerce')
                
                res_df = df_filtered[['ファイル名', '店舗名', col_name_date, col_name_bento, 'sort_date']]
                res_df.columns = ['ファイル名', '店舗名', '日付', '弁当注文人数', 'sort_date']
                combined_list.append(res_df)
        except:
            continue

    progress_window.destroy()

    if combined_list:
        final_df = pd.concat(combined_list, ignore_index=True)
        
        # 1:ファイル名, 2:日付(sort_date) でソート
        final_df = final_df.sort_values(by=['ファイル名', 'sort_date']).reset_index(drop=True)
        
        # 表示形式を「M/D」に変更
        def format_md(x):
            try:
                if pd.isna(x): return ""
                return f"{x.month}/{x.day}"
            except:
                return str(x)
        
        final_df['日付'] = final_df['sort_date'].apply(format_md)
        # ソート用の一時列を削除
        final_df = final_df.drop(columns=['sort_date'])
        
        now = datetime.now().strftime("%Y%m%d_%H%M")
        output_name = f"まとめ_{target_ym}_{now}.xlsx"
        save_path = os.path.join(os.path.expanduser("~"), "Downloads", output_name)
        final_df.to_excel(save_path, index=False)
        
        # --- 装飾処理 ---
        wb = openpyxl.load_workbook(save_path)
        ws = wb.active
        fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        
        current_fill = False
        last_date = None
        for row_idx in range(2, ws.max_row + 1):
            date_val = str(ws.cell(row=row_idx, column=3).value)
            if date_val != last_date:
                current_fill = not current_fill
                last_date = date_val
            if current_fill:
                for col_idx in range(1, 5):
                    ws.cell(row=row_idx, column=col_idx).fill = fill

        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 15

        wb.save(save_path)
        messagebox.showinfo("完了", f"保存しました：\n{output_name}")
        os.startfile(save_path)
    else:
        messagebox.showwarning("結果", "データが見つかりませんでした。")

    root.destroy()
    sys.exit()


