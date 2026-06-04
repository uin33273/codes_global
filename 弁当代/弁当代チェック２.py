import pandas as pd
import openpyxl
import os
import zipfile
import tempfile
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import sys
import warnings
from datetime import datetime
from openpyxl.styles import PatternFill, Font

warnings.simplefilter('ignore', UserWarning)

REF_PATH = r"C:\Users\owner\Desktop\works\保管書類\codes\弁当代\table_HHHL共通店舗一覧m.xlsm"

def load_url_map():
    url_map = {}
    try:
        wb = openpyxl.load_workbook(REF_PATH, read_only=True, keep_vba=True, data_only=True)
        ws = wb["リスト"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            shop = row[0]
            url  = row[3]
            if shop:
                url_map[str(shop).strip()] = str(url).strip() if url else ""
        wb.close()
    except Exception as e:
        messagebox.showwarning("参照ファイル読み込みエラー", f"レク費urlリストを読み込めませんでした:\n{e}")
    return url_map

def main():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    url_map = load_url_map()

    target_ym = simpledialog.askstring("入力", "抽出する年月を入力してください\n（例: 202604）", parent=root)
    if not target_ym: return

    # フォルダまたはZIP選択ダイアログ
    choice_win = tk.Toplevel(root)
    choice_win.title("入力元選択")
    choice_win.geometry("400x110")
    choice_win.attributes("-topmost", True)
    choice_win.resizable(False, False)
    choice_win.grab_set()

    folder_path = [None]
    tmpdir      = [None]

    def pick_folder():
        path = filedialog.askdirectory(title="エクセルがあるフォルダを選択してください")
        if path:
            folder_path[0] = path
        choice_win.destroy()

    def pick_zip():
        path = filedialog.askopenfilename(
            title="エクセルがあるZIPファイルを選択してください",
            filetypes=[("ZIPファイル", "*.zip")]
        )
        if path:
            td = tempfile.mkdtemp()
            with zipfile.ZipFile(path, 'r') as zf:
                zf.extractall(td)
            tmpdir[0] = td
            found_dir = td
            for dirpath, _, filenames in os.walk(td):
                if any(f.endswith((".xlsx", ".xls")) and not f.startswith("~$") for f in filenames):
                    found_dir = dirpath
                    break
            folder_path[0] = found_dir
        choice_win.destroy()

    tk.Label(choice_win, text="エクセルがあるフォルダまたはZIPファイルを選択してください", pady=8).pack()
    btn_frame = tk.Frame(choice_win)
    btn_frame.pack()
    tk.Button(btn_frame, text="フォルダを選択",    width=18, height=2, command=pick_folder).pack(side=tk.LEFT, padx=12)
    tk.Button(btn_frame, text="ZIPファイルを選択", width=18, height=2, command=pick_zip).pack(side=tk.LEFT, padx=12)

    choice_win.wait_window()
    if not folder_path[0]: return
    folder_path = folder_path[0]

    try:
        target_files = [f for f in os.listdir(folder_path)
                        if (f.endswith(".xlsx") or f.endswith(".xls")) and not f.startswith("~$")]

        combined_list = []
        zero_shops    = []

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
                col_name_date  = df.columns[date_col_idx] if date_col_idx is not None else "日付"

                def is_text_entry(x):
                    if pd.isna(x): return False
                    if isinstance(x, (int, float)): return False
                    s = str(x).strip()
                    if s == "": return False
                    try:
                        float(s)
                        return False
                    except ValueError:
                        return True

                mask_text  = df[col_name_bento].apply(is_text_entry)
                df_text    = df[mask_text].copy()

                df['_bento_num'] = pd.to_numeric(df[col_name_bento], errors='coerce')
                total        = df['_bento_num'].sum()
                mask_numeric = df['_bento_num'] > 0
                df_numeric   = df[mask_numeric].copy()

                df_text['_bento_num']      = df_text[col_name_bento]
                df_numeric[col_name_bento] = df_numeric['_bento_num']

                df_filtered = pd.concat([df_numeric, df_text]).drop_duplicates()

                if df_filtered.empty and total <= 0:
                    zero_shops.append({
                        'ファイル名': filename,
                        '店舗名': shop_name,
                        'url': url_map.get(str(shop_name).strip(), "")
                    })
                    continue

                if not df_filtered.empty:
                    df_filtered = df_filtered.copy()
                    df_filtered['ファイル名'] = filename
                    df_filtered['店舗名']    = shop_name
                    df_filtered['sort_date'] = pd.to_datetime(df_filtered[col_name_date], errors='coerce')

                    res_df = df_filtered[['ファイル名', '店舗名', col_name_date, col_name_bento, 'sort_date']]
                    res_df.columns = ['ファイル名', '店舗名', '日付', '弁当注文人数', 'sort_date']
                    combined_list.append(res_df)
                else:
                    zero_shops.append({
                        'ファイル名': filename,
                        '店舗名': shop_name,
                        'url': url_map.get(str(shop_name).strip(), "")
                    })

            except:
                continue

        progress_window.destroy()

        if combined_list or zero_shops:
            now = datetime.now().strftime("%Y%m%d_%H%M")
            output_name = f"弁当チェック_{target_ym}_{now}.xlsx"
            save_path = os.path.join(os.path.expanduser("~"), "Downloads", output_name)

            if combined_list:
                final_df = pd.concat(combined_list, ignore_index=True)
                final_df = final_df.sort_values(by=['ファイル名', 'sort_date']).reset_index(drop=True)

                def format_md(x):
                    try:
                        if pd.isna(x): return ""
                        return f"{x.month}/{x.day}"
                    except:
                        return str(x)

                final_df['日付'] = final_df['sort_date'].apply(format_md)
                final_df = final_df.drop(columns=['sort_date'])

                final_df.insert(0, 'レク費url', final_df['店舗名'].apply(
                    lambda s: url_map.get(str(s).strip(), "")
                ))
            else:
                final_df = pd.DataFrame(columns=['レク費url', 'ファイル名', '店舗名', '日付', '弁当注文人数'])

            final_df.to_excel(save_path, index=False)

            wb = openpyxl.load_workbook(save_path)
            ws = wb.active

            # A列をハイパーリンクに変換
            hyperlink_font = Font(color="0563C1", underline="single")
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=1)
                url_val = cell.value
                if url_val and str(url_val).startswith("http"):
                    cell.hyperlink = url_val
                    cell.value     = url_val
                    cell.font      = hyperlink_font

            # 縞模様（日付列=4列目で切り替え）
            fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            current_fill = False
            last_date = None
            for row_idx in range(2, ws.max_row + 1):
                date_val = str(ws.cell(row=row_idx, column=4).value)
                if date_val != last_date:
                    current_fill = not current_fill
                    last_date = date_val
                if current_fill:
                    for col_idx in range(1, 6):
                        ws.cell(row=row_idx, column=col_idx).fill = fill_gray

            ws.column_dimensions['A'].width  = 10
            ws.column_dimensions['B'].hidden = True
            ws.column_dimensions['C'].width  = 30
            ws.column_dimensions['D'].width  = 10
            ws.column_dimensions['E'].width  = 15

            # 末尾に注文なし店舗セクション
            if zero_shops:
                blank_row = ws.max_row + 2
                fill_h = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                fill_r = PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")
                bold   = Font(bold=True)
                for col_idx, text in enumerate(["【注文なし店舗】レク費url", "ファイル名", "店舗名"], start=1):
                    cell = ws.cell(row=blank_row, column=col_idx, value=text)
                    cell.fill = fill_h
                    cell.font = bold
                for shop in zero_shops:
                    blank_row += 1
                    url_val  = shop.get('url', "")
                    url_cell = ws.cell(row=blank_row, column=1)
                    url_cell.fill = fill_r
                    if url_val and str(url_val).startswith("http"):
                        url_cell.hyperlink = url_val
                        url_cell.value     = url_val
                        url_cell.font      = Font(color="0563C1", underline="single")
                    ws.cell(row=blank_row, column=2, value=shop['ファイル名']).fill = fill_r
                    ws.cell(row=blank_row, column=3, value=shop['店舗名']).fill     = fill_r

            wb.save(save_path)
            messagebox.showinfo("完了", f"保存しました：\n{output_name}")
            os.startfile(save_path)
        else:
            messagebox.showwarning("結果", "データが見つかりませんでした。")

    finally:
        if tmpdir[0] and os.path.exists(tmpdir[0]):
            import shutil
            shutil.rmtree(tmpdir[0], ignore_errors=True)

    root.destroy()
    sys.exit()

if __name__ == "__main__":
    main()