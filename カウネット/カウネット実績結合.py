#このアプリは、複数のCSVファイルを選択して、特定の末尾番号に基づいて口座番号をマッピングし、最終行と「キャンセル」を含む行を削除してから、すべてのデータを結合して保存するものです。
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

def pre_process_csv():
    root = tk.Tk()
    root.withdraw()
    
    # 1. 複数のCSVを選択
    messagebox.showinfo("準備", "処理するCSVファイルをすべて選択してください")
    file_paths = filedialog.askopenfilenames(
        title="CSV選択", 
        filetypes=[("CSV files", "*.csv")]
    )
    
    if not file_paths:
        return

    # 口座番号のマッピング定義
    account_map = {
        "4042": "6361863",
        "5757": "6354106",
        "8405": "6354102",
        "9873": "6361845",
        "8397": "6361865"
    }

    combined_list = []

    try:
        for p in file_paths:
            file_path = Path(p)
            file_name = file_path.stem
            
            # ファイル末尾4桁を確認
            found_account = ""
            for suffix, acc_num in account_map.items():
                if file_name.endswith(suffix):
                    found_account = acc_num
                    break
            
            if not found_account:
                continue

            # CSV読み込み（最初の4行を削除）
            df = pd.read_csv(file_path, encoding='cp932', skiprows=4)
            df.columns = df.columns.str.strip()

            # --- 最終行を削除 ---
            if len(df) > 0:
                df = df.drop(df.index[-1])

            # --- 【追加】「キャンセル」を含む行を削除 ---
            if '伝票タイプ' in df.columns:
                # 「キャンセル」という文字が含まれる行を除外（NaNは無視）
                df = df[~df['伝票タイプ'].str.contains('キャンセル', na=False)]

            # 最右端に「口座番号」列を作成して入力
            df['口座番号'] = found_account
            
            combined_list.append(df)

        if not combined_list:
            messagebox.showwarning("警告", "対象の末尾番号を持つファイルがありませんでした。")
            return

        # すべてのデータを結合
        combined_df = pd.concat(combined_list, ignore_index=True)

        # 保存
        save_path = Path.home() / "Downloads" / "カウネット中間データ.csv"
        combined_df.to_csv(save_path, index=False, encoding='cp932')
        
        messagebox.showinfo("完了", f"前処理完了（キャンセル行・合計行 削除済み）\n保存先: {save_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"前処理中にエラーが発生しました:\n{e}")

if __name__ == "__main__":
    pre_process_csv()