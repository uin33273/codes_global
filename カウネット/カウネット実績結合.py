#このアプリは、複数のCSVファイルを選択して、特定の末尾番号に基づいて口座番号をマッピングし、最終行と「キャンセル」を含む行を削除してから、すべてのデータを結合して保存するものです。
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

def get_account_map_from_user(default_map):
    """
    口座番号のマッピングをユーザーに入力させるウィンドウ
    """
    # サブウィンドウを作成
    dialog = tk.Toplevel()
    dialog.title("口座番号マッピング設定")
    dialog.geometry("400x450")
    dialog.attributes("-topmost", True)  # 最前面に表示
    
    # 閉じるボタンが押された時の対策
    is_cancelled = [True] 

    tk.Label(dialog, text="ファイル末尾(4桁) と 口座番号 を入力", font=("", 10, "bold")).pack(pady=10)

    container = tk.Frame(dialog)
    container.pack(pady=5)

    tk.Label(container, text="末尾4桁").grid(row=0, column=0, padx=5)
    tk.Label(container, text="口座番号").grid(row=0, column=1, padx=5)

    entries = []
    default_keys = list(default_map.keys())
    
    for i in range(10):
        k_val = default_keys[i] if i < len(default_keys) else ""
        v_val = default_map[k_val] if i < len(default_keys) else ""
        
        e1 = tk.Entry(container, width=10)
        e2 = tk.Entry(container, width=20)
        e1.insert(0, k_val)
        e2.insert(0, v_val)
        
        e1.grid(row=i+1, column=0, pady=2, padx=5)
        e2.grid(row=i+1, column=1, pady=2, padx=5)
        entries.append((e1, e2))

    final_map = {}

    def on_confirm():
        for e1, e2 in entries:
            k, v = e1.get().strip(), e2.get().strip()
            if k and v:
                final_map[k] = v
        is_cancelled[0] = False
        dialog.destroy()

    tk.Button(dialog, text="設定完了してCSV選択へ", command=on_confirm, bg="lightblue", width=25).pack(pady=20)
    
    # ウィンドウが閉じられるまで待機
    dialog.grab_set()
    dialog.wait_window()
    
    return None if is_cancelled[0] else final_map

def pre_process_csv():
    # メインウィンドウの設定
    root = tk.Tk()
    root.withdraw()
    # 常に最前面でダイアログを出すための設定
    root.attributes("-topmost", True)

    # デフォルトのマッピング
    default_account_map = {
        "4042": "6361863",
        "5757": "6354106",
        "8405": "6354102",
        "9873": "6361845",
        "8397": "6361865"
    }

    # 入力画面を表示
    account_map = get_account_map_from_user(default_account_map)
    
    # キャンセル（右上の×で閉じるなど）された場合は終了
    if account_map is None:
        return

    # 1. 複数のCSVを選択
    file_paths = filedialog.askopenfilenames(
        title="処理するCSVファイルをすべて選択してください", 
        filetypes=[("CSV files", "*.csv")]
    )
    
    if not file_paths:
        return

    combined_list = []

    try:
        for p in file_paths:
            file_path = Path(p)
            file_name = file_path.stem
            
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

            # 最終行を削除
            if len(df) > 0:
                df = df.drop(df.index[-1])

            # 「キャンセル」行を削除
            if '伝票タイプ' in df.columns:
                df = df[~df['伝票タイプ'].str.contains('キャンセル', na=False)]

            df['口座番号'] = found_account
            combined_list.append(df)

        if not combined_list:
            messagebox.showwarning("警告", "対象の末尾番号を持つファイルがありませんでした。")
            return

        combined_df = pd.concat(combined_list, ignore_index=True)
        save_path = Path.home() / "Downloads" / "カウネット中間データ.csv"
        combined_df.to_csv(save_path, index=False, encoding='cp932')
        
        messagebox.showinfo("完了", f"保存完了: {save_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました:\n{e}")
    finally:
        root.destroy() # 最後に確実に破棄

if __name__ == "__main__":
    pre_process_csv()