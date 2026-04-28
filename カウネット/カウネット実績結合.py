#このアプリは、複数のCSVファイルを選択して、特定の末尾番号に基づいて口座番号をマッピングし、最終行と「キャンセル」を含む行を削除してから、すべてのデータを結合して保存するものです。
import pandas as pd
import tkinter as tk
from tkinter import font, filedialog, messagebox
from pathlib import Path
import webbrowser
import json  # 設定保存用に追加

# 設定ファイルのパス（スクリプトと同じ場所に保存）
CONFIG_FILE = Path(__file__).parent / "config.json"

def load_config(default_map):
    """設定を読み込む。ファイルがなければデフォルトを返す"""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return default_map
    return default_map

def save_config(account_map):
    """設定をファイルに保存する"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(account_map, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"設定の保存に失敗しました: {e}")

def get_account_map_from_user(default_map):
    dialog = tk.Toplevel()
    dialog.title("口座番号マッピング設定")
    dialog.geometry("400x550")
    dialog.attributes("-topmost", True)
    
    is_cancelled = [True] 

    # --- 右クリックメニュー ---
    context_menu = tk.Menu(dialog, tearoff=0)
    context_menu.add_command(label="コピー", command=lambda: dialog.focus_get().event_generate("<<Copy>>"))
    context_menu.add_command(label="貼り付け", command=lambda: dialog.focus_get().event_generate("<<Paste>>"))
    context_menu.add_command(label="すべて選択", command=lambda: dialog.focus_get().event_generate("<<SelectAll>>"))

    def show_context_menu(event):
        event.widget.focus_set()
        context_menu.post(event.x_root, event.y_root)

    tk.Label(dialog, text="ファイル末尾(4桁) と 口座番号 を入力", font=("", 10, "bold")).pack(pady=10)

    container = tk.Frame(dialog)
    container.pack(pady=5)

    tk.Label(container, text="末尾4桁").grid(row=0, column=0, padx=5)
    tk.Label(container, text="口座番号").grid(row=0, column=1, padx=5)

    entries = []
    current_keys = list(default_map.keys())
    
    for i in range(10):
        k_val = current_keys[i] if i < len(current_keys) else ""
        v_val = default_map[k_val] if i < len(current_keys) else ""
        
        e1 = tk.Entry(container, width=10)
        e2 = tk.Entry(container, width=20)
        e1.insert(0, k_val)
        e2.insert(0, v_val)
        
        e1.grid(row=i+1, column=0, pady=2, padx=5)
        e2.grid(row=i+1, column=1, pady=2, padx=5)
        
        e1.bind("<Button-3>", show_context_menu) # Win
        e1.bind("<Button-2>", show_context_menu) # Mac
        e2.bind("<Button-3>", show_context_menu)
        e2.bind("<Button-2>", show_context_menu)
        
        entries.append((e1, e2))

    final_map = {}

    def on_confirm():
        for e1, e2 in entries:
            k, v = e1.get().strip(), e2.get().strip()
            if k and v:
                final_map[k] = v
        # 設定を保存
        save_config(final_map)
        is_cancelled[0] = False
        dialog.destroy()

    tk.Button(dialog, text="設定完了してCSV選択へ", command=on_confirm, bg="lightblue", width=25).pack(pady=15)

    # URL表示
    url_frame = tk.Frame(dialog)
    url_frame.pack(pady=5)
    tk.Label(url_frame, text="データダウンロードURLは").pack(side="left")
    url = "https://www.kaunet.com/kaunet/login/"
    link_label = tk.Label(url_frame, text=url, fg="blue", cursor="hand2")
    link_label.pack(side="left")
    
    f = font.Font(link_label, link_label.cget("font"))
    f.configure(underline=True)
    link_label.configure(font=f)
    link_label.bind("<Button-1>", lambda e: webbrowser.open(url))
    
    tk.Label(dialog, text="です。").pack()

    dialog.grab_set()
    dialog.wait_window()
    return None if is_cancelled[0] else final_map

def pre_process_csv():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # 初回起動時やファイルがない場合のデフォルト
    initial_default_map = {
        "4042": "6361863", "5757": "6354106",
        "8405": "6354102", "9873": "6361859"
    }

    # 保存された設定があれば読み込む
    account_map = get_account_map_from_user(load_config(initial_default_map))
    
    if account_map is None:
        root.destroy()
        return

    file_paths = filedialog.askopenfilenames(
        title="処理するCSVファイルをすべて選択してください", 
        filetypes=[("CSV files", "*.csv")]
    )
    
    if not file_paths:
        root.destroy()
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

            df = pd.read_csv(file_path, encoding='cp932', skiprows=4)
            df.columns = df.columns.str.strip()
            if len(df) > 0:
                df = df.drop(df.index[-1])
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
        root.destroy()

if __name__ == "__main__":
    pre_process_csv()