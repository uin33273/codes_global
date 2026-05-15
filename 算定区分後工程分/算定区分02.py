#サーバーからダウンロードしたZIPファイルや、ローカルのフォルダ内のCSV/Excelファイルの総計行を削除、利用サービスごとに分割(児=P、放=Mして各ファイルの縦合計を入力し、Excelファイルとして保存し、最後にまとめてZIPファイルにするスクリプトです。
import os
import shutil
import zipfile
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

class SplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("算定区分02: ファイル分割処理")
        self.root.geometry("400x120")
        self.root.attributes('-topmost', True)
        
        # 修正：os._exit(0)は親ごと終了してしまうので使用しない
        self.root.bind("<Escape>", lambda e: self.root.destroy())

        self.label = tk.Label(root, text="準備中...", font=("MS Gothic", 10), wraplength=350)
        self.label.pack(pady=15)
        self.progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=5)

    def start_process(self):
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
        target_path = filedialog.askopenfilename(
            title="converted_excels.zipを選択してください", 
            initialdir=downloads_path,
            filetypes=[("ZIP files", "*.zip")],
            parent=self.root
        )
        
        if target_path:
            # 【重要修正】os.path.splitext(target_path)[0] にしないとタプルになりエラーになります
            input_base_path = os.path.splitext(target_path)[0]
            temp_dir = input_base_path + "_temp"
            
            if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
            os.makedirs(temp_dir, exist_ok=True)
            self.run_task(temp_dir, input_base_path, target_path)
            
            # threading.Thread(target=self.run_task, args=(temp_dir, input_base_path, target_path), daemon=True).start()
        else:
            self.root.destroy()

    def run_task(self, temp_dir, input_base_path, target_path):
        try:
            with zipfile.ZipFile(target_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            files = [f for f in os.listdir(temp_dir) if f.lower().endswith((".csv", ".xlsx"))]
            self.progress["maximum"] = len(files)

            for i, filename in enumerate(files):
                self.label.config(text=f"分割・合計処理中: {filename}")
                self.root.update()
                file_path = os.path.join(temp_dir, filename)
                # 【重要修正】ここも[0]が必要
                base_name = os.path.splitext(filename)[0]
                
                if filename.lower().endswith(".csv"):
                    try: df = pd.read_csv(file_path, encoding='cp932')
                    except: df = pd.read_csv(file_path, encoding='utf-8-sig')
                else:
                    df = pd.read_excel(file_path)

                if not df.empty:
                    df = df[~df.iloc[:, 0].astype(str).str.contains(r"総計|\*集計", na=False)]
                    
                    if '利用サービス' in df.columns:
                        for prefix in ["放", "児"]:
                            sub_df = df[df['利用サービス'].astype(str).str.startswith(prefix)].copy()
                            if not sub_df.empty:
                                total_row = {}
                                for col in sub_df.columns:
                                    if col == sub_df.columns[0]: # 【修正】ここもindex指定に変更
                                        total_row[col] = "総計"
                                    else:
                                        num_val = pd.to_numeric(sub_df[col], errors='coerce')
                                        if not num_val.isna().all():
                                            total_row[col] = num_val.sum()
                                        else:
                                            total_row[col] = ""
                                
                                sub_df = pd.concat([sub_df, pd.DataFrame([total_row])], ignore_index=True)
                                sub_df.to_excel(os.path.join(temp_dir, f"{base_name}_{prefix}.xlsx"), index=False)
                
                os.remove(file_path)
                self.progress["value"] = i + 1

            output_zip_base = input_base_path + "_分割済み"
            shutil.make_archive(output_zip_base, 'zip', temp_dir)
            shutil.rmtree(temp_dir)
            os.remove(target_path)

            messagebox.showinfo("完了", "分割と合計処理が完了しました。", parent=self.root)
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました:\n{e}", parent=self.root)
        finally:
            self.root.destroy()

# --- 修正後の main 関数 ---
def main(root=None):
    if root is None:
        # 単体起動の場合
        root = tk.Tk()
        standalone = True
    else:
        # 親スクリプトから呼ばれた場合
        # 親の root を土台にして、新しいウィンドウ(Toplevel)を作る
        root = tk.Toplevel(root)
        standalone = False

    # ウィンドウの設定を確実に反映
    root.deiconify() 
    root.lift()
    root.focus_force()
    
    app = SplitterApp(root)
    
    # 起動後に処理を開始
    root.after(200, app.start_process)
    
    if standalone:
        root.mainloop()
    else:
        # 【重要】このウィンドウが閉じられる（destroy）まで、ここで待機する
        # これにより、親スクリプトは02が終わるまで03を呼びません
        root.wait_window()

if __name__ == "__main__":
    main()
   # main()