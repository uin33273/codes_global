#csvファイルをエクセルファイルに変換するプログラム
#csvのままでは列方向に使いにくいので、エクセルファイルに変換するプログラムを作成しました。
#変換するファイルは、downloads.zipの中に入っています。
import os
import zipfile
import pandas as pd
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import threading
import sys

# --- 1. クラスを関数の「外」に出す（NameError対策） ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("算定区分01: CSV to Excel 変換")
        self.root.geometry("400x120") # ボタンをなくしたので高さを詰めました
        self.root.attributes('-topmost', True)
        self.root.bind("<Escape>", self.force_quit)

        self.label = tk.Label(root, text="処理の準備中...", font=("MS Gothic", 11), wraplength=350)
        self.label.pack(pady=15)
        
        self.progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

    def force_quit(self, event=None):
        os._exit(0)

    def start_process(self):
        """起動直後に呼び出されるメイン処理"""
        downloads_path = Path.home() / "Downloads"
        zip_path = filedialog.askopenfilename(
            title="downloads.zipを選択してください",
            initialdir=str(downloads_path),
            filetypes=[("ZIPファイル", "*.zip")],
            parent=self.root
        )
        
        if zip_path:
            # 選択されたら即座にバックグラウンドで処理開始
            # 
            self.run_conversion(zip_path)
        else:
            # キャンセルされた場合は即座に終了して次へ
            self.root.destroy()

    def run_conversion(self, zip_path):
        downloads_path = Path.home() / "Downloads"
        output_dir = downloads_path / "excel_results"
        output_zip_base = downloads_path / "converted_excels"
        input_temp_dir = Path("extracted_files_temp")
        
        try:
            # ディレクトリ準備
            for d in [input_temp_dir, output_dir]:
                if d.exists(): shutil.rmtree(d)
                d.mkdir(parents=True, exist_ok=True)
            
            self.label.config(text="ファイルを解凍中...")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(input_temp_dir)
            
            csv_files = list(input_temp_dir.rglob('*.csv'))
            total = len(csv_files)
            
            if total == 0:
                messagebox.showwarning("警告", "CSVファイルが見つかりませんでした。", parent=self.root)
                self.root.destroy()
                return

            self.progress["maximum"] = total
            
            for i, file_path in enumerate(csv_files, 1):
                self.root.update()
                self.label.config(text=f"変換中 ({i}/{total}): {file_path.name}")
                
                # 文字コードを判定して読み込み
                df = None
                for enc in ['cp932', 'utf-8-sig', 'shift_jis']:
                    try:
                        df = pd.read_csv(file_path, encoding=enc, index_col=False)
                        break
                    except: continue
                
                if df is not None:
                    df.columns = [str(c).strip() for c in df.columns]
                    save_path = output_dir / f"{file_path.stem}.xlsx"
                    df.to_excel(save_path, index=False, engine='openpyxl')
                
                self.progress["value"] = i
            
            self.root.after(0, lambda: self.label.config(text="ファイルを解凍中..."))
            shutil.make_archive(str(output_zip_base), 'zip', str(output_dir))
            
            messagebox.showinfo("完了", "変換完了。次へ進みます。", parent=self.root)
            
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました:\n{e}", parent=self.root)
        finally:
            if input_temp_dir.exists(): shutil.rmtree(input_temp_dir)
            
            # 【重要】直接 destroy() せず、メインスレッドに「閉じて」と依頼する
            self.root.after(0, self.root.destroy)
            self.root.destroy() 

# --- 2. main関数（実行.pyから呼ばれる入口） ---
def main(root=None):
    # 親から root が渡されなかった場合（単体起動）だけ新しく作る
    if root is None:
        root = tk.Tk()
        standalone = True
    else:
        # 親がいる場合は、新しいウィンドウ(Toplevel)を作る
        root = tk.Toplevel(root)
        standalone = False

    app = App(root) 
    root.after(100, app.start_process)
    
    # 単体起動のときだけ mainloop を動かす
    if standalone:
        root.mainloop()
    else:
        # 【重要】このウィンドウが閉じられるまで、ここ（01）で処理を止める！
        # これを入れないと、スレッドを開始した直後に02へ進んでしまいます
        root.wait_window()

if __name__ == "__main__":
    main()