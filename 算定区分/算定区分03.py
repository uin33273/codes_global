#各店舗のExcelデータを1つのファイルにまとめるツール
import pandas as pd
import glob
import os
import zipfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import threading

# --- クラスを外に出して定義 ---
class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("算定区分03: Excelデータ統合")
        self.root.geometry("450x180")
        self.root.attributes('-topmost', True)
        self.root.bind("<Escape>", lambda e: self.root.destroy())

        self.label = tk.Label(root, text="準備中...\n(Escキーで強制終了)", wraplength=400)
        self.label.pack(pady=15)
        self.progress = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
        self.progress.pack(pady=10)
        self.start_btn = tk.Button(root, text="ファイルを選択して開始", command=self.start_process)
        self.start_btn.pack(pady=10)

    def force_quit(self, event=None):
        os._exit(0)

    # 実行.pyから呼ばれる開始メソッド名を「start_process」に統一
    def start_process(self):
        downloads_path = Path.home() / "Downloads"
        target_files = filedialog.askopenfilenames(
            title="分割済みZIPまたはExcelを選択してください",
            initialdir=str(downloads_path),
            filetypes=[("Excel/ZIP files", "*.xlsx *.zip")],
            parent=self.root # 親ウィンドウ(03の窓)の上に表示
        )
        
        if not target_files:
            # キャンセルされたらこのウィンドウを閉じて、司令塔(実行.py)を次の04へ進める
            self.root.destroy()
            return
            
        self.start_btn.config(state="disabled")
        # threading.Thread(target=self.process_files, args=(target_files,), daemon=True).start()
        self.process_files(target_files)

    def to_num(self, val):
        try:
            return float(str(val).replace(',', '')) if pd.notna(val) and str(val).strip() != '' else 0
        except:
            return 0

    def process_files(self, target_files):
        downloads_path = Path.home() / "Downloads"
        output_file_path = downloads_path / 'anyfiles_to_1files.xlsx'
        target_dir = Path('excel_data_temp')
        try:
            if target_dir.exists(): shutil.rmtree(target_dir)
            target_dir.mkdir(exist_ok=True)

            self.label.config(text="ファイルを展開中...")
            for file_path in target_files:
                p = Path(file_path)
                if p.suffix.lower() == '.zip':
                    with zipfile.ZipFile(file_path, 'r') as zip_ref:
                        zip_ref.extractall(target_dir)
                elif p.suffix.lower() == '.xlsx':
                    shutil.copy(file_path, target_dir / p.name)

            excel_files = [f for f in target_dir.rglob('*.xlsx') if not Path(f).name.startswith('~$')]
            total = len(excel_files)
            if total == 0:
                messagebox.showwarning("警告", "Excelファイルが見つかりませんでした。", parent=self.root)
                self.root.destroy()
                return

            self.progress["maximum"] = total
            combined_data = []

            for i, file in enumerate(excel_files, 1):
                clean_filename = Path(file).stem.replace('店', '')
                self.label.config(text=f"読み込み中 ({i}/{total}):\n{clean_filename}")
                self.progress["value"] = i
                self.root.update()
                try:
                    df = pd.read_excel(file, index_col=None)
                    if df.empty: continue
                    df.columns = [str(c).strip() for c in df.columns]
                    total_row = df.iloc[-1]
                    combined_data.append({
                        'Folder': Path(file).parent.name,
                        'filename': clean_filename,
                        'J_稼働数': total_row.get("利用日数", 0),
                        'S_算定区分1': total_row.get("算定区分１（日数）", 0),
                        'U_算定区分2': total_row.get("算定区分２（日数）", 0),
                        'W_算定区分3': total_row.get("算定区分３（日数）", 0),
                        'ce+cg+ci_延長支援': self.to_num(total_row.get("支援費合計内訳(延長支援区分１日数)", 0)) + 
                                        self.to_num(total_row.get("支援費合計内訳(延長支援区分２日数)", 0)) + 
                                        self.to_num(total_row.get("支援費合計内訳(延長支援区分３日数)", 0)),
                        'BG_専門実施': total_row.get("支援費合計内訳(専門的支援実施加算日数)", 0),
                        'CC_送迎加算': total_row.get("支援費合計内訳(送迎加算日数)", 0),
                        'AS_欠席加算': total_row.get("支援費合計内訳(欠席時対応加算日数)", 0)
                    })
                except Exception as e:
                    print(f"エラースキップ: {file.name} - {e}")

            if combined_data:
                self.label.config(text="Excelファイルを作成中...")
                pd.DataFrame(combined_data).to_excel(output_file_path, index=False)
                messagebox.showinfo("完了", "統合完了。紐づけ処理へ進みます。", parent=self.root)
            
        except Exception as e:
            messagebox.showerror("エラー", f"失敗しました:\n{e}", parent=self.root)
        finally:
            if target_dir.exists(): shutil.rmtree(target_dir)
            # 完了時にウィンドウを壊すことで実行.pyが次(04)へ進む
            self.root.destroy()

# --- 実行用関数（単体テスト用） ---
def main(root=None):
    if root is None:
        # 単体起動の場合
        root = tk.Tk()
        standalone = True
    else:
        # 親（実行.py）から呼ばれた場合、子ウィンドウとして作成
        root = tk.Toplevel(root)
        standalone = False

    # ウィンドウを確実に前面に出す
    root.deiconify()
    root.lift()
    root.focus_force()

    app = ExcelMergerApp(root)
    
    # 起動直後にファイル選択を開始
    root.after(100, app.start_process)
    
    if standalone:
        root.mainloop()
    else:
        # この03のウィンドウが閉じられるまで、親プログラムをここで待機させる
        # これにより04が同時に立ち上がるのを防ぎます
        root.wait_window()

if __name__ == "__main__":
    main()