#csvファイルをエクセルファイルに変換するプログラム
#csvのままでは列方向に使いにくいので、エクセルファイルに変換するプログラムを作成しました。
#変換するファイルは、downloads.zipの中に入っています。
import os
import pandas as pd
import shutil
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path

# --- 地域振り分けの共通定義（02.py等からも import して使用） ---
REGION_MAP = {
    "01": "宇都宮",
    "02": "埼玉・群馬",
    "03": "茨城・千葉",
    "04": "その他",
    "05": "栃木①",
    "06": "栃木②",
}
TOCHIGI_CODES = {"01", "05", "06"}
TOCHIGI_GROUP_NAME = "2026.06 算定区分　栃木"
OTHER_GROUP_NAME = "2026.06 算定区分　栃木以外"


def _looks_like_source(p):
    """番号フォルダ(01~06)、または振り分け済みフォルダのどちらかがあればソースとみなす"""
    return (p / "01").exists() or (p / TOCHIGI_GROUP_NAME).exists() or (p / OTHER_GROUP_NAME).exists()


def find_source_root(downloads_path):
    """番号フォルダ(01~06)、または既に振り分け済みのフォルダが置かれている場所を探す"""
    # Downloads直下にあればそれを使う
    if _looks_like_source(downloads_path):
        return downloads_path
    base = downloads_path / "算定区分データ"
    if base.exists():
        # 算定区分データの直下にあればそれを使う
        if _looks_like_source(base):
            return base
        # 算定区分データ配下の年月フォルダ(例: 202606)を新しい順に探す
        for sub in sorted((p for p in base.iterdir() if p.is_dir()), reverse=True):
            if _looks_like_source(sub):
                return sub
    return downloads_path


def get_destination_dir(source_root, code):
    """地域コード(01~06)から振り分け先フォルダ(例: .../栃木/宇都宮/01)を得る"""
    group_name = TOCHIGI_GROUP_NAME if code in TOCHIGI_CODES else OTHER_GROUP_NAME
    return source_root / group_name / REGION_MAP[code] / code


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

    def create_organization_folders(self, source_root):
        """番号フォルダと同じ場所に振り分け用フォルダを作成する"""
        for code in REGION_MAP:
            get_destination_dir(source_root, code).parent.mkdir(parents=True, exist_ok=True)

    def move_numbered_folders(self, source_root):
        """ダウンロード済みの番号フォルダ(01~06)を振り分け先フォルダへ移動する"""
        for code in REGION_MAP:
            src = source_root / code
            dest_dir = get_destination_dir(source_root, code)
            if src.exists() and src.is_dir():
                try:
                    shutil.move(str(src), str(dest_dir))
                except Exception as e:
                    print(f"移動失敗: {code} - {e}")

    def start_process(self):
        """起動直後に呼び出されるメイン処理"""
        downloads_path = Path.home() / "Downloads"
        source_root = find_source_root(downloads_path)
        self.create_organization_folders(source_root)
        self.move_numbered_folders(source_root)
        self.convert_all_regions(source_root)

    def convert_all_regions(self, source_root):
        """振り分け済みの各地域フォルダ内のCSVを一括でExcelに変換する（手選択なし）"""
        try:
            region_dirs = [get_destination_dir(source_root, code) for code in REGION_MAP]
            region_dirs = [d for d in region_dirs if d.exists()]

            csv_files = []
            for d in region_dirs:
                csv_files.extend(d.rglob('*.csv'))

            total = len(csv_files)
            if total == 0:
                messagebox.showwarning("警告", "変換対象のCSVファイルが見つかりませんでした。", parent=self.root)
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
                    save_path = file_path.with_suffix('.xlsx')
                    df.to_excel(save_path, index=False, engine='openpyxl')
                    file_path.unlink()

                self.progress["value"] = i

            messagebox.showinfo("完了", "変換完了。次へ進みます。", parent=self.root)

        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました:\n{e}", parent=self.root)
        finally:
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