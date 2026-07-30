#csvファイルをエクセルファイルに変換するプログラム
#csvのままでは列方向に使いにくいので、エクセルファイルに変換するプログラムを作成しました。
#変換するファイルは、downloads.zipの中に入っています。
import os
import re
import pandas as pd
import shutil
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path
from datetime import datetime

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


def get_group_names(source_root):
    """source_root(年月フォルダ、例: '202606')の名前から '2026.06 算定区分　栃木' 等の
    振り分け先フォルダ名を作る。年月が読み取れない場合は今日の年月を使う。
    (以前はこの文字列が '2026.06' 固定だったため、他の月のデータでもフォルダ名が
    「2026.06」のままになる不具合があった)"""
    m = re.fullmatch(r"(\d{4})(\d{2})", source_root.name)
    if m:
        label = f"{m.group(1)}.{m.group(2)}"
    else:
        today = datetime.now()
        label = f"{today.year}.{today.month:02d}"
    return f"{label} 算定区分　栃木", f"{label} 算定区分　栃木以外"


def _looks_like_source(p):
    """番号フォルダ(01~06)、または振り分け済みフォルダのどちらかがあればソースとみなす"""
    tochigi_name, other_name = get_group_names(p)
    return (p / "01").exists() or (p / tochigi_name).exists() or (p / other_name).exists()


def find_source_root(downloads_path):
    """番号フォルダ(01~06)、または既に振り分け済みのフォルダが置かれている場所を探す。
    算定区分CSVダウンロード.py の保存先(Downloads\\算定区分データ\\年月)を優先的に探し、
    そこに何も無い場合だけ Downloads 直下を見る。
    (以前は逆順だったため、Downloads直下に古い/空の振り分けフォルダが残っていると、
    算定区分データ配下に本物のデータがあってもそちらを見つけられない不具合があった)"""
    base = downloads_path / "算定区分データ"
    if base.exists():
        # 算定区分データの直下にあればそれを使う
        if _looks_like_source(base):
            return base
        # 算定区分データ配下の「年月」という名前のフォルダ(例: 202606)だけを新しい順に探す。
        # (フォルダ名を年月形式に限定することで、「バック」や「{年月}生データ」など
        #  無関係な名前のフォルダを誤ってソースとして拾ってしまう不具合を防ぐ)
        for sub in sorted(
            (p for p in base.iterdir() if p.is_dir() and re.fullmatch(r"\d{4}\d{2}", p.name)),
            reverse=True,
        ):
            if _looks_like_source(sub):
                return sub
    # 算定区分データ配下に見つからない場合のみ、Downloads直下を見る(互換用)
    if _looks_like_source(downloads_path):
        return downloads_path
    # どこにもソースが見つからない場合の既定値は「算定区分データ」にする。
    # (以前はここでdownloads_path自体を返してしまい、算定区分データフォルダが
    #  無い状態で実行すると、Downloads直下に誤って振り分け先フォルダが
    #  作られてしまう不具合があった)
    return base


def get_destination_dir(source_root, code):
    """地域コード(01~06)から振り分け先フォルダ(例: .../栃木/宇都宮/01)を得る"""
    tochigi_name, other_name = get_group_names(source_root)
    group_name = tochigi_name if code in TOCHIGI_CODES else other_name
    return source_root / group_name / REGION_MAP[code] / code


def get_yyyymm(source_root):
    """source_root(例: '202606')から年月文字列を得る。読み取れない場合は今日の年月を使う。"""
    m = re.fullmatch(r"(\d{4})(\d{2})", source_root.name)
    if m:
        return source_root.name
    today = datetime.now()
    return f"{today.year}{today.month:02d}"


def archive_raw_folders(source_root):
    """栃木/栃木以外への振り分け(コピー)が終わった後、ダウンロード直後の番号フォルダ
    (01~06)を Downloads\\算定区分データ\\{年月}生データ\\ へフォルダごと移動して片付ける。
    振り分け済みデータ(栃木/栃木以外の下)はコピー済みなので影響しない。
    移動はコピー→元削除の順で行い、生データフォルダに既に前回分がある場合も
    (dirs_exist_ok=True で)中身を壊さずマージする。"""
    raw_root = source_root.parent / f"{get_yyyymm(source_root)}生データ"

    for code in REGION_MAP:
        src = source_root / code
        if src.exists() and src.is_dir():
            dest = raw_root / code
            try:
                shutil.copytree(str(src), str(dest), dirs_exist_ok=True)
                shutil.rmtree(str(src))
            except Exception as e:
                print(f"生データ移動失敗: {code} - {e}")


def convert_csv_files(region_dirs, progress_callback=None):
    """GUI無しでCSV→Excel変換だけを行う(02.pyから「強制やり直し」時に呼び出すための版)。
    戻り値は実際に変換できた件数。"""
    csv_files = []
    for d in region_dirs:
        csv_files.extend(d.rglob('*.csv'))

    converted = 0
    total = len(csv_files)
    for i, file_path in enumerate(csv_files, 1):
        df = None
        for enc in ['cp932', 'utf-8-sig', 'shift_jis']:
            try:
                df = pd.read_csv(file_path, encoding=enc, index_col=False)
                break
            except Exception:
                continue

        if df is not None:
            df.columns = [str(c).strip() for c in df.columns]
            save_path = file_path.with_suffix('.xlsx')
            df.to_excel(save_path, index=False, engine='openpyxl')
            converted += 1

        if progress_callback:
            progress_callback(i, total, file_path.name)

    return converted


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

    def copy_numbered_folders(self, source_root):
        """ダウンロード済みの番号フォルダ(01~06)を振り分け先フォルダへコピーする。
        ダウンロード元の番号フォルダ自体は移動・削除せず、そのまま残す
        (誤って振り分け先が消えても、ここから何度でもやり直せるようにするため)"""
        for code in REGION_MAP:
            src = source_root / code
            dest_dir = get_destination_dir(source_root, code)
            if src.exists() and src.is_dir():
                try:
                    shutil.copytree(str(src), str(dest_dir), dirs_exist_ok=True)
                except Exception as e:
                    print(f"コピー失敗: {code} - {e}")

    def start_process(self):
        """起動直後に呼び出されるメイン処理"""
        downloads_path = Path.home() / "Downloads"
        source_root = find_source_root(downloads_path)
        self.create_organization_folders(source_root)
        self.copy_numbered_folders(source_root)
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
                xlsx_count = sum(len(list(d.rglob('*.xlsx'))) for d in region_dirs)
                if region_dirs and xlsx_count > 0:
                    # 振り分け先フォルダは既に存在し、xlsxもある = 前回既に変換済みという意味
                    messagebox.showinfo(
                        "変換対象なし(処理済みの可能性)",
                        "変換対象のCSVが見つかりませんでした。\n\n"
                        f"振り分け先フォルダには既にExcelファイルが{xlsx_count}件あるため、\n"
                        "前回の実行で既に変換済みの可能性があります。\n"
                        "そのまま次(02)へ進みます。",
                        parent=self.root,
                    )
                    self.root.destroy()
                    return
                else:
                    messagebox.showwarning(
                        "変換対象のCSVが見つかりません",
                        "変換対象のCSVファイルが見つかりませんでした。\n\n"
                        "・HUGからのダウンロードが完了しているか\n"
                        "・ダウンロード先フォルダ(Downloads\\算定区分データ)が正しいか\n"
                        "をご確認ください。",
                        parent=self.root,
                    )
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
                    # 元CSVは削除せずそのまま残す(再実行時にやり直せるようにするため)

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