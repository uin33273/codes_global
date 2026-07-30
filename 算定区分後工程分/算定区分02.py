#サーバーからダウンロードしたZIPファイルや、ローカルのフォルダ内のCSV/Excelファイルの総計行を削除、利用サービスごとに分割(児=P、放=Mして各ファイルの縦合計を入力し、Excelファイルとして保存し、最後にまとめてZIPファイルにするスクリプトです。
import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path

import 算定区分01

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
        downloads_path = Path(os.path.expanduser("~")) / "Downloads"
        source_root = 算定区分01.find_source_root(downloads_path)
        self.split_all_regions(source_root)

    def split_all_regions(self, source_root):
        """振り分け済みの各地域フォルダ内のExcelを一括で_放/_児に分割し、地域ごとにdownload.zipを作る（手選択なし）"""
        try:
            region_dirs = [算定区分01.get_destination_dir(source_root, code) for code in 算定区分01.REGION_MAP]
            region_dirs = [d for d in region_dirs if d.exists()]

            # 分割対象＝まだ _放/_児 になっていないExcelファイル
            def scan_target_files():
                files = []
                for d in region_dirs:
                    for f in d.glob('*.xlsx'):
                        if f.name.startswith('~$'): continue
                        if f.stem.endswith('_放') or f.stem.endswith('_児'): continue
                        files.append((d, f))
                return files

            def count_already_split():
                return sum(
                    1 for d in region_dirs for f in d.glob('*.xlsx')
                    if not f.name.startswith('~$') and (f.stem.endswith('_放') or f.stem.endswith('_児'))
                )

            target_files = scan_target_files()

            total = len(target_files)
            if total == 0:
                already_split_count = count_already_split()
                if region_dirs and already_split_count > 0:
                    # 新旧データが混在しないよう、既存ファイルを使い回すのではなく、
                    # 01のCSV→Excel変換を強制的にやり直して最新化する。
                    # (CSV→Excelは同名ファイルに上書き保存されるため、削除しなくても
                    #  最新のCSV内容に揃う。このデータは店舗ごとに_児/_放が別CSVとして
                    #  ダウンロードされる仕様なので、変換後もファイル名はそのまま_児/_放になる)
                    self.label.config(text="01を再実行して最新のCSVに更新しています...")
                    self.root.update()
                    converted = 算定区分01.convert_csv_files(region_dirs)

                    # 再変換後に改めてスキャンし直す。結合レポート(まだ_放/_児に分かれて
                    # いないCSV)があれば、そこから新たな分割対象が見つかる
                    target_files = scan_target_files()
                    total = len(target_files)

                    if total == 0:
                        # 追加で分割が必要な結合レポートは無かった(=このデータは元々
                        # 店舗ごとに_児/_放が分かれているため、02独自の分割処理は不要)。
                        # 常に最新のExcel群でZIPを作り直してから次(03)へ進む
                        downloads_path = Path(os.path.expanduser("~")) / "Downloads"
                        bundle_zip_base = downloads_path / "converted_excels_分割済み"
                        shutil.make_archive(str(bundle_zip_base), 'zip', str(source_root))

                        note = (
                            f"01を再実行し、最新のCSVから{converted}件のExcelを作り直しました。"
                            if converted > 0
                            else "CSVが見つからなかったため、既存のExcelをそのままZIP化しました。"
                        )
                        messagebox.showinfo(
                            "分割対象なし(既に単体ファイル)",
                            "追加で分割が必要なExcelは見つかりませんでした。\n\n"
                            f"振り分け先フォルダには_放/_児のExcelが{already_split_count}件あります。\n"
                            f"{note}\n"
                            "そのまま次(03)へ進みます。",
                            parent=self.root,
                        )
                        self.root.destroy()
                        return
                    # else: 結合レポートが見つかったので、下の通常の分割処理へ進む
                else:
                    messagebox.showwarning(
                        "分割対象のExcelが見つかりません",
                        "分割対象のExcelファイルが見つかりませんでした。\n\n"
                        "・算定区分01(CSV→Excel変換)が完了しているか\n"
                        "をご確認ください。",
                        parent=self.root,
                    )
                    self.root.destroy()
                    return

            self.progress["maximum"] = total

            for i, (dest_dir, file_path) in enumerate(target_files, 1):
                self.label.config(text=f"分割・合計処理中 ({i}/{total}): {file_path.name}")
                self.root.update()
                base_name = file_path.stem

                df = pd.read_excel(file_path)
                split_written = False

                if not df.empty:
                    df = df[~df.iloc[:, 0].astype(str).str.contains(r"総計|\*集計", na=False)]

                    if '利用サービス' in df.columns:
                        for prefix in ["放", "児"]:
                            sub_df = df[df['利用サービス'].astype(str).str.startswith(prefix)].copy()
                            if not sub_df.empty:
                                total_row = {}
                                for col in sub_df.columns:
                                    if col == sub_df.columns[0]:
                                        total_row[col] = "総計"
                                    else:
                                        num_val = pd.to_numeric(sub_df[col], errors='coerce')
                                        if not num_val.isna().all():
                                            total_row[col] = num_val.sum()
                                        else:
                                            total_row[col] = ""

                                sub_df = pd.concat([sub_df, pd.DataFrame([total_row])], ignore_index=True)
                                sub_df.to_excel(dest_dir / f"{base_name}_{prefix}.xlsx", index=False)
                                split_written = True

                # 振り分け後に必要なのは分割済み(_放/_児)のExcelだけなので、分割元の
                # 統合Excelと、同名のCSV(振り分け先へコピーされた分)はここで削除する。
                # 生データ自体は archive_raw_folders で {年月}生データ に退避済みなので
                # ここで消しても元データが失われることはない
                if split_written:
                    try:
                        file_path.unlink()
                    except OSError:
                        pass
                    csv_path = dest_dir / f"{base_name}.csv"
                    if csv_path.exists():
                        try:
                            csv_path.unlink()
                        except OSError:
                            pass

                self.progress["value"] = i

            # 202606フォルダ全体をまとめて、03.pyが読み込むconverted_excels_分割済み.zipを作成
            downloads_path = Path(os.path.expanduser("~")) / "Downloads"
            bundle_zip_base = downloads_path / "converted_excels_分割済み"
            shutil.make_archive(str(bundle_zip_base), 'zip', str(source_root))

            messagebox.showinfo("完了", f"分割と振り分けが完了しました。\n\n作成したZIP:\n{bundle_zip_base}.zip", parent=self.root)
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