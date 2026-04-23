#メインループ
#このコードは、ユーザーに対して一連の処理を実行する前に確認を求め、複数のPythonスクリプトを順番に実行し、完了後にメッセージを表示するためのものです。escapeキーで終了することもできます。
import tkinter as tk
from tkinter import messagebox
import sys
import os
import shutil
from pathlib import Path

# 1. 各スクリプトをモジュールとしてインポート
import 算定区分01
import 算定区分02
import 算定区分03
import 算定区分04
import 算定区分05

def exit_program(event=None):
    """Escキーが押された時の処理"""
    if messagebox.askyesno("中断", "プログラム全体を終了しますか？"):
        sys.exit()

def run():
    root = tk.Tk()
    root.withdraw()

    # Escキーを監視
    # root.bind('<Escape>', exit_program)
    
    #from tkinter import messagebox
    confirm_msg = (
        "【Excel版】算定区分・加算等集計表（運営用）は準備済ですか？\n\n"
        "「はい」を選択すると処理を開始します。\n"
        "「いいえ」を選択すると終了します。"
    )
    
    if messagebox.askyesno("確認", confirm_msg):
        try:
            # 2. 各ファイルの main() 関数を順番に実行
            # これにより PyInstaller は自動的にこれらのファイルを EXE に含めます
            
            #print("--- 01開始 ---")
            算定区分01.main(root)
            #print("--- 01終了 / 02を呼び出します ---") # 
            算定区分02.main(root)
            #print("--- 02終了 / 03を呼び出します ---")
            算定区分03.main(root)
            算定区分04.main(root)
            算定区分05.main(root)
                
            messagebox.showinfo("完了", "すべての処理が正常に終了しました。")

            # --- フォルダ削除機能の追加 ---
            downloads = Path.home() / "Downloads"
            # 削除対象のリスト（Downloads配下と、実行ファイルと同階層の両方をチェック）
            cleanup_targets = [
                downloads / "excel_results",
                downloads / "downroads_temp",
                Path("excel_results"),
                Path("downroads_temp")
            ]
            for target in cleanup_targets:
                if target.exists() and target.is_dir():
                    try:
                        shutil.rmtree(target)
                        print(f"削除成功: {target}")
                    except Exception as e:
                        print(f"削除失敗: {target} - {e}")
                        
        except Exception:
            # エラーの詳細（どこが間違っているか）を画面に出す
            import traceback
            error_details = traceback.format_exc()
            print(error_details) # コンソールに表示
            messagebox.showerror("エラー詳細", error_details) # 画面に表示

            # except Exception as e:
            #     # 各スクリプト内で root.destroy() された際のエラーや予期せぬエラーをキャッチ
            #     messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
            
    #root.destroy()

if __name__ == "__main__":
    run()
    #pyinstaller 算定区分処理実行.py --onefile --noconsole
    #cd C:\Users\owner\Documents\GitHub\global_python