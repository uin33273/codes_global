import tkinter as tk
from tkinter import messagebox
import カウネット実績結合
import カウネット集計

def run_kekko():
    try:
        カウネット実績結合.main()
    except Exception as e:
        messagebox.showerror("エラー", f"実績結合でエラーが発生しました:\n{e}")

def run_shukei():
    try:
        カウネット集計.main()
    except Exception as e:
        messagebox.showerror("エラー", f"集計処理でエラーが発生しました:\n{e}")

def main():
    root = tk.Tk()
    root.title("カウネット処理ツール一式")
    root.geometry("400x320")

    # --- 最前面に表示する設定 ---
    root.attributes("-topmost", True)
    # --------------------------

    label = tk.Label(root, text="実行したいメニューを選択してください", font=("MS Gothic", 11))
    label.pack(pady=20)

    btn1 = tk.Button(root, text="1. 実績データのダウンロードと結合を実行", 
                     command=run_kekko, width=35, height=2, bg="#f0f0f0")
    btn1.pack(pady=5)

    btn2 = tk.Button(root, text="2. コピペ用データの集計を実行", 
                     command=run_shukei, width=35, height=2, bg="#f0f0f0")
    btn2.pack(pady=5)

    btn3 = tk.Button(root, text="3. アプリ終了", 
                     command=root.destroy, width=35, height=2, bg="#ffcccc")
    btn3.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()



