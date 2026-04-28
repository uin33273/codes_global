import sys
import os

def run_script_content(script_name):
    # ファイルの場所を特定
    if hasattr(sys, '_MEIPASS'):
        path = os.path.join(sys._MEIPASS, script_name)
    else:
        path = os.path.join(os.path.dirname(__file__), script_name)

    print(f"--- {script_name} を実行中 ---")
    
    if not os.path.exists(path):
        print(f"エラー: {path} が見つかりません。")
        return

    # ファイルの中身を読み込んで実行
    with open(path, encoding='utf-8') as f:
        code = f.read()
        # ここがポイント：__name__ を __main__ に設定して実行する
        exec(code, {'__name__': '__main__', '__file__': path})

if __name__ == "__main__":
    try:
        # 1. 結合を実行
        run_script_content("カウネット実績結合.py")
        
        # 2. 集計を実行
        run_script_content("カウネット集計.py")
        
        print("\nすべての処理が完了しました。")
    except Exception as e:
        print(f"\nエラーが発生しました:\n{e}")

    input("閉じるにはEnterキーを押してください...")
