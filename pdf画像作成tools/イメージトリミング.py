r"""
Auto-crop an image's whitespace margins to the limit and save the result.

- Dropping a single *.jpg file saves the cropped copy directly into the
  fixed OUTPUT_DIR folder.
- Dropping a folder processes every .jpg/.jpeg file inside it
  (non-recursive) and saves the cropped copies into a subfolder of
  OUTPUT_DIR named after the source folder (so results from different
  dropped folders don't mix together).

Two ways to use:

1) Drag & drop mode (no arguments): double-click this script (or run
   `python イメージトリミング.py`), leave the console window open, then
   drag a *.jpg file (or a folder containing *.jpg files) from Explorer
   onto the console window and press Enter. It crops + saves, then waits
   for the next file/folder. Type nothing and press Enter (or type exit)
   to quit.

2) One-shot mode: python イメージトリミング.py "path\to\image_or_folder" [threshold]

threshold (0-255, default 245): pixels lighter than this (per channel,
near-white) are treated as margin/background to be cropped away.
"""
import sys
import os
import re
import numpy as np
from PIL import Image

OUTPUT_DIR = r"C:\Users\owner\Downloads\トリミング済画像"
DEFAULT_THRESHOLD = 245


def autocrop(im, threshold=DEFAULT_THRESHOLD):
    rgb = im.convert("RGB")
    arr = np.array(rgb)
    # a pixel counts as "content" if any channel is darker than threshold
    mask = np.any(arr < threshold, axis=2)
    rows = np.where(mask.any(axis=1))[0]
    cols = np.where(mask.any(axis=0))[0]
    if rows.size == 0 or cols.size == 0:
        return im  # nothing but background found; return unchanged
    top, bottom = rows[0], rows[-1]
    left, right = cols[0], cols[-1]
    return im.crop((left, top, right + 1, bottom + 1))


def split_paths(line):
    """Split a line of one or more (optionally quoted) Windows paths
    without mangling backslashes the way shlex.split() would."""
    return [quoted or bare for quoted, bare in re.findall(r'"([^"]+)"|(\S+)', line)]


def process_one(path, threshold=DEFAULT_THRESHOLD, out_dir=None):
    path = path.strip().strip('"').strip("'")
    if not path:
        return
    if not os.path.isfile(path):
        print(f"  !! ファイルが見つかりません: {path}")
        return
    if not path.lower().endswith((".jpg", ".jpeg")):
        print(f"  !! .jpg / .jpeg ではありません: {path}")
        return

    im = Image.open(path)
    print(f"  元サイズ: {im.size}")

    cropped = autocrop(im, threshold)
    print(f"  トリミング後: {cropped.size}")

    out_dir = out_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, os.path.basename(path))
    cropped.convert("RGB").save(out_path, quality=95)
    print(f"  保存先: {out_path}")


def unique_dir(base_dir):
    """base_dir if it doesn't exist yet, otherwise base_dir(1), base_dir(2), ..."""
    if not os.path.exists(base_dir):
        return base_dir
    n = 1
    while True:
        candidate = f"{base_dir}({n})"
        if not os.path.exists(candidate):
            return candidate
        n += 1


def process_folder(folder, threshold=DEFAULT_THRESHOLD):
    names = sorted(
        name for name in os.listdir(folder)
        if name.lower().endswith((".jpg", ".jpeg"))
    )
    if not names:
        print(f"  !! フォルダ内に.jpg/.jpegが見つかりません: {folder}")
        return None
    out_dir = unique_dir(os.path.join(OUTPUT_DIR, f"{os.path.basename(os.path.normpath(folder))}_トリミング済"))
    print(f"  フォルダ内の画像 {len(names)} 件を処理します: {folder}")
    print(f"  保存先フォルダ: {out_dir}")
    for name in names:
        print(f"[{name}]")
        process_one(os.path.join(folder, name), threshold, out_dir=out_dir)
    return out_dir


def process_path(path, threshold=DEFAULT_THRESHOLD):
    path = path.strip().strip('"').strip("'")
    if not path:
        return None
    if os.path.isdir(path):
        return process_folder(path, threshold)
    elif os.path.isfile(path):
        process_one(path, threshold)
        return None
    else:
        print(f"  !! ファイル/フォルダが見つかりません: {path}")
        return None


def interactive_loop():
    print("=== JPEGトリミング (ドラッグ&ドロップ待機モード) ===")
    print(f"保存先フォルダ: {OUTPUT_DIR}")
    print("フォルダをドロップした場合はその中に、フォルダ名のサブフォルダを作成して保存します。")
    print(".jpg ファイル、またはそれらを含むフォルダをこのウィンドウにドラッグ&ドロップして Enter を押してください。")
    print("終了するには何も入力せず Enter、または exit と入力してください。")
    print()
    while True:
        try:
            line = input("> ").strip()
        except EOFError:
            break
        if not line or line.lower() in ("exit", "quit"):
            print("終了します。")
            break
        # Explorer can drop multiple files/folders on one line, each quoted separately
        folder_outputs = [d for d in (process_path(path) for path in split_paths(line)) if d]
        print("作業が終了しました。")
        if folder_outputs:
            for out_dir in folder_outputs:
                os.startfile(out_dir)
            return
        print()


def main():
    if len(sys.argv) > 1:
        threshold = int(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_THRESHOLD
        process_path(sys.argv[1], threshold)
        print()
        print("作業が終了しました。")
    else:
        interactive_loop()


if __name__ == "__main__":
    main()
