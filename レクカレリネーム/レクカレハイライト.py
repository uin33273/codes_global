#「レクカレ」(月次レクリエーションカレンダー画像)から「引き落し」の記載を自動検出してハイライトするツール
#対象は3つのカレンダーテンプレート(赤バナー / Global Kids Park の虹背景 / Global Kids method の水色背景)
#前の2つは「引き落し」が赤字、Global Kids method は黒字(費用：〜円（引き落し）)で表記されるため、
#赤系文字と黒系文字の両方の候補を検出し、それぞれ専用のテンプレート画像と照合する
#手順:
#  1) 同じ店舗(ファイル名末尾が0/1)は0のファイルだけを対象にする(1は重複のため無視)
#  2) 画像内の赤系文字ブロック・黒系文字ブロックをそれぞれ候補として検出する(ノイズも混ざる)
#  3) 候補ごとに、あらかじめ用意した「引き落し」の見本画像とテンプレートマッチングして判定する
#  4) 「引き落し」と判定された箇所だけを黄色マーカー+赤枠でハイライトし、別フォルダに保存する
#
#使い方: python レクカレハイライト.py を実行すると、フォルダ選択ダイアログが開く
import os
import re
import sys
import csv
import numpy as np
import cv2
import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image, ImageDraw

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATHS = [
    os.path.join(SCRIPT_DIR, "templates", "hikiotoshi_1.png"),
    os.path.join(SCRIPT_DIR, "templates", "hikiotoshi_2.png"),
]
# Global Kids method テンプレート用(「引き落し」が黒字で表記される)
DARK_TEMPLATE_PATHS = [
    os.path.join(SCRIPT_DIR, "templates", "hikiotoshi_dark_1.png"),
    os.path.join(SCRIPT_DIR, "templates", "hikiotoshi_dark_2.png"),
]

# 「引き落し」と判定するテンプレート一致度のしきい値（0〜1、検証の結果0.65〜0.75で誤り0件だったため中間値を採用）
MATCH_THRESHOLD = 0.7
# しきい値未満だが際どいスコアのものは見逃し防止のため別フォルダに集めて目視確認できるようにする
REVIEW_THRESHOLD = 0.5


def dedup_files(folder):
    """同じ店舗（ファイル名末尾の0/1を除いた部分が同じ）は、0の方を優先して1ファイルだけ残す"""
    files = [f for f in os.listdir(folder) if f.lower().endswith(".jpg")]
    groups = {}
    for f in files:
        base = re.sub(r"[01]\.jpg$", "", f)
        groups.setdefault(base, []).append(f)
    picked = []
    for base, fs in groups.items():
        zero = [f for f in fs if f.endswith("0.jpg")]
        picked.append(zero[0] if zero else sorted(fs)[0])
    return sorted(picked)


def _merge_close_boxes(boxes, v_gap_max=12, h_overlap_min=5):
    """縦に近接する行（折り返し2行の金額表記など）を1つの枠にまとめる"""
    boxes = list(boxes)
    changed = True
    while changed:
        changed = False
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                x1, y1, w1, h1 = boxes[i]
                x2, y2, w2, h2 = boxes[j]
                h_overlap = min(x1 + w1, x2 + w2) - max(x1, x2)
                v_gap = max(y1, y2) - min(y1 + h1, y2 + h2)
                if h_overlap > h_overlap_min and v_gap < v_gap_max:
                    nx0, ny0 = min(x1, x2), min(y1, y2)
                    nx1, ny1 = max(x1 + w1, x2 + w2), max(y1 + h1, y2 + h2)
                    boxes[i] = (nx0, ny0, nx1 - nx0, ny1 - ny0)
                    del boxes[j]
                    changed = True
                    break
            if changed:
                break
    return boxes


def find_red_text_lines(arr, y_min=255, y_max=820, row_thresh=3, gap_rows=3,
                         min_line_height=4, max_line_height=16, col_gap=10,
                         min_width=8, max_width=65):
    """画像内の赤系文字ブロックの候補（バウンディングボックス）を検出する。

    以前は行単位の赤ピクセル数を横一列ぶん積算してから空白行(gap_rows)で
    区切っていたが、店舗によっては右上のマスコットのイラストが縦に長く
    赤系の色を含み、実際の金額テキストと(隙間なく)縦につながって1つの
    巨大なブロックとみなされ、max_line_height を超えて丸ごと検出漏れに
    なるケースがあった(例: 上越店6/13・6/27, 新西茂呂店6/13・6/20など)。
    connectedComponents を使い、横方向だけを軽く膨張させて縦のにじみを
    抑えることでこの誤結合を防ぐ。折り返した2行の金額表記は縦に近い
    2つの行として検出されるため、_merge_close_boxes で1つの枠にまとめる。
    """
    region = arr[y_min:y_max, :, :]
    r = region[:, :, 0].astype(int)
    g = region[:, :, 1].astype(int)
    b = region[:, :, 2].astype(int)
    mask = ((r > 130) & (r - g > 50) & (r - b > 50)).astype(np.uint8) * 255

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (11, 1))
    merged_mask = cv2.dilate(mask, kernel, iterations=1)
    n, labels, stats, centroids = cv2.connectedComponentsWithStats(merged_mask, connectivity=8)

    boxes = []
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if w / h < 3.0:
            continue
        if not (min_line_height <= h <= max_line_height + 4):
            continue
        if not (min_width <= w <= max_width * 2.5):
            continue
        boxes.append((int(x), int(y), int(w), int(h)))

    boxes = _merge_close_boxes(boxes)

    results = []
    for (x, y, w, h) in boxes:
        width = w
        if width < min_width or width > max_width:
            continue
        results.append((x, y_min + y, x + w, y_min + y + h))
    return results


def find_dark_text_lines(arr, y_min=180, y_max=780, min_line_height=4, max_line_height=13,
                          min_width=15, max_width=260):
    """画像内の黒系文字ブロックの候補（バウンディングボックス）を検出する(Global Kids method テンプレート用)。

    赤バナー系と違い「引き落し」が黒字の本文中に出てくるため、候補は本文の行単位で
    広めに取る(「費用：100円（材料費として）（引き落し）」のように前後にテキストが
    連なっていても、テンプレートマッチングは候補領域内の最良一致箇所を探すため問題ない)。
    見出し等の水色文字はグレースケール変換後の輝度が高く出るため、閾値により自然に除外される。
    """
    region = arr[y_min:y_max, :, :]
    gray = cv2.cvtColor(region, cv2.COLOR_RGB2GRAY)
    mask = (gray < 120).astype(np.uint8) * 255

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 1))
    merged_mask = cv2.dilate(mask, kernel, iterations=1)
    n, labels, stats, centroids = cv2.connectedComponentsWithStats(merged_mask, connectivity=8)

    results = []
    for i in range(1, n):
        x, y, w, h, area = stats[i]
        if w / max(h, 1) < 2.5:
            continue
        if not (min_line_height <= h <= max_line_height):
            continue
        if not (min_width <= w <= max_width):
            continue
        results.append((int(x), y_min + int(y), int(x + w), y_min + int(y + h)))
    return results


def load_templates(paths):
    templates = []
    for path in paths:
        img = Image.open(path).convert("RGB")
        gray = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2GRAY)
        templates.append(gray)
    return templates


def match_score(gray_crop, templates):
    """候補領域が「引き落し」テンプレートとどれだけ一致するか（0〜1、複数テンプレート・複数スケールの最大値）を返す"""
    best = -1.0
    for tmpl in templates:
        th, tw = tmpl.shape
        for scale in (0.9, 0.95, 1.0, 1.05, 1.1):
            sw, sh = max(1, int(tw * scale)), max(1, int(th * scale))
            if gray_crop.shape[0] < sh or gray_crop.shape[1] < sw:
                continue
            resized = cv2.resize(tmpl, (sw, sh))
            res = cv2.matchTemplate(gray_crop, resized, cv2.TM_CCOEFF_NORMED)
            m = res.max()
            if m > best:
                best = m
    return best


def show_progress_dialog(root, total):
    dialog = tk.Toplevel(root)
    dialog.title("ハイライト加工中")
    dialog.geometry("360x110")
    dialog.resizable(False, False)
    dialog.protocol("WM_DELETE_WINDOW", lambda: None)  # 完了まで閉じさせない

    label_var = tk.StringVar(value=f"0 / {total} 件")
    tk.Label(dialog, textvariable=label_var).pack(pady=(15, 5))

    bar = ttk.Progressbar(dialog, orient="horizontal", length=300, mode="determinate", maximum=max(total, 1))
    bar.pack(pady=5)

    dialog.lift()
    dialog.attributes("-topmost", True)
    dialog.update_idletasks()

    def update(step, step_total):
        bar["maximum"] = max(step_total, 1)
        bar["value"] = step
        label_var.set(f"{step} / {step_total} 件")
        dialog.update_idletasks()
        dialog.update()

    return dialog, update


def process_folder(folder, progress_callback=None):
    folder_name = os.path.basename(folder.rstrip("\\/"))
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    out_dir = os.path.join(downloads_dir, folder_name + "_highlight")
    os.makedirs(out_dir, exist_ok=True)
    review_dir = os.path.join(out_dir, "_要確認")
    os.makedirs(review_dir, exist_ok=True)
    no_check_dir = os.path.join(out_dir, "チェック不要")
    os.makedirs(no_check_dir, exist_ok=True)

    templates = load_templates(TEMPLATE_PATHS)
    dark_templates = load_templates(DARK_TEMPLATE_PATHS)
    files = dedup_files(folder)
    print(f"対象ファイル数(重複除外後): {len(files)}")

    report_rows = []
    hit_file_count = 0
    hit_box_count = 0
    review_count = 0
    no_check_count = 0

    for fi, filename in enumerate(files):
        path = os.path.join(folder, filename)
        try:
            img = Image.open(path).convert("RGB")
        except Exception as e:
            print(f"  読み込み失敗: {filename} ({e})")
            continue
        arr = np.array(img)
        candidates = [(x0, y0, x1, y1, templates) for (x0, y0, x1, y1) in find_red_text_lines(arr)]
        candidates += [(x0, y0, x1, y1, dark_templates) for (x0, y0, x1, y1) in find_dark_text_lines(arr)]

        hit_boxes = []
        review_boxes = []
        for (x0, y0, x1, y1, tmpl_set) in candidates:
            pad = 3
            crop = img.crop((max(0, x0 - pad), max(0, y0 - pad), x1 + pad, y1 + pad))
            gray_crop = cv2.cvtColor(np.array(crop), cv2.COLOR_RGB2GRAY)
            score = match_score(gray_crop, tmpl_set)
            if score >= MATCH_THRESHOLD:
                hit_boxes.append((x0, y0, x1, y1, score))
            elif score >= REVIEW_THRESHOLD:
                review_boxes.append((x0, y0, x1, y1, score))

        # hit / 要確認 / チェック不要 のいずれか1箇所にのみ振り分ける(出力ファイル数を処理対象ファイル数と一致させるため)
        if hit_boxes:
            overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(overlay)
            pad = 4
            for (x0, y0, x1, y1, score) in hit_boxes:
                box = (x0 - pad, y0 - pad, x1 + pad, y1 + pad)
                draw.rectangle(box, fill=(255, 255, 0, 120))
                draw.rectangle(box, outline=(255, 0, 0, 255), width=2)
                report_rows.append([filename, x0, y0, x1, y1, f"{score:.3f}", "hit"])
            result = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
            result.save(os.path.join(out_dir, filename), quality=95)
            hit_file_count += 1
            hit_box_count += len(hit_boxes)
        elif review_boxes:
            overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(overlay)
            pad = 4
            for (x0, y0, x1, y1, score) in review_boxes:
                box = (x0 - pad, y0 - pad, x1 + pad, y1 + pad)
                draw.rectangle(box, outline=(0, 120, 255, 255), width=2)
                report_rows.append([filename, x0, y0, x1, y1, f"{score:.3f}", "review"])
            result = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
            result.save(os.path.join(review_dir, filename), quality=95)
            review_count += 1
        else:
            img.save(os.path.join(no_check_dir, filename), quality=95)
            no_check_count += 1

        if progress_callback:
            progress_callback(fi + 1, len(files))
        if (fi + 1) % 50 == 0:
            print(f"  {fi + 1}/{len(files)} 件処理済み")

    report_path = os.path.join(out_dir, "ハイライト結果一覧.csv")
    with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["ファイル名", "x0", "y0", "x1", "y1", "一致度", "判定"])
        w.writerows(report_rows)

    print()
    print("=== 完了 ===")
    print(f"ハイライトを付与したファイル数: {hit_file_count}")
    print(f"ハイライト箇所の合計: {hit_box_count}")
    print(f"要確認(際どいスコア)のファイル数: {review_count} → {review_dir}")
    print(f"チェック不要のファイル数: {no_check_count} → {no_check_dir}")
    print(f"出力フォルダ: {out_dir}")
    print(f"詳細一覧: {report_path}")

    return out_dir


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="引き落しハイライト加工するフォルダを選択してください。")
    if not folder:
        print("フォルダが選択されなかったため終了します。")
        sys.exit(1)
    progress_dialog, update_progress = show_progress_dialog(root, 1)
    out_dir = process_folder(folder, progress_callback=update_progress)
    progress_dialog.destroy()
    os.startfile(out_dir)
