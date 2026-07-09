import os
import re
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl

# 店舗リストのExcelファイル(固定パス)
LIST_PATH = r"C:\Users\owner\Desktop\works\保管書類\店舗リスト\table_HHHL共通店舗一覧m.xlsm"
LIST_SHEET = "リスト"
NAME_COL = "店舗名"   # 例: 001GH宇都宮
AREA_COL = "区分"     # エリアとして扱う列

FRONT_PATTERN = re.compile(r'^(\d+)(.*)$')
CODE_PREFIX_PATTERN = re.compile(r'^[A-Za-z]+')  # 店舗名先頭のコード(GH/GP/KM等)


class GoBackToFolderSelection(Exception):
    """「戻る」ボタンが押されたときに、フォルダ選択からやり直すための合図"""
    pass


def normalize_text(s):
    """全角半角ゆれを吸収し、写真ファイル名で付きがちな「店」を無視する"""
    s = unicodedata.normalize("NFKC", s)
    return s.replace("店", "")


def core_place_name(rest):
    """店舗名からコード(GH/GP/KM等の先頭アルファベット)を除いた地名部分を取り出す"""
    return CODE_PREFIX_PATTERN.sub("", rest)


def bring_to_front(win):
    """新しく開いたウィンドウが裏に隠れて止まって見える問題への対策。
    ウィンドウが開いている間はずっと最前面(topmost)を維持する。"""
    win.lift()
    win.attributes("-topmost", True)
    win.focus_force()


def load_store_table(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[LIST_SHEET]

    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    name_idx = header.index(NAME_COL)
    area_idx = header.index(AREA_COL)

    stores = []  # (num3, rest, area, raw_store_name)
    for row in rows[1:]:
        store_name = row[name_idx]
        area = row[area_idx]
        if not store_name:
            continue
        m = FRONT_PATTERN.match(str(store_name))
        if not m:
            continue
        num, rest = m.groups()
        stores.append((num.zfill(3), rest, area, store_name))

    return stores


def find_area(stores, file_num, file_place):
    matches = []
    if not file_place:
        return matches
    normalized_place = normalize_text(file_place)
    for num3, rest, area, raw in stores:
        if num3 != file_num:
            continue
        core = normalize_text(core_place_name(rest))
        if core and core in normalized_place:
            matches.append((area, raw))
    return matches


def classify_files(folder, stores):
    """自動でエリアが確定するものと、確認が必要なものに分ける"""
    entries = sorted(os.listdir(folder))
    files = [n for n in entries if os.path.isfile(os.path.join(folder, n))]

    auto = []        # (old_name, area) 自動で確定
    need_confirm = []  # (old_name, initial_query) 手動確認が必要

    for name in files:
        stem, ext = os.path.splitext(name)
        m = FRONT_PATTERN.match(stem)
        if not m:
            need_confirm.append((name, stem))
            continue
        num, place = m.groups()
        num3 = num.zfill(3)

        matches = find_area(stores, num3, place)
        distinct_areas = {a for a, r in matches if a}

        if len(distinct_areas) == 1:
            auto.append((name, next(iter(distinct_areas))))
        else:
            need_confirm.append((name, place))

    return files, auto, need_confirm


def ask_store_for_file(root, filename, stores, folder, initial_query=""):
    """店舗名を検索して選ばせるダイアログ。選ばれた raw 店舗名を返す(スキップ時はNone)"""
    print(f"    -> ダイアログ作成中: {filename}", flush=True)
    result = {"choice": None}

    dialog = tk.Toplevel(root)
    dialog.title("店舗を選択してください")
    dialog.geometry("520x420+200+200")
    dialog.grab_set()
    # 注意: root が withdraw() されているため transient(root) は付けない。
    # (非表示のウィンドウに transient すると、このダイアログ自体も描画されなくなる)

    tk.Label(
        dialog,
        text=f"エリアを自動判定できませんでした。\nファイル名: {filename}\n該当する店舗名を検索して選んでください。",
        justify="left",
        anchor="w",
    ).pack(fill="x", padx=10, pady=(10, 5))

    search_var = tk.StringVar(value=initial_query)
    entry = tk.Entry(dialog, textvariable=search_var)
    entry.pack(fill="x", padx=10, pady=5)

    listbox = tk.Listbox(dialog)
    listbox.pack(fill="both", expand=True, padx=10, pady=5)

    display_list = []  # (raw_name, area)

    def refresh(*_args):
        query = search_var.get().strip()
        listbox.delete(0, tk.END)
        display_list.clear()
        for _num3, _rest, area, raw in stores:
            if query == "" or query in str(raw) or (area and query in str(area)):
                display_list.append((raw, area))
        for raw, area in display_list:
            listbox.insert(tk.END, f"{raw}   [{area}]")
        if display_list:
            listbox.selection_set(0)

    def confirm(_event=None):
        sel = listbox.curselection()
        if not sel:
            return
        raw, _area = display_list[sel[0]]
        result["choice"] = raw
        dialog.destroy()

    def skip():
        result["choice"] = None
        dialog.destroy()

    def back_to_folder_select():
        print("[戻る] フォルダ選択画面に戻ります", flush=True)
        result["back"] = True
        dialog.destroy()

    def abort_program():
        print(f"[中止] ユーザー操作によりプログラムを中止します。フォルダを開きます: {folder}", flush=True)
        os.startfile(folder)
        os._exit(0)

    search_var.trace_add("write", refresh)
    listbox.bind("<Double-Button-1>", confirm)
    entry.bind("<Return>", confirm)

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(fill="x", padx=10, pady=10)
    tk.Button(btn_frame, text="決定", command=confirm).pack(side="left")
    tk.Button(btn_frame, text="この回はスキップ", command=skip).pack(side="left", padx=5)
    tk.Button(btn_frame, text="戻る", command=back_to_folder_select).pack(side="left", padx=5)
    tk.Button(btn_frame, text="プログラム中止", command=abort_program, fg="red").pack(side="right")

    refresh()
    dialog.deiconify()
    dialog.update()
    bring_to_front(dialog)
    entry.focus_set()
    entry.selection_range(0, tk.END)
    print(f"    -> ダイアログ表示完了。入力待ちです: {filename}", flush=True)

    dialog.protocol("WM_DELETE_WINDOW", skip)
    root.wait_window(dialog)
    if result.get("back"):
        raise GoBackToFolderSelection()
    print(f"    -> ダイアログ終了: {filename} -> {result['choice']!r}", flush=True)
    return result["choice"]


def resolve_need_confirm(root, need_confirm, stores, folder):
    print(f"[確認] 要確認ファイル {len(need_confirm)} 件の処理を開始します", flush=True)
    area_by_raw = {raw: area for _n, _r, area, raw in stores}

    resolved = []   # (old_name, area)
    skipped = []    # old_name

    for i, (name, initial_query) in enumerate(need_confirm, start=1):
        print(f"[確認] {i}/{len(need_confirm)}: {name}", flush=True)
        raw = ask_store_for_file(root, name, stores, folder, initial_query)
        if raw is None:
            skipped.append(name)
            continue
        area = area_by_raw.get(raw)
        if not area:
            skipped.append(name)
            continue
        resolved.append((name, area))

    return resolved, skipped


def build_safe_renames(files, decided):
    """decided: [(old_name, area), ...] から、重複しないリネーム計画を作る"""
    renames = []
    for old_name, area in decided:
        new_name = f"{area}_{old_name}"
        if new_name != old_name:
            renames.append((old_name, new_name))

    final_names = {}
    decided_names = {old for old, _ in decided}
    for name in files:
        if name not in decided_names:
            final_names.setdefault(name, []).append(name)
    for old_name, new_name in renames:
        final_names.setdefault(new_name, []).append(old_name)

    conflicts = {new: olds for new, olds in final_names.items() if len(olds) > 1}

    safe_renames = []
    errors = []
    for old_name, new_name in renames:
        if new_name in conflicts:
            errors.append(f"{old_name} -> {new_name} (名前が重複するためスキップ)")
        else:
            safe_renames.append((old_name, new_name))

    return safe_renames, errors


def show_progress_dialog(root, total):
    dialog = tk.Toplevel(root)
    dialog.title("リネーム中")
    dialog.geometry("360x110")
    dialog.resizable(False, False)
    # 注意: root が withdraw() されているため transient(root) は付けない(描画されなくなるため)
    dialog.grab_set()
    dialog.protocol("WM_DELETE_WINDOW", lambda: None)  # 完了まで閉じさせない

    label_var = tk.StringVar(value=f"0 / {total} 件")
    tk.Label(dialog, textvariable=label_var).pack(pady=(15, 5))

    bar = ttk.Progressbar(dialog, orient="horizontal", length=300, mode="determinate", maximum=max(total, 1))
    bar.pack(pady=5)

    bring_to_front(dialog)
    dialog.update_idletasks()

    def update(step, step_total):
        bar["maximum"] = max(step_total, 1)
        bar["value"] = step
        label_var.set(f"{step} / {step_total} 件")
        dialog.update_idletasks()
        dialog.update()

    return dialog, update


def apply_renames(folder, renames, progress_callback=None):
    total_steps = len(renames) * 2
    step = 0

    temp_map = []
    for old_name, new_name in renames:
        old_path = os.path.join(folder, old_name)
        temp_path = os.path.join(folder, old_name + ".__tmp_rename__")
        os.rename(old_path, temp_path)
        temp_map.append((temp_path, os.path.join(folder, new_name)))
        step += 1
        if progress_callback:
            progress_callback(step, total_steps)

    for temp_path, new_path in temp_map:
        os.rename(temp_path, new_path)
        step += 1
        if progress_callback:
            progress_callback(step, total_steps)


def main():
    print("[1/7] tkinter初期化中...", flush=True)
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    print("[2/7] 初期化完了", flush=True)

    print(f"[3/7] 店舗リストを読み込み中: {LIST_PATH}", flush=True)
    try:
        stores = load_store_table(LIST_PATH)
    except Exception as e:
        messagebox.showerror("エラー", f"店舗リストの読み込みに失敗しました:\n{LIST_PATH}\n\n{e}", parent=root)
        return
    print(f"[4/7] 店舗リスト読み込み完了: {len(stores)} 件", flush=True)

    while True:
        print("[5/7] フォルダ選択ダイアログを表示します(画面裏に隠れていないか確認してください)", flush=True)
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(title="画像が入っているフォルダを選択してください", parent=root)
        root.attributes("-topmost", False)
        print(f"[6/7] 選択されたフォルダ: {folder!r}", flush=True)
        if not folder:
            print("フォルダが選択されなかったため終了します", flush=True)
            return

        files, auto, need_confirm = classify_files(folder, stores)
        print(f"[7/7] 対象ファイル数: {len(files)} / 自動判定: {len(auto)} / 要確認: {len(need_confirm)}", flush=True)

        resolved, skipped = ([], [])
        try:
            if need_confirm:
                resolved, skipped = resolve_need_confirm(root, need_confirm, stores, folder)
        except GoBackToFolderSelection:
            continue

        break

    decided = auto + resolved
    renames, errors = build_safe_renames(files, decided)

    if renames:
        print(f"リネーム開始: {len(renames)} 件", flush=True)
        progress_dialog, update_progress = show_progress_dialog(root, len(renames) * 2)
        apply_renames(folder, renames, progress_callback=update_progress)
        progress_dialog.destroy()
        print("リネーム処理完了", flush=True)

    lines = [f"リネーム完了: {len(renames)} 件"]
    if renames:
        lines.append("")
        lines.extend(f"  {old} -> {new}" for old, new in renames)
    if skipped:
        lines.append("")
        lines.append(f"確認の結果スキップ: {len(skipped)} 件")
        lines.extend(f"  {n}" for n in skipped)
    if errors:
        lines.append("")
        lines.append(f"名前の重複でスキップ: {len(errors)} 件")
        lines.extend(f"  {e}" for e in errors)

    result_text = "\n".join(lines)
    print(result_text)
    root.attributes("-topmost", True)
    messagebox.showinfo("結果", result_text, parent=root)
    root.attributes("-topmost", False)


if __name__ == "__main__":
    main()
