r"""
Place an image onto a slide, sized by orientation and centered on the
page (both left-right and top-bottom).

Rule:
    landscape image (width >= height): displayed width = LANDSCAPE_WIDTH_CM
    portrait image  (height > width):  displayed height = PORTRAIT_HEIGHT_CM
    (aspect ratio is preserved either way)

Two ways to use:

1) Drag & drop mode (no arguments): double-click this script (or run
   `python place_image_centered.py`), leave the console window open,
   then drag a .pptx file from Explorer onto the console window and
   press Enter. Every slide that already has a picture on it gets
   that picture resized/re-centered per the rule above (the image
   itself is kept, only its size and position are fixed), and the
   "図NN"/"QRコードNN" label is moved to the image's top-left corner.
   The result is saved into Downloads as "<元のファイル名>_加工済み.pptx".
   It then waits for the next file. Type nothing and press Enter (or
   type exit) to quit.

2) One-shot CLI mode (insert a new image into one slide):
   python place_image_centered.py "template.pptx" 1 "image.jpg" "output.pptx"

    template.pptx : the pptx to edit
    1             : 1-based slide number whose picture should be replaced
                    (existing PICTURE shape(s) on that slide are removed)
    image.jpg     : the new image to insert
    output.pptx   : where to save the result (omit to overwrite the input)
"""
import sys
import os
import re
import tempfile
from pptx import Presentation
from pptx.util import Cm, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from PIL import Image

LANDSCAPE_WIDTH_CM = 12.76
PORTRAIT_HEIGHT_CM = 16.7
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")

# textboxes like "図00", "図1", "QRコード1" are treated as the image's label
LABEL_PATTERN = re.compile(r"^(図|QRコード)\s*[0-9０-９]+$")


def is_picture_shape(shape):
    """True for any <p:pic> element, including pictures that were inserted
    into a placeholder (python-pptx reports those as shape_type PLACEHOLDER,
    not PICTURE, even though they hold real image data)."""
    return shape._element.tag == qn("p:pic")


def find_label_shape(slide):
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX and shape.has_text_frame:
            if LABEL_PATTERN.match(shape.text_frame.text.strip()):
                return shape
    return None


def place_image_centered(slide, slide_width, slide_height, image_path):
    # remove existing pictures on the slide so the new one takes their place
    for shape in list(slide.shapes):
        if is_picture_shape(shape):
            shape._element.getparent().remove(shape._element)

    with Image.open(image_path) as im:
        iw, ih = im.size

    if iw >= ih:  # landscape (or square)
        width = Cm(LANDSCAPE_WIDTH_CM)
        height = Emu(int(width * ih / iw))
    else:  # portrait
        height = Cm(PORTRAIT_HEIGHT_CM)
        width = Emu(int(height * iw / ih))
        # a near-square "portrait" image can end up wider than the page;
        # cap the width and derive height from that instead
        if width > Cm(LANDSCAPE_WIDTH_CM):
            width = Cm(LANDSCAPE_WIDTH_CM)
            height = Emu(int(width * ih / iw))

    left = Emu(int((slide_width - width) / 2))
    top = Emu(int((slide_height - height) / 2))

    slide.shapes.add_picture(image_path, left, top, width, height)

    # move the "図NN" label so its bottom-left corner sits at the image's
    # top-left corner (label directly above the image, text stays visible),
    # wherever it happened to be placed originally
    label = find_label_shape(slide)
    if label is not None:
        label.left = left
        label.top = Emu(int(top - label.height))


def fix_existing_pictures(prs):
    """Re-size/re-center every picture already present in the deck,
    keeping the image itself. Returns the number of pictures fixed."""
    n = 0
    with tempfile.TemporaryDirectory() as tmp_dir:
        for slide in prs.slides:
            pics = [s for s in slide.shapes if is_picture_shape(s)]
            for pic in pics:
                tmp_path = os.path.join(tmp_dir, f"_extract_{n}.{pic.image.ext}")
                with open(tmp_path, "wb") as f:
                    f.write(pic.image.blob)
                place_image_centered(slide, prs.slide_width, prs.slide_height, tmp_path)
                n += 1
    return n


def split_paths(line):
    """Split a line of one or more (optionally quoted) Windows paths
    without mangling backslashes the way shlex.split() would."""
    return [quoted or bare for quoted, bare in re.findall(r'"([^"]+)"|(\S+)', line)]


def process_one(path):
    path = path.strip().strip('"').strip("'")
    if not path:
        return None
    if not os.path.isfile(path):
        print(f"  !! ファイルが見つかりません: {path}")
        return None
    if not path.lower().endswith(".pptx"):
        print(f"  !! .pptx ではありません: {path}")
        return None

    try:
        prs = Presentation(path)
        n = fix_existing_pictures(prs)

        base = os.path.splitext(os.path.basename(path))[0]
        os.makedirs(DOWNLOADS_DIR, exist_ok=True)
        output_path = os.path.join(DOWNLOADS_DIR, f"{base}_加工済み.pptx")
        prs.save(output_path)

        print(f"  {n}枚の画像を調整しました")
        print(f"  保存先: {output_path}")
        return output_path
    except Exception as e:
        print(f"  !! エラー: {e}")
        return None


def interactive_loop():
    print("=== PDF画像調整 (ドラッグ&ドロップ待機モード) ===")
    print(f"保存先フォルダ: {DOWNLOADS_DIR}")
    print(".pptx ファイルをこのウィンドウにドラッグ&ドロップして Enter を押してください。")
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
        output_paths = [p for p in (process_one(path) for path in split_paths(line)) if p]
        print("完了しました。")
        for output_path in output_paths:
            os.startfile(output_path)
        return


def main():
    if len(sys.argv) == 1:
        interactive_loop()
        return

    # a single .pptx argument (e.g. dragged onto the .bat) is handled the
    # same way as a line typed into the interactive loop
    if len(sys.argv) == 2 and sys.argv[1].lower().endswith(".pptx"):
        process_one(sys.argv[1])
        return

    template_path = sys.argv[1]
    slide_no = int(sys.argv[2])
    image_path = sys.argv[3]
    output_path = sys.argv[4] if len(sys.argv) > 4 else template_path

    prs = Presentation(template_path)
    slide = list(prs.slides)[slide_no - 1]

    place_image_centered(slide, prs.slide_width, prs.slide_height, image_path)

    prs.save(output_path)
    print("saved:", output_path)


if __name__ == "__main__":
    main()
