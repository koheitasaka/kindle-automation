"""
Kindle スクリーンショット → PDF 変換 (Windows版)

macOS版の ImageMagick + ocrmypdf を img2pdf（無劣化・高速）に置換。

Usage:
    python make_pdf_win.py
    python make_pdf_win.py --input-dir ./screenshots --output book.pdf
"""

import argparse
import glob
import os
import sys

import img2pdf
from PIL import Image


def collect_images(input_dir: str) -> list[str]:
    """指定ディレクトリからPNG画像を連番順に収集する"""
    patterns = ["*.png", "*.jpg", "*.jpeg"]
    files: list[str] = []
    for pat in patterns:
        files.extend(glob.glob(os.path.join(input_dir, pat)))

    # ファイル名でソート（page_0001.png, page_0002.png, ...）
    files.sort()
    return files


def validate_images(files: list[str]) -> list[str]:
    """画像ファイルが正常に読み込めるか検証し、有効なもののみ返す"""
    valid: list[str] = []
    for f in files:
        try:
            with Image.open(f) as img:
                img.verify()
            valid.append(f)
        except Exception as e:
            print(f"  Warning: スキップ — {os.path.basename(f)} ({e})")
    return valid


def create_pdf(image_files: list[str], output_path: str):
    """画像リストから無劣化PDFを生成する"""
    print(f"PDF生成中... ({len(image_files)} ページ)")

    with open(output_path, "wb") as f:
        f.write(img2pdf.convert(image_files))

    size_mb = os.path.getsize(output_path) / (1024 * 1024)
    print(f"完了！ {output_path} ({size_mb:.1f} MB)")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Kindle スクリーンショットからPDFを生成 (Windows版)"
    )
    parser.add_argument(
        "--input-dir",
        type=str,
        default=None,
        help="画像ディレクトリ (default: ~/Desktop/KindleScreenshots)",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="出力PDFパス (default: ~/Desktop/KindleBook.pdf)",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    input_dir = args.input_dir or os.path.expanduser("~/Desktop/KindleScreenshots")
    output = args.output or os.path.expanduser("~/Desktop/KindleBook.pdf")

    input_dir = os.path.abspath(input_dir)
    output = os.path.abspath(output)

    if not os.path.isdir(input_dir):
        print(f"Error: ディレクトリが見つかりません — {input_dir}")
        sys.exit(1)

    # 画像収集
    files = collect_images(input_dir)
    if not files:
        print(f"Error: 画像ファイルが見つかりません — {input_dir}")
        sys.exit(1)

    print(f"=== Kindle PDF生成 ===")
    print(f"  入力: {input_dir} ({len(files)} ファイル)")
    print(f"  出力: {output}")
    print()

    # 検証
    valid_files = validate_images(files)
    if not valid_files:
        print("Error: 有効な画像ファイルがありません。")
        sys.exit(1)

    if len(valid_files) < len(files):
        print(f"  {len(files) - len(valid_files)} ファイルをスキップしました。")
        print()

    # PDF生成
    create_pdf(valid_files, output)


if __name__ == "__main__":
    main()
