"""
Kindle スクリーンショット → PDF 変換 (Windows版)

img2pdf で無劣化PDF変換 + ocrmypdf で日本語OCR対応。

Usage:
    python make_pdf_win.py                          # OCRなし（高速）
    python make_pdf_win.py --ocr                    # 日本語OCR付き
    python make_pdf_win.py --ocr --lang jpn_vert    # 縦書きOCR
    python make_pdf_win.py --input-dir ./screenshots --output book.pdf
"""

import argparse
import glob
import os
import shutil
import subprocess
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


def check_tesseract() -> bool:
    """Tesseract OCRがインストールされているか確認する"""
    if shutil.which("tesseract"):
        return True
    # wingetデフォルトインストール先を確認
    default_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(default_path):
        os.environ["PATH"] = os.path.dirname(default_path) + os.pathsep + os.environ.get("PATH", "")
        return True
    return False


def run_ocr(input_pdf: str, output_pdf: str, lang: str):
    """ocrmypdfでOCR処理を実行する"""
    if not check_tesseract():
        print("Error: Tesseract OCR が見つかりません。")
        print("  インストール: winget install UB-Mannheim.TesseractOCR")
        print("  日本語データ: https://github.com/tesseract-ocr/tessdata_best")
        sys.exit(1)

    print(f"OCR処理中 (lang={lang})... これには時間がかかります。")

    try:
        import ocrmypdf
        ocrmypdf.ocr(
            input_pdf,
            output_pdf,
            language=lang,
            optimize=1,
            skip_text=True,
            progress_bar=True,
        )
        size_mb = os.path.getsize(output_pdf) / (1024 * 1024)
        print(f"OCR完了！ {output_pdf} ({size_mb:.1f} MB)")
    except ocrmypdf.exceptions.MissingDependencyError as e:
        print(f"Error: 依存ツールが不足しています — {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: OCR処理に失敗しました — {e}")
        sys.exit(1)


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
    parser.add_argument(
        "--ocr",
        action="store_true",
        help="OCR処理を実行してテキスト検索可能なPDFを生成する",
    )
    parser.add_argument(
        "--lang",
        type=str,
        default="jpn",
        help="OCR言語 (default: jpn)。縦書きは jpn_vert を指定。",
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
    print(f"  OCR: {'ON (' + args.lang + ')' if args.ocr else 'OFF'}")
    print()

    # 検証
    valid_files = validate_images(files)
    if not valid_files:
        print("Error: 有効な画像ファイルがありません。")
        sys.exit(1)

    if len(valid_files) < len(files):
        print(f"  {len(files) - len(valid_files)} ファイルをスキップしました。")
        print()

    if args.ocr:
        # OCRモード: img2pdf → 一時PDF → ocrmypdf → 最終PDF
        temp_pdf = output.replace(".pdf", "_temp.pdf")
        create_pdf(valid_files, temp_pdf)
        print()
        run_ocr(temp_pdf, output, args.lang)
        # 一時ファイル削除
        if os.path.exists(temp_pdf) and os.path.exists(output):
            os.remove(temp_pdf)
    else:
        # 通常モード: img2pdf のみ
        create_pdf(valid_files, output)


if __name__ == "__main__":
    main()
