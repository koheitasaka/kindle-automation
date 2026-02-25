import os
import subprocess
import sys

def create_ocr_pdf(image_dir, output_filename):
    """画像ディレクトリからOCR済みPDFを生成する"""
    # 1. 画像をPDFに変換 (ImageMagick)
    print("Converting images to a single PDF...")
    images_path = os.path.join(image_dir, "*.png")
    temp_pdf = "temp_combined.pdf"
    
    # 画像ファイルが連番（page_0001.png...）なので、ワイルドカードで順番通りになる
    try:
        subprocess.run(["magick", images_path, temp_pdf], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error during image conversion: {e}")
        return

    # 2. OCRを実行 (ocrmypdf)
    print("Running OCR (Japanese)... This might take a while.")
    try:
        subprocess.run([
            "ocrmypdf", 
            "-l", "jpn", 
            "--optimize", "1", 
            "--skip-text", 
            temp_pdf, 
            output_filename
        ], check=True)
        print(f"Success! Saved as {output_filename}")
    except subprocess.CalledProcessError as e:
        print(f"Error during OCR: {e}")
    finally:
        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)

if __name__ == "__main__":
    IMAGE_DIR = os.path.expanduser("~/Desktop/KindleScreenshots")
    OUTPUT = os.path.expanduser("~/Desktop/Kindle_OCR_Book.pdf")
    create_ocr_pdf(IMAGE_DIR, OUTPUT)
