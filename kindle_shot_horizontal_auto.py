import time
import os
import subprocess
import cv2
import numpy as np
from PIL import Image

# 保存先ディレクトリの作成（横書き用）
SAVE_DIR = os.path.expanduser("~/Desktop/KindleScreenshots_Horizontal")
if not os.path.exists(SAVE_DIR):
    os.makedirs(SAVE_DIR)

# レビューモーダルのテンプレート画像パス
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "review_modal_template_v2.png")

def capture_kindle_window_safe(index, temp_filepath="/tmp/temp_kindle_screenshot.png"):
    """Kindleのウィンドウだけを確実にキャプチャし、指定された一時ファイルに保存する"""
    
    # AppleScriptで座標を取得（文字列で返す）
    script = '''
    tell application "System Events"
        tell process "Kindle"
            set frontmost to true
            if exists window 1 then
                set win to window 1
                set pos to position of win
                set siz to size of win
                return (item 1 of pos as string) & "," & (item 2 of pos as string) & "," & (item 1 of siz as string) & "," & (item 2 of siz as string)
            else
                return "ERROR: No Window"
            end if
        end tell
    end tell
    '''
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    coords = result.stdout.strip()
    
    if "ERROR" in coords or not coords:
        print("Kindleのウィンドウが見つかりませんでした。")
        return False, None
        
    subprocess.run(["screencapture", "-x", "-R" + coords, temp_filepath])
    
    # スクリーンショットを正式なパスに移動
    filename = f"page_{index:04d}.png"
    filepath = os.path.join(SAVE_DIR, filename)
    os.rename(temp_filepath, filepath) # 一時ファイルから正式な場所に移動
    
    return True, filepath

def turn_page():
    """右矢印キー（横書き本の次ページ）を送信する"""
    script = '''
    tell application "System Events"
        tell process "Kindle"
            set frontmost to true
            key code 124 -- Right Arrow
        end tell
    end tell
    '''
    subprocess.run(["osascript", "-e", script])

def detect_modal(screenshot_path, template_path, threshold=0.8):
    """スクリーンショット内にモーダルテンプレートを検出する"""
    if not os.path.exists(template_path):
        print(f"Error: Template image not found at {template_path}")
        return False

    try:
        # テンプレートとスクリーンショットを読み込む
        template = cv2.imread(template_path, cv2.IMREAD_COLOR)
        screenshot = cv2.imread(screenshot_path, cv2.IMREAD_COLOR)

        if template is None:
            print(f"Error: Could not load template image from {template_path}")
            return False
        if screenshot is None:
            print(f"Error: Could not load screenshot image from {screenshot_path}")
            return False

        # グレースケールに変換（処理を高速化するため）
        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

        # テンプレートマッチングを実行
        res = cv2.matchTemplate(screenshot_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

        # マッチングの類似度がしきい値を超えているかチェック
        if max_val >= threshold:
            print(f"Modal detected with confidence: {max_val:.2f}")
            return True
        else:
            return False
    except Exception as e:
        print(f"Error during modal detection: {e}")
        return False

def main():
    print(f"Saving screenshots to {SAVE_DIR}")
    print("Kindle 横書き本対応版 - 右矢印でページ送り (自動停止機能付き)")
    
    # Kindleを前面に
    subprocess.run(["osascript", "-e", 'tell application id "com.amazon.Lassen" to activate'])
    time.sleep(2)
    
    print("Kindleの本の部分を一度クリックしてください（フォーカスを当てるため）...")
    time.sleep(3)
    
    index = 1
    while True:
        print(f"--- Processing Page {index} ---")
        success, current_screenshot_path = capture_kindle_window_safe(index, temp_filepath="/tmp/temp_kindle_screenshot_h.png")
        if not success:
            print("スクリーンショットの取得に失敗しました。")
            break
            
        # モーダル検知
        if detect_modal(current_screenshot_path, TEMPLATE_PATH):
            print("レビューを促すモーダルを検出しました。スクリーンショットの取得を停止します。")
            break
            
        time.sleep(0.5)
        
        print(f"Sending Right Arrow key (Next Page for horizontal books)...")
        turn_page()
        index += 1
        
        # ページめくりアニメーションを待つ
        time.sleep(1.5)
        
        # 安全のために1000ページで停止
        if index > 1000:
            print("1000ページの上限に達しました。処理を停止します。")
            break

    print("Done!")

if __name__ == "__main__":
    main()
