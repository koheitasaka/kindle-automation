"""
Kindle for PC 自動スクリーンショット (Windows版)

macOS版 kindle_shot_v7_auto.py / kindle_shot_horizontal_auto.py を統合。
AppleScript → pygetwindow, screencapture → pyautogui, osascript → pyautogui.press に置換。

Usage:
    python kindle_capture.py                  # 縦書き（左矢印でページ送り）
    python kindle_capture.py --horizontal     # 横書き（右矢印でページ送り）
    python kindle_capture.py --delay 2.0      # ページめくり待機時間を変更
    python kindle_capture.py --max-pages 500  # 最大ページ数を変更
    python kindle_capture.py --output-dir ./out  # 出力先を変更
"""

import argparse
import os
import sys
import time

import cv2
import numpy as np
import pyautogui
import pygetwindow as gw
from PIL import Image


TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "review_modal_template_v2.png")


def find_kindle_window():
    """Kindle for PC のウィンドウを検索して返す"""
    candidates = [w for w in gw.getAllWindows() if w.title and "kindle" in w.title.lower()]
    if not candidates:
        return None
    # 最も大きいウィンドウを優先（メインウィンドウの可能性が高い）
    candidates.sort(key=lambda w: w.width * w.height, reverse=True)
    return candidates[0]


def activate_kindle(win):
    """Kindleウィンドウを前面に持ってくる"""
    try:
        if win.isMinimized:
            win.restore()
        win.activate()
        time.sleep(0.3)
    except Exception as e:
        print(f"Warning: ウィンドウのアクティブ化に失敗 ({e})")


def capture_kindle_window(win, index: int, save_dir: str):
    """Kindleウィンドウ領域をスクリーンショットし、保存する"""
    left = win.left
    top = win.top
    width = win.width
    height = win.height

    if width <= 0 or height <= 0:
        print("Error: ウィンドウサイズが不正です。")
        return False, None

    screenshot = pyautogui.screenshot(region=(left, top, width, height))

    filename = f"page_{index:04d}.png"
    filepath = os.path.join(save_dir, filename)
    screenshot.save(filepath)

    return True, filepath


def turn_page(horizontal: bool):
    """ページを送る（縦書き: 左矢印, 横書き: 右矢印）"""
    key = "right" if horizontal else "left"
    pyautogui.press(key)


def detect_modal(screenshot_path: str, template_path: str, threshold: float = 0.8) -> bool:
    """スクリーンショット内にレビューモーダルを検出する"""
    if not os.path.exists(template_path):
        return False

    try:
        template = cv2.imread(template_path, cv2.IMREAD_COLOR)
        screenshot = cv2.imread(screenshot_path, cv2.IMREAD_COLOR)

        if template is None or screenshot is None:
            return False

        # テンプレートがスクリーンショットより大きい場合はスキップ
        if (template.shape[0] > screenshot.shape[0] or
                template.shape[1] > screenshot.shape[1]):
            return False

        template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

        res = cv2.matchTemplate(screenshot_gray, template_gray, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, _ = cv2.minMaxLoc(res)

        print(f"  Modal detection: {max_val:.2f} (threshold: {threshold:.2f})")
        if max_val >= threshold:
            print(f"  -> レビューモーダル検出 (confidence: {max_val:.2f})")
            return True
        return False
    except Exception as e:
        print(f"  Warning: モーダル検出エラー ({e})")
        return False


def detect_duplicate(prev_path: str, curr_path: str, threshold: float = 0.99) -> bool:
    """前ページと現ページが同一（最終ページ到達）かを判定する"""
    if prev_path is None or not os.path.exists(prev_path):
        return False

    try:
        prev = cv2.imread(prev_path, cv2.IMREAD_GRAYSCALE)
        curr = cv2.imread(curr_path, cv2.IMREAD_GRAYSCALE)

        if prev is None or curr is None:
            return False
        if prev.shape != curr.shape:
            return False

        # 正規化相関で比較
        res = cv2.matchTemplate(curr, prev, cv2.TM_CCOEFF_NORMED)
        similarity = res[0][0]
        if similarity >= threshold:
            print(f"  -> 同一ページ検出 (similarity: {similarity:.4f}) — 最終ページと判断")
            return True
        return False
    except Exception:
        return False


def parse_args():
    parser = argparse.ArgumentParser(
        description="Kindle for PC 自動スクリーンショット (Windows版)"
    )
    parser.add_argument(
        "--horizontal",
        action="store_true",
        help="横書きモード（右矢印でページ送り）。デフォルトは縦書き（左矢印）。",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.5,
        help="ページめくり後の待機秒数 (default: 1.5)",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=1000,
        help="最大ページ数 (default: 1000)",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="保存先ディレクトリ (default: ~/Desktop/KindleScreenshots)",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.8,
        help="モーダル検出の閾値 (default: 0.8)",
    )
    parser.add_argument(
        "--no-modal-detection",
        action="store_true",
        help="モーダル検出を無効化（Ctrl+C で手動停止）",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    # 保存先ディレクトリ
    if args.output_dir:
        save_dir = os.path.abspath(args.output_dir)
    else:
        suffix = "_Horizontal" if args.horizontal else ""
        save_dir = os.path.expanduser(f"~/Desktop/KindleScreenshots{suffix}")
    os.makedirs(save_dir, exist_ok=True)

    mode = "横書き (→)" if args.horizontal else "縦書き (←)"
    print(f"=== Kindle for PC 自動キャプチャ ===")
    print(f"  モード: {mode}")
    print(f"  保存先: {save_dir}")
    print(f"  ページめくり待機: {args.delay}s")
    print(f"  最大ページ数: {args.max_pages}")

    # テンプレート有無の確認
    has_template = os.path.exists(TEMPLATE_PATH) and not args.no_modal_detection
    if has_template:
        print(f"  モーダル検出: ON (threshold={args.threshold})")
    else:
        print(f"  モーダル検出: OFF (Ctrl+C で停止してください)")

    print()

    # Kindleウィンドウ検索
    win = find_kindle_window()
    if win is None:
        print("Error: Kindle for PC のウィンドウが見つかりません。")
        print("  Kindle for PC を起動し、本を開いた状態にしてください。")
        sys.exit(1)

    print(f"Kindle ウィンドウ検出: \"{win.title}\" ({win.width}x{win.height})")
    activate_kindle(win)

    print("3秒後にキャプチャを開始します... Kindleの本の部分をクリックしてフォーカスを当ててください。")
    time.sleep(3)

    prev_path = None
    index = 1

    try:
        while index <= args.max_pages:
            print(f"--- Page {index} ---")

            # Kindleを再検索（ウィンドウが移動/リサイズされた場合に対応）
            win = find_kindle_window()
            if win is None:
                print("Error: Kindleウィンドウが見つかりません。終了します。")
                break

            success, filepath = capture_kindle_window(win, index, save_dir)
            if not success:
                print("Error: スクリーンショットの取得に失敗しました。")
                break

            # モーダル検出
            if has_template and detect_modal(filepath, TEMPLATE_PATH, args.threshold):
                print("レビューモーダルを検出しました。キャプチャを停止します。")
                # モーダルが写ったスクリーンショットを削除
                os.remove(filepath)
                break

            # 同一ページ検出（最終ページ判定）
            if detect_duplicate(prev_path, filepath):
                # 重複スクリーンショットを削除
                os.remove(filepath)
                break

            prev_path = filepath
            time.sleep(0.3)

            turn_page(args.horizontal)
            index += 1
            time.sleep(args.delay)

    except KeyboardInterrupt:
        print("\n\nCtrl+C を検出しました。キャプチャを停止します。")

    total = index - 1
    print(f"\n完了！ {total} ページをキャプチャしました。")
    print(f"保存先: {save_dir}")


if __name__ == "__main__":
    main()
