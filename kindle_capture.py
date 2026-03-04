"""
Kindle for PC 自動スクリーンショット (Windows版)

macOS版 kindle_shot_v7_auto.py / kindle_shot_horizontal_auto.py を統合。
AppleScript → pygetwindow, screencapture → pyautogui, osascript → pyautogui.press に置換。

Usage:
    python kindle_capture.py                  # 縦書き（左矢印でページ送り）
    python kindle_capture.py --horizontal     # 横書き（右矢印でページ送り）
    python kindle_capture.py --background     # バックグラウンド実行（作業しながらOK）
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
from PIL import Image


TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "review_modal_template_v2.png")


# ---------------------------------------------------------------------------
# フォアグラウンドモード (pyautogui + pygetwindow)
# ---------------------------------------------------------------------------

def fg_find_kindle_window():
    """Kindle for PC のウィンドウを検索して返す"""
    import pygetwindow as gw
    candidates = [w for w in gw.getAllWindows() if w.title and "kindle" in w.title.lower()]
    if not candidates:
        return None
    candidates.sort(key=lambda w: w.width * w.height, reverse=True)
    return candidates[0]


def fg_activate_kindle(win):
    """Kindleウィンドウを前面に持ってくる"""
    try:
        if win.isMinimized:
            win.restore()
        win.activate()
        time.sleep(0.3)
    except Exception as e:
        print(f"Warning: ウィンドウのアクティブ化に失敗 ({e})")


def fg_capture(win, index: int, save_dir: str):
    """pyautogui でウィンドウ領域をキャプチャする"""
    import pyautogui
    left, top, width, height = win.left, win.top, win.width, win.height
    if width <= 0 or height <= 0:
        return False, None

    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    filepath = os.path.join(save_dir, f"page_{index:04d}.png")
    screenshot.save(filepath)
    return True, filepath


def fg_turn_page(horizontal: bool):
    """pyautogui でキー送信する"""
    import pyautogui
    pyautogui.press("right" if horizontal else "left")


# ---------------------------------------------------------------------------
# バックグラウンドモード (Win32 API: PrintWindow + PostMessage)
# ---------------------------------------------------------------------------

def bg_find_kindle_hwnd():
    """Win32 API で Kindle ウィンドウハンドルを取得する"""
    import win32gui
    result = []

    def enum_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and "kindle" in title.lower():
                result.append((hwnd, title))

    win32gui.EnumWindows(enum_callback, None)
    if not result:
        return None, None

    best_hwnd, best_title = result[0]
    best_area = 0
    for hwnd, title in result:
        rect = win32gui.GetWindowRect(hwnd)
        area = (rect[2] - rect[0]) * (rect[3] - rect[1])
        if area > best_area:
            best_area = area
            best_hwnd = hwnd
            best_title = title

    return best_hwnd, best_title


def bg_get_window_size(hwnd):
    """ウィンドウのサイズを返す"""
    import win32gui
    rect = win32gui.GetWindowRect(hwnd)
    return rect[2] - rect[0], rect[3] - rect[1]


def bg_capture(hwnd, index: int, save_dir: str):
    """PrintWindow API でバックグラウンドキャプチャする"""
    import win32gui
    import win32ui
    from ctypes import windll

    rect = win32gui.GetWindowRect(hwnd)
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]

    if width <= 0 or height <= 0:
        return False, None

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(bitmap)

    # PW_RENDERFULLCONTENT=2 でハードウェアアクセラレーション対応
    result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)
    if result == 0:
        windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 0)

    bmp_info = bitmap.GetInfo()
    bmp_bits = bitmap.GetBitmapBits(True)

    img = Image.frombuffer(
        "RGB",
        (bmp_info["bmWidth"], bmp_info["bmHeight"]),
        bmp_bits,
        "raw",
        "BGRX",
        0,
        1,
    )

    filepath = os.path.join(save_dir, f"page_{index:04d}.png")
    img.save(filepath)

    win32gui.DeleteObject(bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)

    return True, filepath


def bg_turn_page(hwnd, horizontal: bool):
    """PostMessage でバックグラウンドキー送信する"""
    import win32con
    import win32gui

    vk = win32con.VK_RIGHT if horizontal else win32con.VK_LEFT
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
    time.sleep(0.05)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0)


# ---------------------------------------------------------------------------
# 共通ユーティリティ
# ---------------------------------------------------------------------------

def detect_modal(screenshot_path: str, template_path: str, threshold: float = 0.8) -> bool:
    """スクリーンショット内にレビューモーダルを検出する"""
    if not os.path.exists(template_path):
        return False

    try:
        template = cv2.imread(template_path, cv2.IMREAD_COLOR)
        screenshot = cv2.imread(screenshot_path, cv2.IMREAD_COLOR)

        if template is None or screenshot is None:
            return False

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

        res = cv2.matchTemplate(curr, prev, cv2.TM_CCOEFF_NORMED)
        similarity = res[0][0]
        if similarity >= threshold:
            print(f"  -> 同一ページ検出 (similarity: {similarity:.4f}) — 最終ページと判断")
            return True
        return False
    except Exception:
        return False


# ---------------------------------------------------------------------------
# CLI / メインループ
# ---------------------------------------------------------------------------

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
        "--background",
        action="store_true",
        help="バックグラウンドモード（Kindleが背面でも動作。作業しながら実行可能）。",
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
    capture_mode = "バックグラウンド (Win32 API)" if args.background else "フォアグラウンド (pyautogui)"
    print(f"=== Kindle for PC 自動キャプチャ ===")
    print(f"  ページ送り: {mode}")
    print(f"  キャプチャ: {capture_mode}")
    print(f"  保存先: {save_dir}")
    print(f"  ページめくり待機: {args.delay}s")
    print(f"  最大ページ数: {args.max_pages}")

    has_template = os.path.exists(TEMPLATE_PATH) and not args.no_modal_detection
    if has_template:
        print(f"  モーダル検出: ON (threshold={args.threshold})")
    else:
        print(f"  モーダル検出: OFF (Ctrl+C で停止してください)")

    print()

    # --- バックグラウンドモード ---
    if args.background:
        hwnd, title = bg_find_kindle_hwnd()
        if hwnd is None:
            print("Error: Kindle for PC のウィンドウが見つかりません。")
            print("  Kindle for PC を起動し、本を開いた状態にしてください。")
            sys.exit(1)

        w, h = bg_get_window_size(hwnd)
        print(f"Kindle ウィンドウ検出: \"{title}\" ({w}x{h})")
        print("バックグラウンドで実行します。他のウィンドウで作業しても大丈夫です。")
        print()

        prev_path = None
        index = 1

        try:
            while index <= args.max_pages:
                print(f"--- Page {index} ---")

                hwnd_check, _ = bg_find_kindle_hwnd()
                if hwnd_check is None:
                    print("Error: Kindleウィンドウが見つかりません。終了します。")
                    break
                hwnd = hwnd_check

                success, filepath = bg_capture(hwnd, index, save_dir)
                if not success:
                    print("Error: スクリーンショットの取得に失敗しました。")
                    break

                if has_template and detect_modal(filepath, TEMPLATE_PATH, args.threshold):
                    print("レビューモーダルを検出しました。キャプチャを停止します。")
                    os.remove(filepath)
                    break

                if detect_duplicate(prev_path, filepath):
                    os.remove(filepath)
                    break

                prev_path = filepath
                time.sleep(0.3)

                bg_turn_page(hwnd, args.horizontal)
                index += 1
                time.sleep(args.delay)

        except KeyboardInterrupt:
            print("\n\nCtrl+C を検出しました。キャプチャを停止します。")

    # --- フォアグラウンドモード ---
    else:
        win = fg_find_kindle_window()
        if win is None:
            print("Error: Kindle for PC のウィンドウが見つかりません。")
            print("  Kindle for PC を起動し、本を開いた状態にしてください。")
            sys.exit(1)

        print(f"Kindle ウィンドウ検出: \"{win.title}\" ({win.width}x{win.height})")
        fg_activate_kindle(win)

        print("3秒後にキャプチャを開始します... Kindleの本の部分をクリックしてフォーカスを当ててください。")
        time.sleep(3)

        prev_path = None
        index = 1

        try:
            while index <= args.max_pages:
                print(f"--- Page {index} ---")

                win = fg_find_kindle_window()
                if win is None:
                    print("Error: Kindleウィンドウが見つかりません。終了します。")
                    break

                success, filepath = fg_capture(win, index, save_dir)
                if not success:
                    print("Error: スクリーンショットの取得に失敗しました。")
                    break

                if has_template and detect_modal(filepath, TEMPLATE_PATH, args.threshold):
                    print("レビューモーダルを検出しました。キャプチャを停止します。")
                    os.remove(filepath)
                    break

                if detect_duplicate(prev_path, filepath):
                    os.remove(filepath)
                    break

                prev_path = filepath
                time.sleep(0.3)

                fg_turn_page(args.horizontal)
                index += 1
                time.sleep(args.delay)

        except KeyboardInterrupt:
            print("\n\nCtrl+C を検出しました。キャプチャを停止します。")

    total = index - 1
    print(f"\n完了！ {total} ページをキャプチャしました。")
    print(f"保存先: {save_dir}")


if __name__ == "__main__":
    main()
