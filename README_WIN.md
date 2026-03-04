# Kindle 自動スクリーンショット — Windows版

macOS版（AppleScript + screencapture）をWindows 11 + Kindle for PCで動作するように移植したバージョンです。

## 必要環境

- Windows 10/11
- Python 3.11+
- Kindle for PC（Microsoft Store版 or Amazon公式版）

## セットアップ

### 1. Python インストール

[python.org](https://www.python.org/downloads/) から Python 3.11 以上をダウンロード。

> **重要**: インストーラーで **「Add Python to PATH」にチェック** を入れてください。

### 2. 仮想環境 & 依存パッケージ

```bash
cd C:/Users/user/kindle-automation
python -m venv venv
venv/Scripts/activate
pip install -r requirements.txt
```

## 使い方

### スクリーンショットの撮影

```bash
# 仮想環境を有効化
venv/Scripts/activate

# 縦書き本（←キーでページ送り）— デフォルト
python kindle_capture.py

# 横書き本（→キーでページ送り）
python kindle_capture.py --horizontal

# オプション
python kindle_capture.py --delay 2.0        # ページめくり待機時間（秒）
python kindle_capture.py --max-pages 500     # 最大ページ数
python kindle_capture.py --output-dir ./out  # 保存先を変更
python kindle_capture.py --no-modal-detection  # モーダル検出無効化
```

**実行前の準備:**
1. Kindle for PC で本を開く
2. **全画面表示**にする（F11 または表示メニュー）
3. スクリプト実行後、3秒以内にKindleのページ部分をクリックしてフォーカスを当てる

**停止方法:**
- テンプレート画像がある場合 → レビューモーダル検出時に自動停止
- 同一ページが2回続いた場合 → 最終ページと判断して自動停止
- 手動停止 → `Ctrl+C`

### PDF生成

```bash
# デフォルト（~/Desktop/KindleScreenshots → ~/Desktop/KindleBook.pdf）
python make_pdf_win.py

# ディレクトリとファイル名を指定
python make_pdf_win.py --input-dir ~/Desktop/KindleScreenshots_Horizontal --output book.pdf
```

## テンプレート画像の作成方法（モーダル自動停止用）

Kindleは本の最後にレビューを促すモーダルを表示します。このモーダルを検出して自動停止するには、テンプレート画像が必要です。

1. Kindleで何か本を最後まで読み、レビューモーダルが表示された状態にする
2. `Win + Shift + S` でモーダル部分だけを切り取ってスクリーンショット
3. `review_modal_template_v2.png` として `kindle-automation/` フォルダに保存

> テンプレートがなくてもスクリプトは動作します（Ctrl+C で手動停止）。

## macOS版との違い

| 機能 | macOS版 | Windows版 |
|------|---------|-----------|
| ウィンドウ操作 | AppleScript | pygetwindow |
| スクリーンショット | screencapture | pyautogui |
| キー送信 | osascript (key code) | pyautogui.press() |
| ページ送り方向 | 別ファイル (v7 / horizontal) | `--horizontal` フラグで切替 |
| PDF変換 | ImageMagick + ocrmypdf | img2pdf（無劣化） |
| 重複ページ検出 | なし | あり（自動停止） |
| OCR | ocrmypdf (日本語) | なし（将来対応予定） |

## トラブルシューティング

### 「Kindleのウィンドウが見つかりません」

- Kindle for PC が起動しているか確認
- タスクバーでKindleがアクティブになっているか確認
- ウィンドウタイトルに "Kindle" が含まれていない場合は検出できません

### スクリーンショットが黒い / 空白

- Kindle for PC のハードウェアアクセラレーションを無効にしてみてください
  - Kindle > ツール > オプション > 一般 > 「ハードウェアアクセラレーションを使用」のチェックを外す

### ページが送れない

- スクリプト実行後、Kindleのページ部分をクリックしてフォーカスを当ててください
- `--delay` の値を大きくしてみてください（例: `--delay 2.5`）
