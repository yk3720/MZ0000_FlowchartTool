"""MZ0000_フローチャート作成(詳細版) 起動エントリポイント。

Powered by Auto (Cursor) (rev014)
Recommended Python Version: 3.14.2
達人級・体系的必然構成 (rev002) に準拠。
"""
import sys
import os

import pythoncom
import logging
from pathlib import Path

# 【インポート・ハイジーン (K-002)】パスの聖域化
# エントリポイント（main.py）の冒頭において、自身のディレクトリを sys.path の先頭に強制挿入
# これにより、PyInstaller等による配布環境（exe）や、異なるカレントディレクトリから実行された際でも、
# app パッケージの読み込み不全（ModuleNotFoundError）を根絶する
if not getattr(sys, 'frozen', False):
    # 開発環境: __file__ が存在する場合のみパスを追加
    if hasattr(sys.modules[__name__], '__file__') and __file__:
        sys.path.insert(0, str(Path(__file__).resolve().parent))

from app.utils.path_utils import ensure_environment
from app.utils.logging_config import setup_logger
from app.ui.main_window import FlowchartApp
from app.constants import APP_NAME, REVISION, AUTHOR


def main() -> None:
    """アプリケーションの起動。
    
    Raises:
        SystemExit: OS環境がWindowsでない場合、または致命的なエラーが発生した場合。
    """
    # 1. 実行環境の自己修復（ディレクトリ生成等）
    ensure_environment()
    
    # 2. ロガーの初期化
    logger = setup_logger()
    logger.info(f"starting_app | {APP_NAME} {REVISION} | by {AUTHOR}")
    
    # 3. OS環境チェック
    if os.name != "nt":
        logger.error("unsupported_os | Windows is required for Excel COM.")
        print("Error: Windows is required for Excel COM.")
        sys.exit(1)

    # 4. COMの初期化
    pythoncom.CoInitialize()
    
    try:
        # GUIの起動
        app = FlowchartApp()
        app.mainloop()
    except KeyboardInterrupt:
        logger.info("app_interrupted_by_user")
        print("アプリケーションが中断されました。")
    except Exception as e:
        logger.exception("unhandled_exception_at_main")
        print(f"致命的なエラーが発生しました: {e}")
    finally:
        # COMのクリーンアップ
        pythoncom.CoUninitialize()
        logger.info("app_terminated")


if __name__ == "__main__":
    main()
