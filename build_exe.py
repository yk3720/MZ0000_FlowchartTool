# -*- coding: utf-8 -*-
"""
ビルドスクリプト
MZ0000_FlowchartTool rev014

プログラム名: MZ0000_FlowchartTool
操作AI名: Auto (Cursor)
"""

import os
import subprocess
import sys
from pathlib import Path

# 作業ディレクトリをスクリプトのディレクトリに設定
script_dir = Path(__file__).resolve().parent
os.chdir(script_dir)

# PyInstallerコマンドの構築
pyinstaller_cmd = [
    "pyinstaller",
    "--onefile",
    "--windowed",
    "--name=MZ0000_FlowchartTool_rev014",
    "--icon=NONE",  # アイコンは未指定（必要に応じて追加）
    "--add-data=仕様・管理;仕様・管理",  # 仕様書ディレクトリを含める
    "main.py",
]

print("=" * 60)
print("PyInstallerによる実行ファイルのビルドを開始します")
print("=" * 60)
print(f"作業ディレクトリ: {script_dir}")
print(f"コマンド: {' '.join(pyinstaller_cmd)}")
print("=" * 60)

try:
    # PyInstallerの実行
    result = subprocess.run(
        pyinstaller_cmd,
        cwd=script_dir,
        check=True,
    )
    
    print("=" * 60)
    print("ビルドが完了しました")
    print(f"実行ファイル: {script_dir / 'dist' / 'MZ0000_FlowchartTool_rev014.exe'}")
    print("=" * 60)
    
except subprocess.CalledProcessError as e:
    print(f"【エラー】ビルドに失敗しました: {e}")
    sys.exit(1)
except FileNotFoundError:
    print("【エラー】PyInstallerが見つかりません。")
    print("インストール方法: pip install pyinstaller")
    sys.exit(1)
