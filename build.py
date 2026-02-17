"""PyInstallerビルド用スクリプト

使い方（Windows コマンドプロンプト）:
    py -3.11 build.py
"""

import subprocess
import sys


def main():
    # PyInstallerがインストールされているか確認
    try:
        import PyInstaller
    except ImportError:
        print("PyInstallerをインストールしています...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "pdf_converter.spec",
        "--noconfirm",
    ]

    print("ビルドを開始します...")
    print(f"コマンド: {' '.join(cmd)}")
    subprocess.check_call(cmd)

    print()
    print("=" * 50)
    print("ビルド完了!")
    print("出力先: dist/pdf-converter/")
    print("実行: dist/pdf-converter/pdf_converter.exe")
    print("=" * 50)


if __name__ == "__main__":
    main()
