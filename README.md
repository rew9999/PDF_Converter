# PDF変換ツール

PDFファイルを画像（PNG/JPG）に変換するWebアプリケーションです。

## 使い方

### 開発環境での起動

```bash
pip install -r requirements.txt
python main.py
```

ブラウザで http://localhost:5000 が自動的に開きます。

### 操作手順

1. PDFファイルをドラッグ＆ドロップまたはファイル選択でアップロード
2. 変換形式（PNG / JPG）と解像度（150 / 200 / 300 dpi）を選択
3. 「変換開始」ボタンをクリック
4. 変換完了後、出力先フォルダに画像ファイルが保存されます

### 設定

`config.json` で以下の設定を変更できます:

- `default_output_dir`: デフォルトの出力先フォルダ
- `default_format`: デフォルトの変換形式（png / jpg）
- `default_dpi`: デフォルトの解像度
- `max_file_size_mb`: アップロード可能な最大ファイルサイズ（MB）
- `port`: サーバーのポート番号
- `auto_open_browser`: 起動時にブラウザを自動で開くか

## 技術スタック

- Python 3.11+
- FastAPI + Uvicorn
- PyMuPDF (fitz)
- Pillow
