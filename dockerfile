# Pythonの公式イメージをベースにする
FROM python:3.9-slim-buster

# apt-get updateとpoppler-utilsのインストール
RUN apt-get update && apt-get install -y poppler-utils \
    # 不要になったaptキャッシュをクリーンアップしてイメージサイズを削減
    && rm -rf /var/lib/apt/lists/*

# 作業ディレクトリを設定
WORKDIR /app

# requirements.txtをコピーして依存関係をインストール
COPY requirements.txt .
RUN pip install -r requirements.txt

# アプリケーションコードをコピー
COPY . .

# アプリケーションを実行するコマンドを指定 (あなたのアプリのエントリポイントに合わせて変更してください)
CMD ["python", "app.py"]
