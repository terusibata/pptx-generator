FROM python:3.11-slim

WORKDIR /app

# システム依存パッケージのインストール
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# Python依存パッケージのインストール
COPY pyproject.toml .
RUN pip install --no-cache-dir pip --upgrade && \
    pip install --no-cache-dir .

# アプリケーションコードのコピー
COPY app/ app/

# データディレクトリの作成
RUN mkdir -p /app/data

# 環境変数
ENV PPTX_DATA_DIR=/app/data
ENV PYTHONUNBUFFERED=1

# ポート
EXPOSE 8000

# 起動コマンド
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
