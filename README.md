# PPTX Generator API

AIによるPowerPoint生成のためのバックエンドAPI。会社のスライドマスター付きテンプレートを活用し、ブランドガイドライン準拠のプレゼンテーションを生成します。

## 特徴

- **テンプレートベース生成**: スライドマスター・テーマを完全に継承
- **柔軟なコンテンツ指定**: 単純テキスト、箇条書き、テーブル、チャートに対応
- **セッション管理**: 修正依頼に対応した会話継続機能
- **プレビュー生成**: PDF・サムネイル画像の自動生成

## クイックスタート

### Docker を使用

```bash
# 起動
docker-compose up -d

# API確認
curl http://localhost:8000/health
```

### ローカル実行

```bash
# 依存パッケージのインストール
pip install -e .

# 追加依存（プレビュー生成用）
# Ubuntu/Debian
sudo apt-get install libreoffice

# 起動
uvicorn app.main:app --reload
```

## API エンドポイント

### テンプレート管理

| エンドポイント | メソッド | 説明 |
|--------------|---------|------|
| `/api/templates` | GET | テンプレート一覧 |
| `/api/templates/for-ai` | GET | AI向けテンプレート一覧（シンプル版） |
| `/api/templates/{id}` | GET | テンプレート詳細 |
| `/api/templates/{id}/for-ai` | GET | AI向けテンプレート詳細 |
| `/api/templates/upload` | POST | テンプレートアップロード |
| `/api/templates/{id}` | PUT | テンプレート更新 |
| `/api/templates/{id}` | DELETE | テンプレート削除 |

### PPTX生成

| エンドポイント | メソッド | 説明 |
|--------------|---------|------|
| `/api/generate` | POST | PPTX生成 |

### セッション管理

| エンドポイント | メソッド | 説明 |
|--------------|---------|------|
| `/api/sessions/{id}` | GET | セッション情報 |
| `/api/sessions/{id}/files/pptx` | GET | PPTXダウンロード |
| `/api/sessions/{id}/files/preview` | GET | プレビューPDF |
| `/api/sessions/{id}/files/thumbnails/{name}` | GET | サムネイル画像 |
| `/api/sessions/{id}` | DELETE | セッション削除 |

## 使用例

### 1. テンプレートの登録

```bash
curl -X POST http://localhost:8000/api/templates/upload \
  -F "file=@company_template.pptx" \
  -F "name=会社標準テンプレート" \
  -F "description=2024年版の標準テンプレート"
```

### 2. AI向けテンプレート情報の取得

```bash
curl http://localhost:8000/api/templates/company_template/for-ai
```

レスポンス例:
```json
{
  "id": "company_template",
  "name": "会社標準テンプレート",
  "layouts": [
    {
      "name": "タイトル スライド",
      "index": 0,
      "placeholders": ["タイトル", "サブタイトル"]
    },
    {
      "name": "タイトルとコンテンツ",
      "index": 1,
      "placeholders": ["タイトル", "コンテンツ"]
    }
  ]
}
```

### 3. PPTX生成

```bash
curl -X POST http://localhost:8000/api/generate \
  -H "Content-Type: application/json" \
  -d '{
    "template_id": "company_template",
    "slides": [
      {
        "layoutName": "タイトル スライド",
        "content": {
          "タイトル": "2024年度 営業戦略",
          "サブタイトル": "営業本部 | 2024年4月"
        }
      },
      {
        "layoutName": "タイトルとコンテンツ",
        "content": {
          "タイトル": "本日のアジェンダ",
          "コンテンツ": [
            "市場環境の分析",
            "重点施策の説明",
            "KPI目標の設定"
          ]
        }
      }
    ]
  }'
```

レスポンス例:
```json
{
  "session_id": "abc123...",
  "pptx_url": "/api/sessions/abc123.../files/pptx",
  "preview_url": "/api/sessions/abc123.../files/preview",
  "thumbnail_urls": [
    "/api/sessions/abc123.../files/thumbnails/slide_1.png",
    "/api/sessions/abc123.../files/thumbnails/slide_2.png"
  ],
  "slide_count": 2,
  "warnings": []
}
```

### 4. 修正依頼（セッション継続）

```bash
curl -X POST http://localhost:8000/api/generate \
  -H "Content-Type: application/json" \
  -d '{
    "session_id": "abc123...",
    "template_id": "company_template",
    "user_input": "2ページ目のタイトルを変更",
    "slides": [
      {
        "layoutName": "タイトル スライド",
        "content": {
          "タイトル": "2024年度 営業戦略",
          "サブタイトル": "営業本部 | 2024年4月"
        }
      },
      {
        "layoutName": "タイトルとコンテンツ",
        "content": {
          "タイトル": "目次",
          "コンテンツ": [
            "市場環境の分析",
            "重点施策の説明",
            "KPI目標の設定"
          ]
        }
      }
    ]
  }'
```

## コンテンツ形式

### 単純テキスト

```json
"タイトル": "シンプルなテキスト"
```

### 箇条書き（配列）

```json
"コンテンツ": [
  "項目1",
  "項目2",
  "項目3"
]
```

### 詳細テキスト

```json
"本文": {
  "type": "text",
  "paragraphs": [
    {
      "text": "太字のテキスト",
      "style": {"bold": true}
    },
    {
      "text": "通常のテキスト"
    }
  ]
}
```

### テーブル

```json
"データ": {
  "type": "table",
  "headers": ["項目", "2023年", "2024年"],
  "rows": [
    ["売上", "100億円", "120億円"],
    ["利益", "10億円", "15億円"]
  ]
}
```

### チャート

```json
"グラフ": {
  "type": "chart",
  "chart_type": "column",
  "title": "売上推移",
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "2023年", "values": [100, 120, 110, 130]},
    {"name": "2024年", "values": [120, 140, 150, 160]}
  ]
}
```

## Dify との統合

詳細は `docs/dify_prompt_template.md` を参照してください。

### 基本フロー

1. フロントエンドからDifyにリクエスト
2. Difyが本APIからテンプレート情報を取得
3. LLMがスライド構成JSONを生成
4. Difyが本APIにPPTX生成をリクエスト
5. フロントエンドにダウンロードURL/プレビューを返却

## ディレクトリ構造

```
pptx-generator/
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI アプリケーション
│   ├── schemas.py           # Pydantic スキーマ
│   └── services/
│       ├── __init__.py
│       ├── template_analyzer.py  # テンプレート解析
│       ├── template_manager.py   # テンプレート管理
│       ├── pptx_generator.py     # PPTX生成
│       └── session_manager.py    # セッション管理
├── data/                    # ランタイムデータ
│   ├── templates/           # テンプレートファイル
│   ├── meta/                # メタデータJSON
│   └── sessions/            # セッションデータ
├── docs/
│   └── dify_prompt_template.md
├── Dockerfile
├── docker-compose.yml
├── pyproject.toml
└── README.md
```

## 環境変数

| 変数名 | デフォルト | 説明 |
|--------|-----------|------|
| `PPTX_DATA_DIR` | `./data` | データディレクトリ |
| `PPTX_CORS_ORIGINS` | `["*"]` | CORS許可オリジン |
| `PPTX_CLEANUP_INTERVAL_HOURS` | `24` | セッションクリーンアップ間隔 |

## ライセンス

MIT
