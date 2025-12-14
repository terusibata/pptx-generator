"""
PowerPoint生成のためのスキーマ定義
"""
from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


# =============================================================================
# Enums
# =============================================================================

class PlaceholderType(str, Enum):
    """プレースホルダーの種類"""
    TITLE = "TITLE"
    CENTER_TITLE = "CENTER_TITLE"
    SUBTITLE = "SUBTITLE"
    BODY = "BODY"
    OBJECT = "OBJECT"
    PICTURE = "PICTURE"
    CHART = "CHART"
    TABLE = "TABLE"
    FOOTER = "FOOTER"
    DATE = "DATE"
    SLIDE_NUMBER = "SLIDE_NUMBER"


class ContentType(str, Enum):
    """コンテンツの種類"""
    TEXT = "text"
    BULLETS = "bullets"
    IMAGE = "image"
    TABLE = "table"
    CHART = "chart"


class ChartType(str, Enum):
    """グラフの種類"""
    BAR = "bar"
    COLUMN = "column"
    LINE = "line"
    PIE = "pie"
    DOUGHNUT = "doughnut"
    AREA = "area"


# =============================================================================
# Template Metadata Schemas (テンプレート解析結果)
# =============================================================================

class PlaceholderMeta(BaseModel):
    """プレースホルダーのメタデータ"""
    idx: int = Field(..., description="プレースホルダーのインデックス")
    type: PlaceholderType | None = Field(None, description="プレースホルダーの種類")
    name: str = Field(..., description="プレースホルダーの名前")
    left: float = Field(..., description="左位置（インチ）")
    top: float = Field(..., description="上位置（インチ）")
    width: float = Field(..., description="幅（インチ）")
    height: float = Field(..., description="高さ（インチ）")
    has_text_frame: bool = Field(True, description="テキストフレームを持つか")


class LayoutMeta(BaseModel):
    """レイアウトのメタデータ"""
    index: int = Field(..., description="レイアウトのインデックス")
    name: str = Field(..., description="レイアウトの名前")
    description: str | None = Field(None, description="レイアウトの説明（AI用）")
    usage_hint: str | None = Field(None, description="使用時のヒント（AI用）")
    placeholders: list[PlaceholderMeta] = Field(
        default_factory=list, 
        description="プレースホルダー一覧"
    )


class TemplateMeta(BaseModel):
    """テンプレートのメタデータ"""
    id: str = Field(..., description="テンプレートID")
    name: str = Field(..., description="テンプレート名")
    file_name: str = Field(..., description="ファイル名")
    description: str | None = Field(None, description="テンプレートの説明")
    slide_width: int = Field(..., description="スライドの幅（EMU）")
    slide_height: int = Field(..., description="スライドの高さ（EMU）")
    layouts: list[LayoutMeta] = Field(
        default_factory=list,
        description="利用可能なレイアウト一覧"
    )
    custom_config: TemplateCustomConfig | None = Field(
        None, 
        description="カスタム設定"
    )


class TemplateCustomConfig(BaseModel):
    """テンプレートのカスタム設定"""
    default_font: str | None = Field(None, description="デフォルトフォント")
    title_font: str | None = Field(None, description="タイトル用フォント")
    accent_color: str | None = Field(None, description="アクセントカラー（RGB hex）")
    layout_aliases: dict[str, int] = Field(
        default_factory=dict,
        description="レイアウトのエイリアス（名前→インデックス）"
    )


# =============================================================================
# Content Schemas (AIからの入力)
# =============================================================================

class TextStyle(BaseModel):
    """テキストのスタイル設定"""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_size: float | None = Field(None, description="フォントサイズ（pt）")
    font_name: str | None = Field(None, description="フォント名")
    color: str | None = Field(None, description="RGB色（例: 'FF0000'）")
    theme_color: str | None = Field(None, description="テーマカラー名")


class ParagraphContent(BaseModel):
    """段落コンテンツ"""
    text: str = Field(..., description="テキスト内容")
    style: TextStyle | None = Field(None, description="スタイル設定")
    bullet: bool = Field(False, description="箇条書きにするか")
    level: int = Field(0, ge=0, le=8, description="インデントレベル")
    alignment: str | None = Field(None, description="配置（LEFT/CENTER/RIGHT）")


class TextContent(BaseModel):
    """テキストコンテンツ"""
    type: ContentType = ContentType.TEXT
    paragraphs: list[ParagraphContent] = Field(..., description="段落リスト")


class BulletContent(BaseModel):
    """箇条書きコンテンツ"""
    type: ContentType = ContentType.BULLETS
    items: list[str | ParagraphContent] = Field(..., description="箇条書き項目")
    style: TextStyle | None = Field(None, description="共通スタイル")


class ImageContent(BaseModel):
    """画像コンテンツ"""
    type: ContentType = ContentType.IMAGE
    source: str = Field(..., description="画像パスまたはURL")
    alt_text: str | None = Field(None, description="代替テキスト")


class TableCell(BaseModel):
    """テーブルのセル"""
    text: str = Field(..., description="セルのテキスト")
    style: TextStyle | None = Field(None, description="スタイル")
    colspan: int = Field(1, ge=1, description="結合する列数")
    rowspan: int = Field(1, ge=1, description="結合する行数")


class TableContent(BaseModel):
    """テーブルコンテンツ"""
    type: ContentType = ContentType.TABLE
    headers: list[str] | None = Field(None, description="ヘッダー行")
    rows: list[list[str | TableCell]] = Field(..., description="データ行")
    style: TableStyle | None = Field(None, description="テーブルスタイル")


class TableStyle(BaseModel):
    """テーブルのスタイル"""
    header_bg_color: str | None = Field(None, description="ヘッダー背景色")
    header_text_color: str | None = Field(None, description="ヘッダー文字色")
    alternate_row_color: str | None = Field(None, description="交互行の背景色")
    border_color: str | None = Field(None, description="罫線色")


class ChartDataSeries(BaseModel):
    """チャートのデータ系列"""
    name: str = Field(..., description="系列名")
    values: list[float] = Field(..., description="値リスト")
    color: str | None = Field(None, description="系列の色")


class ChartContent(BaseModel):
    """チャートコンテンツ"""
    type: ContentType = ContentType.CHART
    chart_type: ChartType = Field(..., description="グラフの種類")
    title: str | None = Field(None, description="グラフタイトル")
    categories: list[str] = Field(..., description="カテゴリ（X軸ラベル）")
    series: list[ChartDataSeries] = Field(..., description="データ系列")
    show_legend: bool = Field(True, description="凡例を表示するか")


# コンテンツの Union 型
SlideContent = TextContent | BulletContent | ImageContent | TableContent | ChartContent


class PlaceholderContent(BaseModel):
    """プレースホルダーへのコンテンツ割り当て"""
    placeholder_name: str | None = Field(None, description="プレースホルダー名で指定")
    placeholder_idx: int | None = Field(None, description="プレースホルダーインデックスで指定")
    placeholder_type: PlaceholderType | None = Field(None, description="種類で指定")
    content: SlideContent | str | list[str] = Field(..., description="コンテンツ")


class SlideDefinition(BaseModel):
    """スライド定義"""
    layout_name: str | None = Field(None, description="レイアウト名で指定")
    layout_index: int | None = Field(None, description="レイアウトインデックスで指定")
    contents: dict[str, SlideContent | str | list[str]] = Field(
        default_factory=dict,
        description="プレースホルダー名 → コンテンツのマッピング"
    )
    speaker_notes: str | None = Field(None, description="スピーカーノート")


# =============================================================================
# API Request/Response Schemas
# =============================================================================

class GenerateRequest(BaseModel):
    """PPTX生成リクエスト"""
    session_id: str = Field(..., description="セッションID")
    template_id: str = Field(..., description="使用するテンプレートID")
    slides: list[SlideDefinition] = Field(..., description="スライド定義リスト")
    options: GenerateOptions | None = Field(None, description="生成オプション")


class GenerateOptions(BaseModel):
    """生成オプション"""
    generate_preview: bool = Field(True, description="プレビューPDFを生成するか")
    generate_thumbnails: bool = Field(True, description="サムネイル画像を生成するか")
    thumbnail_width: int = Field(400, description="サムネイルの幅（px）")


class GenerateResponse(BaseModel):
    """PPTX生成レスポンス"""
    session_id: str
    pptx_path: str
    preview_path: str | None = None
    thumbnail_paths: list[str] = Field(default_factory=list)
    slide_count: int
    warnings: list[str] = Field(default_factory=list)


class TemplateListResponse(BaseModel):
    """テンプレート一覧レスポンス"""
    templates: list[TemplateMeta]


class TemplateUploadResponse(BaseModel):
    """テンプレートアップロードレスポンス"""
    template_id: str
    meta: TemplateMeta
    message: str


# =============================================================================
# AI向け簡易スキーマ（Difyに渡す用）
# =============================================================================

class LayoutForAI(BaseModel):
    """AI向けのレイアウト情報（シンプル版）"""
    name: str
    index: int
    description: str | None = None
    placeholders: list[str] = Field(
        default_factory=list,
        description="プレースホルダー名のリスト"
    )


class TemplateForAI(BaseModel):
    """AI向けのテンプレート情報（シンプル版）"""
    id: str
    name: str
    layouts: list[LayoutForAI]
    
    @classmethod
    def from_meta(cls, meta: TemplateMeta) -> "TemplateForAI":
        """TemplateMetaからAI向け情報を生成"""
        return cls(
            id=meta.id,
            name=meta.name,
            layouts=[
                LayoutForAI(
                    name=layout.name,
                    index=layout.index,
                    description=layout.description,
                    placeholders=[p.name for p in layout.placeholders]
                )
                for layout in meta.layouts
            ]
        )
