"""
PowerPoint生成サービス
テンプレートとAI出力のJSONからPPTXを生成
"""
from __future__ import annotations

import copy
import re
from pathlib import Path
from typing import TYPE_CHECKING, Any

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from app.schemas import (
    BulletContent,
    ChartContent,
    ChartType,
    ContentType,
    ImageContent,
    ParagraphContent,
    PlaceholderType,
    SlideContent,
    SlideDefinition,
    TableContent,
    TemplateMeta,
    TextContent,
    TextStyle,
)

if TYPE_CHECKING:
    from pptx.shapes.autoshape import Shape
    from pptx.shapes.placeholder import SlidePlaceholder
    from pptx.slide import Slide, SlideLayout
    from pptx.text.text import TextFrame


# =============================================================================
# 定数・マッピング
# =============================================================================

ALIGNMENT_MAP: dict[str, int] = {
    "LEFT": PP_ALIGN.LEFT,
    "CENTER": PP_ALIGN.CENTER,
    "RIGHT": PP_ALIGN.RIGHT,
    "JUSTIFY": PP_ALIGN.JUSTIFY,
}

CHART_TYPE_MAP: dict[ChartType, int] = {
    ChartType.BAR: XL_CHART_TYPE.BAR_CLUSTERED,
    ChartType.COLUMN: XL_CHART_TYPE.COLUMN_CLUSTERED,
    ChartType.LINE: XL_CHART_TYPE.LINE,
    ChartType.PIE: XL_CHART_TYPE.PIE,
    ChartType.DOUGHNUT: XL_CHART_TYPE.DOUGHNUT,
    ChartType.AREA: XL_CHART_TYPE.AREA,
}

THEME_COLOR_MAP: dict[str, int] = {
    "DARK_1": MSO_THEME_COLOR.DARK_1,
    "LIGHT_1": MSO_THEME_COLOR.LIGHT_1,
    "DARK_2": MSO_THEME_COLOR.DARK_2,
    "LIGHT_2": MSO_THEME_COLOR.LIGHT_2,
    "ACCENT_1": MSO_THEME_COLOR.ACCENT_1,
    "ACCENT_2": MSO_THEME_COLOR.ACCENT_2,
    "ACCENT_3": MSO_THEME_COLOR.ACCENT_3,
    "ACCENT_4": MSO_THEME_COLOR.ACCENT_4,
    "ACCENT_5": MSO_THEME_COLOR.ACCENT_5,
    "ACCENT_6": MSO_THEME_COLOR.ACCENT_6,
}


# =============================================================================
# ヘルパー関数
# =============================================================================

def _parse_rgb_color(color_str: str) -> RGBColor:
    """RGB文字列をRGBColorオブジェクトに変換"""
    color_str = color_str.lstrip("#")
    if len(color_str) != 6:
        raise ValueError(f"Invalid color format: {color_str}")
    r = int(color_str[0:2], 16)
    g = int(color_str[2:4], 16)
    b = int(color_str[4:6], 16)
    return RGBColor(r, g, b)


def _apply_text_style(run, style: TextStyle | None) -> None:
    """テキストランにスタイルを適用"""
    if style is None:
        return
    
    if style.bold:
        run.font.bold = True
    if style.italic:
        run.font.italic = True
    if style.underline:
        run.font.underline = True
    if style.font_size:
        run.font.size = Pt(style.font_size)
    if style.font_name:
        run.font.name = style.font_name
    if style.color:
        run.font.color.rgb = _parse_rgb_color(style.color)
    elif style.theme_color and style.theme_color in THEME_COLOR_MAP:
        run.font.color.theme_color = THEME_COLOR_MAP[style.theme_color]


def _find_placeholder(
    slide: "Slide",
    name: str | None = None,
    idx: int | None = None,
    ph_type: PlaceholderType | None = None,
) -> "SlidePlaceholder | None":
    """
    スライドからプレースホルダーを検索
    優先順位: idx > name > ph_type
    """
    for shape in slide.placeholders:
        if not hasattr(shape, "placeholder_format"):
            continue
        
        ph_format = shape.placeholder_format
        
        # インデックスで検索
        if idx is not None and ph_format.idx == idx:
            return shape
        
        # 名前で検索（部分一致）
        if name is not None:
            shape_name = shape.name or ""
            if name.lower() in shape_name.lower() or shape_name.lower() in name.lower():
                return shape
        
        # タイプで検索
        if ph_type is not None:
            type_map = {
                PlaceholderType.TITLE: PP_PLACEHOLDER.TITLE,
                PlaceholderType.CENTER_TITLE: PP_PLACEHOLDER.CENTER_TITLE,
                PlaceholderType.SUBTITLE: PP_PLACEHOLDER.SUBTITLE,
                PlaceholderType.BODY: PP_PLACEHOLDER.BODY,
            }
            if ph_type in type_map and ph_format.type == type_map[ph_type]:
                return shape
    
    return None


def _find_layout_by_name_or_index(
    prs: Presentation,
    name: str | None = None,
    index: int | None = None,
    meta: TemplateMeta | None = None,
) -> "SlideLayout":
    """
    レイアウトを名前またはインデックスで検索
    """
    # インデックス指定の場合
    if index is not None:
        if 0 <= index < len(prs.slide_layouts):
            return prs.slide_layouts[index]
        raise ValueError(f"Layout index {index} out of range (0-{len(prs.slide_layouts)-1})")
    
    # 名前指定の場合
    if name is not None:
        # エイリアス検索
        if meta and meta.custom_config and meta.custom_config.layout_aliases:
            if name in meta.custom_config.layout_aliases:
                alias_index = meta.custom_config.layout_aliases[name]
                return prs.slide_layouts[alias_index]
        
        # 名前で検索（完全一致 → 部分一致）
        for layout in prs.slide_layouts:
            if layout.name == name:
                return layout
        
        # 部分一致
        name_lower = name.lower()
        for layout in prs.slide_layouts:
            if name_lower in (layout.name or "").lower():
                return layout
        
        raise ValueError(f"Layout not found: {name}")
    
    # デフォルトは最初のレイアウト
    return prs.slide_layouts[0]


# =============================================================================
# コンテンツ挿入関数
# =============================================================================

def _insert_text_content(
    text_frame: "TextFrame",
    content: TextContent,
    clear_existing: bool = True,
) -> None:
    """テキストコンテンツをテキストフレームに挿入"""
    if clear_existing:
        text_frame.clear()
    
    for i, para_content in enumerate(content.paragraphs):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        # テキスト設定
        run = p.add_run()
        run.text = para_content.text
        
        # スタイル適用
        _apply_text_style(run, para_content.style)
        
        # 段落プロパティ
        if para_content.bullet:
            p.level = para_content.level
        
        if para_content.alignment and para_content.alignment in ALIGNMENT_MAP:
            p.alignment = ALIGNMENT_MAP[para_content.alignment]


def _insert_bullet_content(
    text_frame: "TextFrame",
    content: BulletContent,
    clear_existing: bool = True,
) -> None:
    """箇条書きコンテンツをテキストフレームに挿入"""
    if clear_existing:
        text_frame.clear()
    
    for i, item in enumerate(content.items):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        p.level = 0  # 箇条書きレベル
        
        if isinstance(item, str):
            run = p.add_run()
            run.text = item
            _apply_text_style(run, content.style)
        else:
            # ParagraphContent の場合
            run = p.add_run()
            run.text = item.text
            _apply_text_style(run, item.style or content.style)
            p.level = item.level


def _insert_simple_text(
    text_frame: "TextFrame",
    text: str,
    clear_existing: bool = True,
) -> None:
    """単純なテキストを挿入"""
    if clear_existing:
        text_frame.clear()
    
    p = text_frame.paragraphs[0]
    run = p.add_run()
    run.text = text


def _insert_simple_bullets(
    text_frame: "TextFrame",
    items: list[str],
    clear_existing: bool = True,
) -> None:
    """単純な箇条書きリストを挿入"""
    if clear_existing:
        text_frame.clear()
    
    for i, item in enumerate(items):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        
        p.level = 0
        run = p.add_run()
        run.text = item


def _insert_table(
    slide: "Slide",
    placeholder: "SlidePlaceholder",
    content: TableContent,
) -> None:
    """テーブルコンテンツを挿入"""
    rows_count = len(content.rows)
    if content.headers:
        rows_count += 1
    
    cols_count = len(content.rows[0]) if content.rows else len(content.headers or [])
    
    # テーブルを作成
    table = slide.shapes.add_table(
        rows_count,
        cols_count,
        placeholder.left,
        placeholder.top,
        placeholder.width,
        placeholder.height,
    ).table
    
    row_offset = 0
    
    # ヘッダー行
    if content.headers:
        for col_idx, header in enumerate(content.headers):
            cell = table.cell(0, col_idx)
            cell.text = header
            if content.style and content.style.header_bg_color:
                cell.fill.solid()
                cell.fill.fore_color.rgb = _parse_rgb_color(content.style.header_bg_color)
        row_offset = 1
    
    # データ行
    for row_idx, row in enumerate(content.rows):
        for col_idx, cell_data in enumerate(row):
            cell = table.cell(row_idx + row_offset, col_idx)
            if isinstance(cell_data, str):
                cell.text = cell_data
            else:
                cell.text = cell_data.text


def _insert_chart(
    slide: "Slide",
    placeholder: "SlidePlaceholder",
    content: ChartContent,
) -> None:
    """チャートコンテンツを挿入"""
    chart_data = CategoryChartData()
    chart_data.categories = content.categories
    
    for series in content.series:
        chart_data.add_series(series.name, series.values)
    
    chart_type = CHART_TYPE_MAP.get(content.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
    
    chart = slide.shapes.add_chart(
        chart_type,
        placeholder.left,
        placeholder.top,
        placeholder.width,
        placeholder.height,
        chart_data,
    ).chart
    
    if content.title:
        chart.has_title = True
        chart.chart_title.text_frame.text = content.title
    
    chart.has_legend = content.show_legend


def _insert_image(
    slide: "Slide",
    placeholder: "SlidePlaceholder",
    content: ImageContent,
) -> None:
    """画像コンテンツを挿入"""
    image_path = Path(content.source)
    
    if not image_path.exists():
        # URL の場合は後で対応（ここではスキップ）
        return
    
    slide.shapes.add_picture(
        str(image_path),
        placeholder.left,
        placeholder.top,
        placeholder.width,
        placeholder.height,
    )


# =============================================================================
# メイン生成クラス
# =============================================================================

class PptxGenerator:
    """
    PowerPoint生成クラス
    
    Usage:
        generator = PptxGenerator(template_path, template_meta)
        generator.add_slide(slide_definition)
        generator.save(output_path)
    """
    
    def __init__(
        self,
        template_path: str | Path,
        template_meta: TemplateMeta | None = None,
    ):
        """
        Args:
            template_path: テンプレートPPTXのパス
            template_meta: テンプレートメタデータ（レイアウト検索に使用）
        """
        self.template_path = Path(template_path)
        self.template_meta = template_meta
        self.prs = Presentation(template_path)
        self.warnings: list[str] = []
        
        # テンプレートに既存スライドがあれば削除
        self._remove_existing_slides()
    
    def _remove_existing_slides(self) -> None:
        """テンプレートの既存スライドを削除"""
        # スライドを逆順で削除（インデックスずれ防止）
        slide_ids = [slide.slide_id for slide in self.prs.slides]
        for slide_id in slide_ids:
            rId = self.prs.part.get_rId(self.prs.slides.get(slide_id).part)
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[self.prs.slides._sldIdLst.index(
                self.prs.slides._sldIdLst.sldId_lst[
                    [s.slide_id for s in self.prs.slides].index(slide_id)
                ]
            )]
    
    def add_slide(self, definition: SlideDefinition) -> "Slide":
        """
        スライドを追加
        
        Args:
            definition: スライド定義
        
        Returns:
            追加されたスライド
        """
        # レイアウトを取得
        layout = _find_layout_by_name_or_index(
            self.prs,
            name=definition.layout_name,
            index=definition.layout_index,
            meta=self.template_meta,
        )
        
        # スライドを追加
        slide = self.prs.slides.add_slide(layout)
        
        # コンテンツを挿入
        for placeholder_key, content in definition.contents.items():
            self._insert_content(slide, placeholder_key, content)
        
        # スピーカーノート
        if definition.speaker_notes:
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = definition.speaker_notes
        
        return slide
    
    def _insert_content(
        self,
        slide: "Slide",
        placeholder_key: str,
        content: SlideContent | str | list[str],
    ) -> None:
        """
        プレースホルダーにコンテンツを挿入
        
        Args:
            slide: スライド
            placeholder_key: プレースホルダーのキー（名前、インデックス、タイプ）
            content: 挿入するコンテンツ
        """
        # プレースホルダーを検索
        placeholder = self._resolve_placeholder(slide, placeholder_key)
        
        if placeholder is None:
            self.warnings.append(f"Placeholder not found: {placeholder_key}")
            return
        
        if not hasattr(placeholder, "text_frame"):
            self.warnings.append(f"Placeholder has no text frame: {placeholder_key}")
            return
        
        text_frame = placeholder.text_frame
        
        # コンテンツタイプに応じて挿入
        if isinstance(content, str):
            _insert_simple_text(text_frame, content)
        
        elif isinstance(content, list):
            # 文字列のリストは箇条書きとして扱う
            _insert_simple_bullets(text_frame, content)
        
        elif isinstance(content, TextContent):
            _insert_text_content(text_frame, content)
        
        elif isinstance(content, BulletContent):
            _insert_bullet_content(text_frame, content)
        
        elif isinstance(content, TableContent):
            _insert_table(slide, placeholder, content)
        
        elif isinstance(content, ChartContent):
            _insert_chart(slide, placeholder, content)
        
        elif isinstance(content, ImageContent):
            _insert_image(slide, placeholder, content)
    
    def _resolve_placeholder(
        self,
        slide: "Slide",
        key: str,
    ) -> "SlidePlaceholder | None":
        """
        キーからプレースホルダーを解決
        
        キーの形式:
        - "タイトル" → 名前で検索
        - "idx:0" → インデックスで検索
        - "type:TITLE" → タイプで検索
        """
        # インデックス指定
        if key.startswith("idx:"):
            idx = int(key[4:])
            return _find_placeholder(slide, idx=idx)
        
        # タイプ指定
        if key.startswith("type:"):
            type_str = key[5:]
            try:
                ph_type = PlaceholderType(type_str)
                return _find_placeholder(slide, ph_type=ph_type)
            except ValueError:
                pass
        
        # 名前で検索
        return _find_placeholder(slide, name=key)
    
    def save(self, output_path: str | Path) -> Path:
        """
        PPTXを保存
        
        Args:
            output_path: 出力パス
        
        Returns:
            保存されたファイルのパス
        """
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(output_path)
        return output_path
    
    @property
    def slide_count(self) -> int:
        """スライド数を取得"""
        return len(self.prs.slides)


# =============================================================================
# 便利関数
# =============================================================================

def generate_pptx(
    template_path: str | Path,
    slides: list[SlideDefinition],
    output_path: str | Path,
    template_meta: TemplateMeta | None = None,
) -> tuple[Path, list[str]]:
    """
    PPTXを生成する便利関数
    
    Args:
        template_path: テンプレートPPTXのパス
        slides: スライド定義のリスト
        output_path: 出力パス
        template_meta: テンプレートメタデータ
    
    Returns:
        (出力パス, 警告リスト)
    """
    generator = PptxGenerator(template_path, template_meta)
    
    for slide_def in slides:
        generator.add_slide(slide_def)
    
    saved_path = generator.save(output_path)
    
    return saved_path, generator.warnings


def generate_from_ai_json(
    template_path: str | Path,
    ai_output: dict[str, Any],
    output_path: str | Path,
    template_meta: TemplateMeta | None = None,
) -> tuple[Path, list[str]]:
    """
    AIからのJSON出力を直接処理してPPTXを生成
    
    Args:
        template_path: テンプレートPPTXのパス
        ai_output: AIからのJSON出力（slides配列を含む）
        output_path: 出力パス
        template_meta: テンプレートメタデータ
    
    Returns:
        (出力パス, 警告リスト)
    """
    slides_data = ai_output.get("slides", [])
    
    slides = []
    for slide_data in slides_data:
        # AIからの出力をSlideDefinitionに変換
        slide_def = SlideDefinition(
            layout_name=slide_data.get("layoutName") or slide_data.get("layout_name"),
            layout_index=slide_data.get("layoutIndex") or slide_data.get("layout_index"),
            contents=_normalize_contents(slide_data.get("content", {})),
            speaker_notes=slide_data.get("speakerNotes") or slide_data.get("speaker_notes"),
        )
        slides.append(slide_def)
    
    return generate_pptx(template_path, slides, output_path, template_meta)


def _normalize_contents(
    contents: dict[str, Any],
) -> dict[str, SlideContent | str | list[str]]:
    """
    AIからのコンテンツをスキーマに正規化
    """
    normalized = {}
    
    for key, value in contents.items():
        if isinstance(value, str):
            normalized[key] = value
        elif isinstance(value, list):
            # 文字列のリストはそのまま
            if all(isinstance(v, str) for v in value):
                normalized[key] = value
            else:
                # 複雑な構造は TextContent に変換
                paragraphs = []
                for item in value:
                    if isinstance(item, str):
                        paragraphs.append(ParagraphContent(text=item))
                    elif isinstance(item, dict):
                        paragraphs.append(ParagraphContent(**item))
                normalized[key] = TextContent(paragraphs=paragraphs)
        elif isinstance(value, dict):
            # typeフィールドで判定
            content_type = value.get("type")
            if content_type == "text":
                normalized[key] = TextContent(**value)
            elif content_type == "bullets":
                normalized[key] = BulletContent(**value)
            elif content_type == "table":
                normalized[key] = TableContent(**value)
            elif content_type == "chart":
                normalized[key] = ChartContent(**value)
            elif content_type == "image":
                normalized[key] = ImageContent(**value)
            else:
                # typeがない場合は paragraphs の有無で判定
                if "paragraphs" in value:
                    normalized[key] = TextContent(**value)
                else:
                    # 単純なkey-valueとして処理
                    normalized[key] = str(value)
    
    return normalized
