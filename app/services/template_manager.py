"""
テンプレート管理サービス
テンプレートファイルとメタデータの保存・取得を管理
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import TYPE_CHECKING

from app.schemas import TemplateMeta, TemplateForAI
from app.services.template_analyzer import (
    analyze_template,
    enrich_template_meta,
    DEFAULT_LAYOUT_DESCRIPTIONS,
    DEFAULT_LAYOUT_HINTS,
)

if TYPE_CHECKING:
    from fastapi import UploadFile


class TemplateManager:
    """
    テンプレート管理クラス
    
    ディレクトリ構造:
        base_dir/
        ├── templates/
        │   ├── template_id_1.pptx
        │   └── template_id_2.pptx
        └── meta/
            ├── template_id_1.json
            └── template_id_2.json
    """
    
    def __init__(self, base_dir: str | Path):
        """
        Args:
            base_dir: ベースディレクトリ
        """
        self.base_dir = Path(base_dir)
        self.templates_dir = self.base_dir / "templates"
        self.meta_dir = self.base_dir / "meta"
        
        # ディレクトリを作成
        self.templates_dir.mkdir(parents=True, exist_ok=True)
        self.meta_dir.mkdir(parents=True, exist_ok=True)
    
    def get_template_path(self, template_id: str) -> Path:
        """テンプレートファイルのパスを取得"""
        # まず直接パスを確認
        direct_path = self.templates_dir / f"{template_id}.pptx"
        if direct_path.exists():
            return direct_path
        
        # メタデータからファイル名を取得
        meta = self.get_template_meta(template_id)
        if meta:
            return self.templates_dir / meta.file_name
        
        raise FileNotFoundError(f"Template not found: {template_id}")
    
    def get_meta_path(self, template_id: str) -> Path:
        """メタデータファイルのパスを取得"""
        return self.meta_dir / f"{template_id}.json"
    
    def list_templates(self) -> list[TemplateMeta]:
        """利用可能なテンプレート一覧を取得"""
        templates = []
        
        for meta_file in self.meta_dir.glob("*.json"):
            try:
                with open(meta_file, encoding="utf-8") as f:
                    data = json.load(f)
                    templates.append(TemplateMeta.model_validate(data))
            except Exception:
                continue
        
        return templates
    
    def list_templates_for_ai(self) -> list[TemplateForAI]:
        """AI向けのテンプレート一覧を取得"""
        templates = self.list_templates()
        return [TemplateForAI.from_meta(t) for t in templates]
    
    def get_template_meta(self, template_id: str) -> TemplateMeta | None:
        """テンプレートのメタデータを取得"""
        meta_path = self.get_meta_path(template_id)
        
        if not meta_path.exists():
            return None
        
        with open(meta_path, encoding="utf-8") as f:
            data = json.load(f)
            return TemplateMeta.model_validate(data)
    
    def get_template_for_ai(self, template_id: str) -> TemplateForAI | None:
        """AI向けのテンプレート情報を取得"""
        meta = self.get_template_meta(template_id)
        if meta:
            return TemplateForAI.from_meta(meta)
        return None
    
    def register_template(
        self,
        pptx_path: str | Path,
        template_id: str | None = None,
        name: str | None = None,
        description: str | None = None,
        layout_descriptions: dict[str, str] | None = None,
        layout_hints: dict[str, str] | None = None,
        layout_aliases: dict[str, int] | None = None,
        use_default_descriptions: bool = True,
    ) -> TemplateMeta:
        """
        テンプレートを登録
        
        Args:
            pptx_path: PPTXファイルのパス
            template_id: テンプレートID（省略時は自動生成）
            name: テンプレート名
            description: テンプレートの説明
            layout_descriptions: レイアウトの説明
            layout_hints: レイアウトの使用ヒント
            layout_aliases: レイアウトのエイリアス
            use_default_descriptions: デフォルトの説明を使用するか
        
        Returns:
            登録されたテンプレートのメタデータ
        """
        pptx_path = Path(pptx_path)
        
        if not pptx_path.exists():
            raise FileNotFoundError(f"Template file not found: {pptx_path}")
        
        # テンプレートを解析
        meta = analyze_template(pptx_path, template_id)
        
        # 名前と説明を更新
        if name:
            meta = meta.model_copy(update={"name": name})
        if description:
            meta = meta.model_copy(update={"description": description})
        
        # デフォルトの説明をマージ
        merged_descriptions = {}
        merged_hints = {}
        
        if use_default_descriptions:
            merged_descriptions.update(DEFAULT_LAYOUT_DESCRIPTIONS)
            merged_hints.update(DEFAULT_LAYOUT_HINTS)
        
        if layout_descriptions:
            merged_descriptions.update(layout_descriptions)
        if layout_hints:
            merged_hints.update(layout_hints)
        
        # 説明を追加
        meta = enrich_template_meta(
            meta,
            layout_descriptions=merged_descriptions,
            layout_hints=merged_hints,
            layout_aliases=layout_aliases,
        )
        
        # ファイルをコピー
        dest_path = self.templates_dir / pptx_path.name
        if dest_path != pptx_path:
            shutil.copy2(pptx_path, dest_path)
        
        # メタデータを保存
        meta_path = self.get_meta_path(meta.id)
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta.model_dump(), f, ensure_ascii=False, indent=2)
        
        return meta
    
    async def upload_template(
        self,
        file: "UploadFile",
        template_id: str | None = None,
        name: str | None = None,
        description: str | None = None,
    ) -> TemplateMeta:
        """
        アップロードされたテンプレートを登録
        
        Args:
            file: アップロードされたファイル
            template_id: テンプレートID
            name: テンプレート名
            description: テンプレートの説明
        
        Returns:
            登録されたテンプレートのメタデータ
        """
        # 一時ファイルに保存
        temp_path = self.templates_dir / (file.filename or "uploaded.pptx")
        
        with open(temp_path, "wb") as f:
            content = await file.read()
            f.write(content)
        
        try:
            return self.register_template(
                temp_path,
                template_id=template_id,
                name=name or Path(file.filename or "").stem,
                description=description,
            )
        except Exception:
            # エラー時は一時ファイルを削除
            if temp_path.exists():
                temp_path.unlink()
            raise
    
    def update_template_meta(
        self,
        template_id: str,
        name: str | None = None,
        description: str | None = None,
        layout_descriptions: dict[str, str] | None = None,
        layout_hints: dict[str, str] | None = None,
        layout_aliases: dict[str, int] | None = None,
    ) -> TemplateMeta:
        """
        テンプレートのメタデータを更新
        """
        meta = self.get_template_meta(template_id)
        
        if meta is None:
            raise FileNotFoundError(f"Template not found: {template_id}")
        
        # 基本情報を更新
        updates = {}
        if name:
            updates["name"] = name
        if description:
            updates["description"] = description
        
        if updates:
            meta = meta.model_copy(update=updates)
        
        # 説明を更新
        if layout_descriptions or layout_hints or layout_aliases:
            meta = enrich_template_meta(
                meta,
                layout_descriptions=layout_descriptions,
                layout_hints=layout_hints,
                layout_aliases=layout_aliases,
            )
        
        # 保存
        meta_path = self.get_meta_path(template_id)
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta.model_dump(), f, ensure_ascii=False, indent=2)
        
        return meta
    
    def delete_template(self, template_id: str) -> bool:
        """テンプレートを削除"""
        meta = self.get_template_meta(template_id)
        
        if meta is None:
            return False
        
        # ファイルを削除
        template_path = self.templates_dir / meta.file_name
        if template_path.exists():
            template_path.unlink()
        
        # メタデータを削除
        meta_path = self.get_meta_path(template_id)
        if meta_path.exists():
            meta_path.unlink()
        
        return True
