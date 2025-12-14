"""
セッション管理サービス
生成セッションの管理とファイル出力を担当
"""
from __future__ import annotations

import json
import shutil
import subprocess
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field


class SessionState(BaseModel):
    """セッション状態"""
    session_id: str
    template_id: str
    created_at: datetime = Field(default_factory=datetime.now)
    updated_at: datetime = Field(default_factory=datetime.now)
    current_slides: list[dict[str, Any]] = Field(default_factory=list)
    generation_history: list[GenerationRecord] = Field(default_factory=list)
    files: SessionFiles = Field(default_factory=lambda: SessionFiles())


class GenerationRecord(BaseModel):
    """生成履歴レコード"""
    id: str = Field(default_factory=lambda: str(uuid.uuid4())[:8])
    timestamp: datetime = Field(default_factory=datetime.now)
    type: str  # "initial" or "modification"
    user_input: str
    slides_snapshot: list[dict[str, Any]]


class SessionFiles(BaseModel):
    """セッションのファイル情報"""
    pptx_path: str | None = None
    preview_path: str | None = None
    thumbnail_paths: list[str] = Field(default_factory=list)


class SessionManager:
    """
    セッション管理クラス
    
    ディレクトリ構造:
        base_dir/
        └── sessions/
            └── {session_id}/
                ├── state.json
                ├── output.pptx
                ├── preview.pdf
                └── thumbnails/
                    ├── slide_1.png
                    └── slide_2.png
    """
    
    def __init__(self, base_dir: str | Path):
        """
        Args:
            base_dir: ベースディレクトリ
        """
        self.base_dir = Path(base_dir)
        self.sessions_dir = self.base_dir / "sessions"
        self.sessions_dir.mkdir(parents=True, exist_ok=True)
    
    def _get_session_dir(self, session_id: str) -> Path:
        """セッションディレクトリのパスを取得"""
        return self.sessions_dir / session_id
    
    def _get_state_path(self, session_id: str) -> Path:
        """状態ファイルのパスを取得"""
        return self._get_session_dir(session_id) / "state.json"
    
    def create_session(self, template_id: str) -> SessionState:
        """
        新しいセッションを作成
        
        Args:
            template_id: 使用するテンプレートID
        
        Returns:
            作成されたセッション状態
        """
        session_id = str(uuid.uuid4())
        session_dir = self._get_session_dir(session_id)
        session_dir.mkdir(parents=True, exist_ok=True)
        
        # サムネイルディレクトリを作成
        (session_dir / "thumbnails").mkdir(exist_ok=True)
        
        state = SessionState(
            session_id=session_id,
            template_id=template_id,
        )
        
        self._save_state(state)
        
        return state
    
    def get_session(self, session_id: str) -> SessionState | None:
        """セッション状態を取得"""
        state_path = self._get_state_path(session_id)
        
        if not state_path.exists():
            return None
        
        with open(state_path, encoding="utf-8") as f:
            data = json.load(f)
            return SessionState.model_validate(data)
    
    def _save_state(self, state: SessionState) -> None:
        """セッション状態を保存"""
        state_path = self._get_state_path(state.session_id)
        state.updated_at = datetime.now()
        
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(state.model_dump(mode="json"), f, ensure_ascii=False, indent=2, default=str)
    
    def update_session(
        self,
        session_id: str,
        slides: list[dict[str, Any]],
        user_input: str,
        generation_type: str = "modification",
    ) -> SessionState:
        """
        セッションを更新
        
        Args:
            session_id: セッションID
            slides: 新しいスライド構成
            user_input: ユーザー入力
            generation_type: 生成タイプ
        
        Returns:
            更新されたセッション状態
        """
        state = self.get_session(session_id)
        
        if state is None:
            raise ValueError(f"Session not found: {session_id}")
        
        # 履歴を追加
        record = GenerationRecord(
            type=generation_type,
            user_input=user_input,
            slides_snapshot=slides,
        )
        state.generation_history.append(record)
        
        # 現在のスライドを更新
        state.current_slides = slides
        
        self._save_state(state)
        
        return state
    
    def update_files(
        self,
        session_id: str,
        pptx_path: str | None = None,
        preview_path: str | None = None,
        thumbnail_paths: list[str] | None = None,
    ) -> SessionState:
        """ファイル情報を更新"""
        state = self.get_session(session_id)
        
        if state is None:
            raise ValueError(f"Session not found: {session_id}")
        
        if pptx_path:
            state.files.pptx_path = pptx_path
        if preview_path:
            state.files.preview_path = preview_path
        if thumbnail_paths is not None:
            state.files.thumbnail_paths = thumbnail_paths
        
        self._save_state(state)
        
        return state
    
    def get_output_paths(self, session_id: str) -> dict[str, Path]:
        """出力ファイルのパスを取得"""
        session_dir = self._get_session_dir(session_id)
        
        return {
            "pptx": session_dir / "output.pptx",
            "preview": session_dir / "preview.pdf",
            "thumbnails_dir": session_dir / "thumbnails",
        }
    
    def delete_session(self, session_id: str) -> bool:
        """セッションを削除"""
        session_dir = self._get_session_dir(session_id)
        
        if not session_dir.exists():
            return False
        
        shutil.rmtree(session_dir)
        return True
    
    def cleanup_old_sessions(self, max_age_hours: int = 24) -> int:
        """
        古いセッションをクリーンアップ
        
        Args:
            max_age_hours: 保持する最大時間（時間）
        
        Returns:
            削除されたセッション数
        """
        deleted = 0
        cutoff = datetime.now().timestamp() - (max_age_hours * 3600)
        
        for session_dir in self.sessions_dir.iterdir():
            if not session_dir.is_dir():
                continue
            
            state_path = session_dir / "state.json"
            if state_path.exists():
                try:
                    with open(state_path) as f:
                        data = json.load(f)
                        updated_at = datetime.fromisoformat(data.get("updated_at", ""))
                        if updated_at.timestamp() < cutoff:
                            shutil.rmtree(session_dir)
                            deleted += 1
                except Exception:
                    pass
        
        return deleted


class PreviewGenerator:
    """プレビュー生成クラス"""
    
    @staticmethod
    def generate_pdf(pptx_path: Path, output_path: Path) -> bool:
        """
        PPTXからPDFを生成
        
        LibreOfficeを使用
        """
        try:
            result = subprocess.run(
                [
                    "soffice",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(output_path.parent),
                    str(pptx_path),
                ],
                capture_output=True,
                timeout=60,
            )
            
            # 出力ファイル名を修正（LibreOfficeはファイル名を変更する）
            generated_pdf = output_path.parent / f"{pptx_path.stem}.pdf"
            if generated_pdf.exists() and generated_pdf != output_path:
                generated_pdf.rename(output_path)
            
            return output_path.exists()
        except Exception:
            return False
    
    @staticmethod
    def generate_thumbnails(
        pdf_path: Path,
        output_dir: Path,
        width: int = 400,
    ) -> list[Path]:
        """
        PDFからサムネイル画像を生成
        
        pdftoppmを使用
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            # DPIを計算（幅400pxの場合、約150DPI）
            dpi = max(72, min(300, int(width * 150 / 400)))
            
            result = subprocess.run(
                [
                    "pdftoppm",
                    "-png",
                    "-r", str(dpi),
                    str(pdf_path),
                    str(output_dir / "slide"),
                ],
                capture_output=True,
                timeout=120,
            )
            
            # 生成されたファイルを収集
            thumbnails = sorted(output_dir.glob("slide-*.png"))
            
            # ファイル名を正規化（slide-1.png → slide_1.png）
            normalized = []
            for thumb in thumbnails:
                new_name = thumb.name.replace("-", "_")
                new_path = thumb.parent / new_name
                if new_path != thumb:
                    thumb.rename(new_path)
                normalized.append(new_path)
            
            return normalized
        except Exception:
            return []
    
    @classmethod
    def generate_all(
        cls,
        pptx_path: Path,
        session_dir: Path,
        generate_pdf: bool = True,
        generate_thumbnails: bool = True,
        thumbnail_width: int = 400,
    ) -> tuple[Path | None, list[Path]]:
        """
        全てのプレビューを生成
        
        Returns:
            (PDFパス, サムネイルパスのリスト)
        """
        pdf_path = None
        thumbnails = []
        
        if generate_pdf:
            pdf_output = session_dir / "preview.pdf"
            if cls.generate_pdf(pptx_path, pdf_output):
                pdf_path = pdf_output
        
        if generate_thumbnails and pdf_path:
            thumbnails_dir = session_dir / "thumbnails"
            thumbnails = cls.generate_thumbnails(
                pdf_path,
                thumbnails_dir,
                width=thumbnail_width,
            )
        
        return pdf_path, thumbnails
