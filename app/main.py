"""
PowerPoint生成API
"""
from __future__ import annotations

from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, UploadFile, File, Form, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    """アプリケーション設定"""
    app_name: str = "PPTX Generator API"
    data_dir: str = "./data"
    cors_origins: list[str] = ["*"]
    cleanup_interval_hours: int = 24
    
    class Config:
        env_prefix = "PPTX_"


settings = Settings()


@asynccontextmanager
async def lifespan(app: FastAPI):
    """アプリケーションのライフサイクル管理"""
    # 起動時の初期化
    data_dir = Path(settings.data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    
    yield
    
    # 終了時のクリーンアップ


app = FastAPI(
    title=settings.app_name,
    description="AIによるPowerPoint生成のためのAPI",
    version="0.1.0",
    lifespan=lifespan,
)

# CORS設定
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =============================================================================
# 依存性注入
# =============================================================================

def get_template_manager():
    """テンプレートマネージャーを取得"""
    from app.services.template_manager import TemplateManager
    return TemplateManager(settings.data_dir)


def get_session_manager():
    """セッションマネージャーを取得"""
    from app.services.session_manager import SessionManager
    return SessionManager(settings.data_dir)


# =============================================================================
# テンプレート関連エンドポイント
# =============================================================================

@app.get("/api/templates")
async def list_templates():
    """利用可能なテンプレート一覧を取得"""
    manager = get_template_manager()
    templates = manager.list_templates()
    return {"templates": [t.model_dump() for t in templates]}


@app.get("/api/templates/for-ai")
async def list_templates_for_ai():
    """AI向けのテンプレート一覧を取得（シンプル版）"""
    manager = get_template_manager()
    templates = manager.list_templates_for_ai()
    return {"templates": [t.model_dump() for t in templates]}


@app.get("/api/templates/{template_id}")
async def get_template(template_id: str):
    """テンプレートの詳細を取得"""
    manager = get_template_manager()
    meta = manager.get_template_meta(template_id)
    
    if meta is None:
        raise HTTPException(status_code=404, detail="Template not found")
    
    return meta.model_dump()


@app.get("/api/templates/{template_id}/for-ai")
async def get_template_for_ai(template_id: str):
    """AI向けのテンプレート情報を取得"""
    manager = get_template_manager()
    template = manager.get_template_for_ai(template_id)
    
    if template is None:
        raise HTTPException(status_code=404, detail="Template not found")
    
    return template.model_dump()


@app.post("/api/templates/upload")
async def upload_template(
    file: UploadFile = File(...),
    template_id: str | None = Form(None),
    name: str | None = Form(None),
    description: str | None = Form(None),
):
    """テンプレートをアップロード"""
    if not file.filename or not file.filename.endswith((".pptx", ".potx")):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Only .pptx and .potx files are allowed.",
        )
    
    manager = get_template_manager()
    
    try:
        meta = await manager.upload_template(
            file,
            template_id=template_id,
            name=name,
            description=description,
        )
        return {
            "message": "Template uploaded successfully",
            "template_id": meta.id,
            "meta": meta.model_dump(),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/templates/{template_id}/analyze")
async def analyze_template_endpoint(template_id: str):
    """テンプレートを再解析"""
    manager = get_template_manager()
    
    try:
        template_path = manager.get_template_path(template_id)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Template not found")
    
    from app.services.template_analyzer import analyze_template
    meta = analyze_template(template_path, template_id)
    
    return meta.model_dump()


@app.put("/api/templates/{template_id}")
async def update_template(template_id: str, data: dict[str, Any]):
    """テンプレートのメタデータを更新"""
    manager = get_template_manager()
    
    try:
        meta = manager.update_template_meta(
            template_id,
            name=data.get("name"),
            description=data.get("description"),
            layout_descriptions=data.get("layout_descriptions"),
            layout_hints=data.get("layout_hints"),
            layout_aliases=data.get("layout_aliases"),
        )
        return meta.model_dump()
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Template not found")


@app.delete("/api/templates/{template_id}")
async def delete_template(template_id: str):
    """テンプレートを削除"""
    manager = get_template_manager()
    
    if manager.delete_template(template_id):
        return {"message": "Template deleted successfully"}
    else:
        raise HTTPException(status_code=404, detail="Template not found")


# =============================================================================
# 生成関連エンドポイント
# =============================================================================

@app.post("/api/generate")
async def generate_pptx(data: dict[str, Any], background_tasks: BackgroundTasks):
    """
    PPTXを生成
    
    Request Body:
        {
            "session_id": "optional - auto-generated if not provided",
            "template_id": "required",
            "slides": [...],
            "options": {
                "generate_preview": true,
                "generate_thumbnails": true,
                "thumbnail_width": 400
            }
        }
    """
    from app.services.pptx_generator import generate_from_ai_json
    from app.services.session_manager import PreviewGenerator
    
    template_manager = get_template_manager()
    session_manager = get_session_manager()
    
    # テンプレート確認
    template_id = data.get("template_id")
    if not template_id:
        raise HTTPException(status_code=400, detail="template_id is required")
    
    meta = template_manager.get_template_meta(template_id)
    if meta is None:
        raise HTTPException(status_code=404, detail="Template not found")
    
    try:
        template_path = template_manager.get_template_path(template_id)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Template file not found")
    
    # セッション作成または取得
    session_id = data.get("session_id")
    if session_id:
        session = session_manager.get_session(session_id)
        if session is None:
            session = session_manager.create_session(template_id)
    else:
        session = session_manager.create_session(template_id)
    
    # 出力パス
    paths = session_manager.get_output_paths(session.session_id)
    
    # PPTX生成
    try:
        output_path, warnings = generate_from_ai_json(
            template_path=template_path,
            ai_output=data,
            output_path=paths["pptx"],
            template_meta=meta,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Generation failed: {str(e)}")
    
    # オプション
    options = data.get("options", {})
    generate_preview = options.get("generate_preview", True)
    generate_thumbnails = options.get("generate_thumbnails", True)
    thumbnail_width = options.get("thumbnail_width", 400)
    
    # プレビュー生成
    preview_path = None
    thumbnail_paths = []
    
    if generate_preview or generate_thumbnails:
        session_dir = session_manager._get_session_dir(session.session_id)
        preview_path, thumbnail_paths = PreviewGenerator.generate_all(
            output_path,
            session_dir,
            generate_pdf=generate_preview,
            generate_thumbnails=generate_thumbnails,
            thumbnail_width=thumbnail_width,
        )
    
    # セッション更新
    session_manager.update_session(
        session.session_id,
        slides=data.get("slides", []),
        user_input=data.get("user_input", ""),
        generation_type="initial" if not data.get("session_id") else "modification",
    )
    
    session_manager.update_files(
        session.session_id,
        pptx_path=str(output_path),
        preview_path=str(preview_path) if preview_path else None,
        thumbnail_paths=[str(p) for p in thumbnail_paths],
    )
    
    # レスポンス
    from pptx import Presentation
    prs = Presentation(output_path)
    
    return {
        "session_id": session.session_id,
        "pptx_url": f"/api/sessions/{session.session_id}/files/pptx",
        "preview_url": f"/api/sessions/{session.session_id}/files/preview" if preview_path else None,
        "thumbnail_urls": [
            f"/api/sessions/{session.session_id}/files/thumbnails/{p.name}"
            for p in thumbnail_paths
        ],
        "slide_count": len(prs.slides),
        "warnings": warnings,
    }


# =============================================================================
# セッション関連エンドポイント
# =============================================================================

@app.get("/api/sessions/{session_id}")
async def get_session(session_id: str):
    """セッション情報を取得"""
    manager = get_session_manager()
    session = manager.get_session(session_id)
    
    if session is None:
        raise HTTPException(status_code=404, detail="Session not found")
    
    return session.model_dump(mode="json")


@app.get("/api/sessions/{session_id}/files/pptx")
async def download_pptx(session_id: str):
    """PPTXファイルをダウンロード"""
    manager = get_session_manager()
    paths = manager.get_output_paths(session_id)
    
    if not paths["pptx"].exists():
        raise HTTPException(status_code=404, detail="PPTX file not found")
    
    return FileResponse(
        paths["pptx"],
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"presentation_{session_id}.pptx",
    )


@app.get("/api/sessions/{session_id}/files/preview")
async def get_preview(session_id: str):
    """プレビューPDFを取得"""
    manager = get_session_manager()
    paths = manager.get_output_paths(session_id)
    
    if not paths["preview"].exists():
        raise HTTPException(status_code=404, detail="Preview not found")
    
    return FileResponse(
        paths["preview"],
        media_type="application/pdf",
    )


@app.get("/api/sessions/{session_id}/files/thumbnails/{filename}")
async def get_thumbnail(session_id: str, filename: str):
    """サムネイル画像を取得"""
    manager = get_session_manager()
    paths = manager.get_output_paths(session_id)
    
    thumbnail_path = paths["thumbnails_dir"] / filename
    
    if not thumbnail_path.exists():
        raise HTTPException(status_code=404, detail="Thumbnail not found")
    
    return FileResponse(
        thumbnail_path,
        media_type="image/png",
    )


@app.delete("/api/sessions/{session_id}")
async def delete_session(session_id: str):
    """セッションを削除"""
    manager = get_session_manager()
    
    if manager.delete_session(session_id):
        return {"message": "Session deleted successfully"}
    else:
        raise HTTPException(status_code=404, detail="Session not found")


@app.post("/api/sessions/cleanup")
async def cleanup_sessions(max_age_hours: int = 24):
    """古いセッションをクリーンアップ"""
    manager = get_session_manager()
    deleted = manager.cleanup_old_sessions(max_age_hours)
    return {"deleted_count": deleted}


# =============================================================================
# ヘルスチェック
# =============================================================================

@app.get("/health")
async def health_check():
    """ヘルスチェック"""
    return {"status": "healthy"}


@app.get("/")
async def root():
    """ルート"""
    return {
        "name": settings.app_name,
        "version": "0.1.0",
        "docs": "/docs",
    }
