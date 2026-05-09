"""FastAPI 应用：上传 → 异步校对 → 轮询 → 下载。

路由
----
- ``GET  /``                     上传页
- ``POST /jobs``                 接受 docx 上传，启动后台校对，返回 HTMX 片段
- ``GET  /jobs/{id}/status``     HTMX 轮询用，返回进度片段
- ``GET  /jobs/{id}/download``   下载结果 docx
- ``GET  /healthz``              健康检查

启动
----
::

    uv run uvicorn src.web.app:app --reload --port 8000
"""

from __future__ import annotations

import asyncio
import logging
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates

# 让 `src/` 内的模块可以裸名 import（沿用项目惯例：from proofreader_factory import …）
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

load_dotenv(PROJECT_ROOT / ".env")

from src.web.jobs import JobStatus, JobStore, run_proofread_job

logger = logging.getLogger(__name__)

# 校对模式 → 用户友好标签 / 描述（与 main.py 保持一致）
MODE_OPTIONS = [
    ("comments", "💬 批注模式", "在 Word 审阅窗格显示建议，原文不变"),
    ("revisions", "✏️ 修订模式", "用样式显示删除/插入"),
    ("track_changes", "🔄 跟踪更改", "原生 Word 跟踪更改标记"),
    ("enhanced", "🌟 增强模式", "跟踪更改 + 详细批注（推荐）"),
]
ALLOWED_MODES = {m[0] for m in MODE_OPTIONS}
MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB

TEMPLATES_DIR = Path(__file__).parent / "templates"
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


def _get_api_key() -> str:
    """复用 main.py 的优先级 licloud > claude > openai。"""
    return (
        os.getenv("LICLOUD_TOKEN")
        or os.getenv("LLM_API_TOKEN")
        or os.getenv("CLAUDE_SESSION_TOKEN")
        or os.getenv("ANTHROPIC_API_KEY")
        or os.getenv("OPENAI_API_KEY")
        or ""
    )


def _detect_provider() -> str:
    if os.getenv("AI_PROVIDER"):
        return os.getenv("AI_PROVIDER", "").lower()
    if os.getenv("LICLOUD_TOKEN") or os.getenv("LLM_API_TOKEN"):
        return "licloud"
    if os.getenv("CLAUDE_SESSION_TOKEN") or os.getenv("ANTHROPIC_API_KEY"):
        return "claude"
    return "openai"


def create_app(*, job_store: JobStore | None = None) -> FastAPI:
    """工厂函数，便于测试时注入自定义 store。"""
    app = FastAPI(title="AI 校对助手", version="1.1.0")
    app.state.job_store = job_store or JobStore()
    # 持有后台任务的强引用避免被GC（RUF006）
    app.state.background_tasks = set()

    @app.get("/healthz")
    async def healthz() -> dict[str, str]:
        return {"status": "ok"}

    @app.get("/", response_class=HTMLResponse)
    async def index(request: Request) -> HTMLResponse:
        return templates.TemplateResponse(
            request,
            "index.html",
            {
                "modes": MODE_OPTIONS,
                "provider": _detect_provider(),
                "has_key": bool(_get_api_key()),
            },
        )

    @app.post("/jobs", response_class=HTMLResponse)
    async def create_job(
        request: Request,
        file: UploadFile = File(...),
        mode: str = Form("comments"),
    ) -> HTMLResponse:
        if mode not in ALLOWED_MODES:
            raise HTTPException(status_code=400, detail=f"未知模式: {mode}")

        # 仅放行 .docx
        if not (file.filename or "").lower().endswith(".docx"):
            raise HTTPException(status_code=400, detail="仅支持 .docx 文件")

        api_key = _get_api_key()
        if not api_key:
            raise HTTPException(
                status_code=503,
                detail="服务端未配置 API 密钥（LICLOUD_TOKEN / CLAUDE_SESSION_TOKEN / OPENAI_API_KEY 任一）",
            )

        contents = await file.read()
        if len(contents) > MAX_UPLOAD_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"文件过大，限 {MAX_UPLOAD_BYTES // 1024 // 1024} MB",
            )
        if not contents:
            raise HTTPException(status_code=400, detail="空文件")

        store: JobStore = request.app.state.job_store
        job = await store.create(
            filename=file.filename or "upload.docx",
            mode=mode,
            source_bytes=contents,
        )
        # 后台跑校对（强引用避免被GC回收，完成后自动从set里移除）
        task = asyncio.create_task(run_proofread_job(store, job, api_key=api_key))
        request.app.state.background_tasks.add(task)
        task.add_done_callback(request.app.state.background_tasks.discard)

        return templates.TemplateResponse(
            request,
            "_job_status.html",
            {"job": job},
        )

    @app.get("/jobs/{job_id}/status", response_class=HTMLResponse)
    async def job_status(request: Request, job_id: str) -> HTMLResponse:
        store: JobStore = request.app.state.job_store
        job = await store.get(job_id)
        if job is None:
            raise HTTPException(status_code=404, detail="任务不存在或已清理")
        return templates.TemplateResponse(
            request,
            "_job_status.html",
            {"job": job},
        )

    @app.get("/jobs/{job_id}/download")
    async def download(request: Request, job_id: str) -> FileResponse:
        store: JobStore = request.app.state.job_store
        job = await store.get(job_id)
        if job is None or job.status != JobStatus.SUCCEEDED:
            raise HTTPException(status_code=404, detail="结果不存在或尚未完成")
        if not job.output_path.exists():
            raise HTTPException(status_code=410, detail="结果文件已被清理")
        return FileResponse(
            path=job.output_path,
            filename=f"校对_{job.filename}",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    return app


app = create_app()
