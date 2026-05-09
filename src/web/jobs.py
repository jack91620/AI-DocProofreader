"""校对任务调度（内存版）。

设计意图
========

Web UI 端到端流程是"上传 → 异步校对 → 轮询进度 → 下载结果"。校对本身可能
耗时数十秒到几分钟（长文档分块处理 + LLM 调用），HTTP 请求不能同步等待。

这里用最朴素的方案：
- 内存 dict 存任务（多 worker 部署时换 Redis；单进程开发足够）
- ``asyncio.create_task`` 跑后台协程
- 阻塞的校对调用包在 ``loop.run_in_executor`` 里，避免阻塞事件循环

不引入 Celery/RQ 这种重量级，因为本工具偏向"单机/小团队"。
"""

from __future__ import annotations

import asyncio
import logging
import shutil
import tempfile
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime
from enum import StrEnum
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


class JobStatus(StrEnum):
    PENDING = "pending"
    RUNNING = "running"
    SUCCEEDED = "succeeded"
    FAILED = "failed"


@dataclass
class Job:
    """单个校对任务的全状态。"""

    id: str
    filename: str  # 用户上传时的原文件名
    mode: str  # comments / revisions / track_changes / enhanced
    input_path: Path  # 服务端临时存储的输入 docx
    output_path: Path  # 校对结果 docx 的预期路径
    status: JobStatus = JobStatus.PENDING
    progress: int = 0  # 0-100
    message: str = ""  # 进度文本
    error: str | None = None
    created_at: datetime = field(default_factory=datetime.now)
    finished_at: datetime | None = None

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["status"] = self.status.value
        d["input_path"] = str(self.input_path)
        d["output_path"] = str(self.output_path)
        d["created_at"] = self.created_at.isoformat()
        d["finished_at"] = self.finished_at.isoformat() if self.finished_at else None
        return d


class JobStore:
    """单进程内存任务存储。线程安全（asyncio Lock）。"""

    def __init__(self, work_dir: Path | None = None) -> None:
        self._jobs: dict[str, Job] = {}
        self._lock = asyncio.Lock()
        self.work_dir = work_dir or Path(tempfile.gettempdir()) / "ai-proofreader-jobs"
        self.work_dir.mkdir(parents=True, exist_ok=True)

    async def create(self, *, filename: str, mode: str, source_bytes: bytes) -> Job:
        """落盘上传文件并登记任务。"""
        job_id = uuid.uuid4().hex[:12]
        job_dir = self.work_dir / job_id
        job_dir.mkdir(parents=True, exist_ok=True)

        # 强制 .docx 后缀，避免别有用心的扩展名
        safe_name = Path(filename).name
        if not safe_name.lower().endswith(".docx"):
            safe_name = f"{safe_name}.docx"
        input_path = job_dir / f"input_{safe_name}"
        output_path = job_dir / f"output_{safe_name}"

        input_path.write_bytes(source_bytes)

        job = Job(
            id=job_id,
            filename=safe_name,
            mode=mode,
            input_path=input_path,
            output_path=output_path,
        )
        async with self._lock:
            self._jobs[job_id] = job
        return job

    async def get(self, job_id: str) -> Job | None:
        async with self._lock:
            return self._jobs.get(job_id)

    async def update(self, job_id: str, **fields: Any) -> Job | None:
        async with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return None
            for key, value in fields.items():
                setattr(job, key, value)
            return job

    async def cleanup(self, job_id: str) -> None:
        """删除任务记录与临时文件。"""
        async with self._lock:
            job = self._jobs.pop(job_id, None)
        if job is not None:
            shutil.rmtree(job.input_path.parent, ignore_errors=True)


# ---------------------------------------------------------------------------
# 任务执行（包装 ProofReader 的同步 API）
# ---------------------------------------------------------------------------
async def run_proofread_job(store: JobStore, job: Job, *, api_key: str) -> None:
    """后台协程：调用 ProofReader 跑校对，更新 job 状态。

    校对本身是同步阻塞的，这里走 ``run_in_executor`` 避免堵事件循环。
    """
    await store.update(job.id, status=JobStatus.RUNNING, progress=10, message="加载文档…")

    loop = asyncio.get_running_loop()
    try:
        ok = await loop.run_in_executor(None, _run_blocking, job, api_key)
    except Exception as exc:
        logger.exception("Job %s failed", job.id)
        await store.update(
            job.id,
            status=JobStatus.FAILED,
            error=str(exc),
            finished_at=datetime.now(),
        )
        return

    if ok and job.output_path.exists():
        await store.update(
            job.id,
            status=JobStatus.SUCCEEDED,
            progress=100,
            message="校对完成",
            finished_at=datetime.now(),
        )
    else:
        await store.update(
            job.id,
            status=JobStatus.FAILED,
            error="校对失败：输出文件未生成（参见服务端日志）",
            finished_at=datetime.now(),
        )


def _run_blocking(job: Job, api_key: str) -> bool:
    """阻塞实际执行——给 executor 用，不要直接 await。"""
    # 延迟 import，避免 web 进程启动时立刻把 LLM 客户端拉起
    from proofreader_factory import create_proofreader

    proofreader = create_proofreader(api_key, job.mode)
    return proofreader.proofread_document(str(job.input_path), str(job.output_path))
