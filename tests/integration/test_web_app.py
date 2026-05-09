"""Web UI 路由测试。

策略：用 FastAPI TestClient 发请求，校对部分 mock 成同步成功/失败，
不真打 LLM——LLM 链路有专门的 cassette 测试覆盖（test_ai_checker_recorded.py）。
"""

from __future__ import annotations

import asyncio
import shutil
from pathlib import Path

import pytest
from docx import Document
from fastapi.testclient import TestClient

from src.web.app import create_app
from src.web.jobs import JobStatus, JobStore

pytestmark = [pytest.mark.integration]


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------
@pytest.fixture
def docx_bytes(tmp_path: Path) -> bytes:
    """生成一个最小 .docx 的字节内容。"""
    p = tmp_path / "src.docx"
    doc = Document()
    doc.add_paragraph("操做系统是计算器科学的基础。")
    doc.save(p)
    return p.read_bytes()


@pytest.fixture
def store(tmp_path: Path) -> JobStore:
    """每个测试一个隔离的 JobStore。"""
    return JobStore(work_dir=tmp_path / "jobs")


@pytest.fixture
def client(store: JobStore, monkeypatch: pytest.MonkeyPatch) -> TestClient:
    """构造客户端 + 注入 fake api_key + 拦截后台校对协程，避免真打 LLM。"""
    monkeypatch.setenv("LICLOUD_TOKEN", "fake-token-for-tests")

    # 拦截后台真校对：用一个能成功生成空 docx 的 stub 替代
    async def stub_run_proofread_job(store, job, *, api_key: str):
        # 直接把输入 copy 为输出，模拟"无问题→直传"的快速路径
        shutil.copy2(job.input_path, job.output_path)
        await store.update(
            job.id,
            status=JobStatus.SUCCEEDED,
            progress=100,
            message="校对完成（stub）",
        )

    monkeypatch.setattr("src.web.app.run_proofread_job", stub_run_proofread_job)

    app = create_app(job_store=store)
    return TestClient(app)


# ---------------------------------------------------------------------------
# 路由级测试
# ---------------------------------------------------------------------------
class TestStaticRoutes:
    def test_healthz(self, client: TestClient):
        r = client.get("/healthz")
        assert r.status_code == 200
        assert r.json() == {"status": "ok"}

    def test_index_renders(self, client: TestClient):
        r = client.get("/")
        assert r.status_code == 200
        assert "AI 校对助手" in r.text
        # 4 种模式都列出来
        for value in ("comments", "revisions", "track_changes", "enhanced"):
            assert f'value="{value}"' in r.text
        # 12 个校对维度都渲染
        for cat in (
            "typo",
            "word_usage",
            "grammar",
            "logic",
            "punctuation",
            "numerals",
            "style",
            "terminology",
            "cn_en_terms",
            "code",
            "formula",
            "knowledge",
            "political",
        ):
            assert f'value="{cat}"' in r.text


class TestUploadValidation:
    def test_rejects_non_docx_extension(self, client: TestClient):
        r = client.post(
            "/jobs",
            files={"file": ("a.txt", b"hello", "text/plain")},
            data={"mode": "comments"},
        )
        assert r.status_code == 400
        assert "docx" in r.json()["detail"]

    def test_rejects_unknown_mode(self, client: TestClient, docx_bytes: bytes):
        r = client.post(
            "/jobs",
            files={
                "file": (
                    "x.docx",
                    docx_bytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            },
            data={"mode": "totally-fake"},
        )
        assert r.status_code == 400

    def test_rejects_empty_file(self, client: TestClient):
        r = client.post(
            "/jobs",
            files={"file": ("x.docx", b"", "application/octet-stream")},
            data={"mode": "comments"},
        )
        assert r.status_code == 400

    def test_503_when_no_api_key(
        self, store: JobStore, monkeypatch: pytest.MonkeyPatch, docx_bytes: bytes
    ):
        # 清空所有可能的 token 环境变量
        for env in (
            "LICLOUD_TOKEN",
            "LLM_API_TOKEN",
            "CLAUDE_SESSION_TOKEN",
            "ANTHROPIC_API_KEY",
            "OPENAI_API_KEY",
            "API_KEY",
        ):
            monkeypatch.delenv(env, raising=False)
        app = create_app(job_store=store)
        c = TestClient(app)
        r = c.post(
            "/jobs",
            files={"file": ("x.docx", docx_bytes, "application/octet-stream")},
            data={"mode": "comments"},
        )
        assert r.status_code == 503


class TestJobLifecycle:
    def test_full_pipeline(self, client: TestClient, docx_bytes: bytes, store: JobStore):
        # 1. 上传
        r = client.post(
            "/jobs",
            files={"file": ("input.docx", docx_bytes, "application/octet-stream")},
            data={"mode": "comments"},
        )
        assert r.status_code == 200
        # 返回的是 _job_status.html 片段，含进度容器
        assert "job-" in r.text
        assert "input.docx" in r.text

        # 提取 job id
        import re

        m = re.search(r'id="job-([a-f0-9]+)"', r.text)
        assert m, f"未在响应中找到 job id: {r.text[:300]}"
        job_id = m.group(1)

        # 2. 等后台 stub 跑完
        async def wait_done():
            for _ in range(50):
                job = await store.get(job_id)
                if job and job.status in (JobStatus.SUCCEEDED, JobStatus.FAILED):
                    return job
                await asyncio.sleep(0.05)
            return None

        job = asyncio.run(wait_done())
        assert job is not None
        assert job.status == JobStatus.SUCCEEDED, f"job 未完成: {job}"

        # 3. 状态接口返回成功片段（含下载链接）
        r2 = client.get(f"/jobs/{job_id}/status")
        assert r2.status_code == 200
        assert "下载校对结果" in r2.text
        assert f"/jobs/{job_id}/download" in r2.text

        # 4. 下载
        r3 = client.get(f"/jobs/{job_id}/download")
        assert r3.status_code == 200
        # docx 是 zip，PK 起头
        assert r3.content[:2] == b"PK"
        assert "wordprocessingml" in r3.headers["content-type"]

    def test_status_404_for_unknown_id(self, client: TestClient):
        r = client.get("/jobs/deadbeefdead/status")
        assert r.status_code == 404

    def test_categories_are_persisted_on_job(
        self, client: TestClient, docx_bytes: bytes, store: JobStore
    ):
        r = client.post(
            "/jobs",
            files={"file": ("x.docx", docx_bytes, "application/octet-stream")},
            data={
                "mode": "comments",
                "categories": ["typo", "punctuation", "code"],
            },
        )
        assert r.status_code == 200
        import re

        m = re.search(r'id="job-([a-f0-9]+)"', r.text)
        assert m, r.text[:300]
        job_id = m.group(1)

        async def fetch():
            return await store.get(job_id)

        job = asyncio.run(fetch())
        assert job is not None
        assert sorted(job.categories) == ["code", "punctuation", "typo"]

    def test_unknown_categories_fall_back_to_default(
        self, client: TestClient, docx_bytes: bytes, store: JobStore
    ):
        r = client.post(
            "/jobs",
            files={"file": ("x.docx", docx_bytes, "application/octet-stream")},
            data={"mode": "comments", "categories": ["totally-fake"]},
        )
        assert r.status_code == 200
        import re

        m = re.search(r'id="job-([a-f0-9]+)"', r.text)
        job_id = m.group(1)

        async def fetch():
            return await store.get(job_id)

        job = asyncio.run(fetch())
        assert job is not None
        # 未知 id 全部被过滤后，会落回 DEFAULT_CATEGORY_IDS
        assert "typo" in job.categories

    def test_download_404_when_not_succeeded(
        self, client: TestClient, store: JobStore, docx_bytes: bytes
    ):
        # 直接造一个 PENDING 的 job
        async def setup():
            return await store.create(filename="x.docx", mode="comments", source_bytes=docx_bytes)

        job = asyncio.run(setup())
        r = client.get(f"/jobs/{job.id}/download")
        assert r.status_code == 404
