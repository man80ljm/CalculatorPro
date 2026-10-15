"""同一受保护请求内报告阶段并发送文件；不新增持久化下载或后台任务。"""
import asyncio
import base64
import json
from functools import partial

import anyio

from .errors import CalendarError
from .service import generate_for_course

CHUNK_BYTES = 64 * 1024
HEARTBEAT_SECONDS = 10


def _line(event):
    return (json.dumps(event, ensure_ascii=False, separators=(",", ":")) + "\n").encode("utf-8")


async def calendar_events(course_name, settings, teaching_assignment=None):
    loop = asyncio.get_running_loop()
    events = asyncio.Queue()

    def progress(percent, message):
        loop.call_soon_threadsafe(events.put_nowait,
            {"type": "progress", "percent": percent, "message": message})

    async def produce():
        try:
            filename, content, summary = await anyio.to_thread.run_sync(
                partial(generate_for_course, course_name, settings, progress=progress, teaching_assignment=teaching_assignment))
            await events.put({"type": "generated", "filename": filename,
                              "content": content, "summary": summary})
        except CalendarError as exc:
            await events.put({"type": "error", "message": str(exc),
                              "status": exc.status, **exc.extra})
        except Exception:
            await events.put({"type": "error", "message": "教学日历生成失败，请稍后重试。",
                              "status": 500, "code": "calendar_generation_failed"})

    worker = asyncio.create_task(produce())
    try:
        while True:
            try:
                event = await asyncio.wait_for(events.get(), timeout=HEARTBEAT_SECONDS)
            except asyncio.TimeoutError:
                yield _line({"type": "heartbeat"})
                continue
            if event["type"] == "generated":
                content = event["content"]
                yield _line({"type": "file", "filename": event["filename"],
                             "size": len(content), "summary": event["summary"]})
                for start in range(0, len(content), CHUNK_BYTES):
                    yield _line({"type": "chunk", "data": base64.b64encode(
                        content[start:start + CHUNK_BYTES]).decode("ascii")})
                yield _line({"type": "complete"})
                break
            yield _line(event)
            if event["type"] == "error":
                break
    finally:
        # 浏览器断开也等有上限的转换结束，让临时目录和请求并发槽正常释放。
        with anyio.CancelScope(shield=True):
            await worker
