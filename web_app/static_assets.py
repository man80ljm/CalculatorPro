"""页面和脚本保持同一版本，避免缓存旧脚本访问新版页面中已删除的元素。"""
from __future__ import annotations

import hashlib
from pathlib import Path

from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles


def page_response(directory: Path, filename: str) -> HTMLResponse:
    html = (directory / filename).read_text(encoding="utf-8")
    assets = sorted(path for path in directory.iterdir() if path.suffix in {".js", ".css"})
    digest = hashlib.sha256(html.encode("utf-8"))
    for asset in assets:
        digest.update(asset.name.encode("utf-8"))
        digest.update(asset.read_bytes())
    version = digest.hexdigest()[:16]
    for asset in assets:
        url = f"/static/{asset.name}"
        html = html.replace(f'"{url}"', f'"{url}?v={version}"')
    return HTMLResponse(html, headers={"Cache-Control": "no-store"})


class RevalidatedStaticFiles(StaticFiles):
    async def get_response(self, path, scope):
        response = await super().get_response(path, scope)
        response.headers["Cache-Control"] = "no-cache"
        return response
