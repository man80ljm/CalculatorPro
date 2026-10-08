"""静态资源缓存：旧浏览器缓存不能让 HTML 和脚本混用。"""
import re

from tests.test_web_flow import app, client
from web_app.static_assets import page_response


def test_page_resources_change_version_when_scripts_or_html_change(tmp_path):
    page = tmp_path / "index.html"
    page.write_text('<link href="/static/style.css"><script src="/static/app.js"></script>', encoding="utf-8")
    script = tmp_path / "app.js"
    script.write_text("const old = true;", encoding="utf-8")
    (tmp_path / "style.css").write_text("body {}", encoding="utf-8")
    before = page_response(tmp_path, "index.html")
    assert before.headers["cache-control"] == "no-store"
    urls = re.findall(r'/static/[^" ]+', before.body.decode())
    assert len(urls) == 2 and all("?v=" in url for url in urls)
    assert urls[0].split("?v=")[1] == urls[1].split("?v=")[1]
    script.write_text("const updated = true;", encoding="utf-8")
    assert page_response(tmp_path, "index.html").body != before.body
    newer = page_response(tmp_path, "index.html").body
    page.write_text(page.read_text(encoding="utf-8") + '<div id="newUi"></div>', encoding="utf-8")
    assert page_response(tmp_path, "index.html").body != newer


def test_authenticated_html_and_static_assets_revalidate(client):
    page = client.get("/")
    assert page.headers["cache-control"] == "no-store"
    urls = re.findall(r'/static/[^" ]+\?v=[0-9a-f]+', page.text)
    assert len(urls) == 3
    script_url = next(url for url in urls if "/app.js?" in url)
    script = client.get(script_url)
    assert script.status_code == 200
    assert script.headers["cache-control"] == "no-cache"
    cached = client.get(script_url, headers={"if-none-match": script.headers["etag"]})
    assert cached.status_code == 304
    assert cached.headers["cache-control"] == "no-cache"
    assert "filesMore" not in script.text
    assert 'id="filesMore"' not in page.text
