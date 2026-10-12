"""Independent local API configuration. No key appears in metadata or logs."""
import hashlib
import json
import os
import re
import threading
import time
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from cryptography.fernet import Fernet, InvalidToken
import base64,hmac
from sqlalchemy.orm import Session
from web_app.db import AdminApiConfiguration
from web_app.auth import get_secret_key
from web_app.deepseek_pool import load_specs, PoolError

BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-flash"

class ApiError(ValueError):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.status = status

class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None

class ApiSettings:
    def __init__(self, engine):
        self.engine = engine
        self.lock = threading.RLock()
        self.last_test = None

    def cipher(self):
        derived = hmac.new(get_secret_key().encode(), b"admin-api-encryption-v1", hashlib.sha256).digest()
        return Fernet(base64.urlsafe_b64encode(derived))

    def project(self):
        try: keys = [key for _, values in load_specs() for key in values]
        except PoolError: keys = []
        return {"DEEPSEEK_API_KEY": keys[0] if keys else "", "DEEPSEEK_MODEL": os.environ.get("DEEPSEEK_MODEL", "").strip() or DEFAULT_MODEL,
                "DEEPSEEK_BASE_URL": os.environ.get("DEEPSEEK_BASE_URL", "").strip() or BASE_URL}

    def load(self):
        with Session(self.engine) as db:
            row = db.get(AdminApiConfiguration, 1)
            if row is None:
                project = self.project()
                return {"source": "project" if project["DEEPSEEK_API_KEY"] else "custom", "model": project["DEEPSEEK_MODEL"], "key": "", "savedAt": None}
            try: return json.loads(self.cipher().decrypt(row.encrypted_config.encode()).decode())
            except (InvalidToken, ValueError, UnicodeError):
                raise ApiError("加密 API 配置无法读取，请检查服务器的会话密钥配置。", 503) from None

    def resolve(self, body, saved=None):
        saved = self.load() if saved is None else saved
        source = body.get("source", saved["source"])
        model = str(body.get("model", saved["model"])).strip()
        if not isinstance(source, str) or source not in {"custom", "project"} or not re.fullmatch(r"[a-zA-Z0-9_.:-]{1,100}", model):
            raise ApiError("请选择配置来源并填写有效模型名称。")
        if source == "project":
            project = self.project()
            if project.get("DEEPSEEK_BASE_URL", BASE_URL).rstrip("/") != BASE_URL:
                raise ApiError("服务器配置使用其他接口地址，课表识别只向 DeepSeek 官方接口发送密钥。")
            key = project.get("DEEPSEEK_API_KEY", "")
        else:
            incoming = body.get("key", "")
            if not isinstance(incoming, str):
                raise ApiError("API 密钥格式不正确。")
            key = incoming.strip() or (saved.get("key", "") if saved["source"] == "custom" else "")
        if not key:
            raise ApiError("尚未配置密钥，请填写 SK 或选择服务器配置。")
        if not key.startswith("sk-") or not 12 <= len(key) <= 512 or any(c.isspace() for c in key):
            raise ApiError("请填写完整的 sk- 密钥，不要包含空格或换行。")
        return {"source":source, "model":model, "key":key}

    @staticmethod
    def fingerprint(config):
        return hashlib.sha256((config["source"] + "\0" + config["model"] + "\0" + config["key"]).encode()).hexdigest()

    def metadata(self):
        with self.lock:
            saved = self.load()
            project = self.project()
            configured = bool(project.get("DEEPSEEK_API_KEY")) if saved["source"] == "project" else bool(saved.get("key"))
            if self.last_test is None:
                with Session(self.engine) as db:
                    row = db.get(AdminApiConfiguration, 1)
                    if row and row.last_test: self.last_test = json.loads(row.last_test)
            data = {"source":saved["source"], "model":saved["model"], "baseUrl":BASE_URL, "configured":configured, "projectAvailable":bool(project.get("DEEPSEEK_API_KEY")), "projectModel":project.get("DEEPSEEK_MODEL") or DEFAULT_MODEL, "savedAt":saved.get("savedAt"), "lastTest":None}
            if configured and self.last_test:
                try:
                    if self.last_test["fingerprint"] == self.fingerprint(self.resolve({}, saved)):
                        data["lastTest"] = {k:v for k,v in self.last_test.items() if k != "fingerprint"}
                except ApiError:
                    pass
            return data

    def save(self, body):
        with self.lock:
            config = self.resolve(body)
            config["savedAt"] = datetime.now(timezone.utc).isoformat()
            if config["source"] == "project": config["key"] = ""
            content = self.cipher().encrypt(json.dumps(config, ensure_ascii=False).encode()).decode()
            with Session(self.engine) as db:
                row = db.get(AdminApiConfiguration, 1)
                if row is None: db.add(AdminApiConfiguration(id=1, encrypted_config=content))
                else: row.encrypted_config = content
                db.commit()
            return self.metadata()

    def test(self, body, opener=None):
        with self.lock:
            config = self.resolve(body)
        request = urllib.request.Request(BASE_URL + "/models", headers={"Authorization":"Bearer " + config["key"], "Accept":"application/json"})
        from .provider import ProviderOpener
        client = opener or ProviderOpener(config)
        started = time.perf_counter()
        try:
            with client.open(request, timeout=15) as response:
                raw = response.read(1048577)
                if len(raw) > 1048576:
                    raise ApiError("模型列表响应过大，连接测试未通过。", 502)
                payload = json.loads(raw)
            models = [m["id"] for m in payload.get("data", []) if isinstance(m, dict) and isinstance(m.get("id"), str) and len(m["id"]) <= 100][:100]
            if not models:
                raise ApiError("接口没有返回有效模型列表，连接测试未通过。", 502)
            if config["model"] not in models:
                raise ApiError("密钥认证成功，但所选模型不在可用列表中，请核对模型名称。")
            result = {"ok":True, "message":"连接与密钥认证成功，所选模型可用。", "models":models, "latencyMs":round((time.perf_counter()-started)*1000), "testedAt":datetime.now(timezone.utc).isoformat(), "scope":"仅验证模型列表与密钥认证；未验证文本生成或余额。"}
        except urllib.error.HTTPError as exc:
            messages = {401:"密钥认证失败，请检查 SK。",403:"当前密钥无权访问接口。",429:"接口限流，请稍后再测试。"}
            raise ApiError(messages.get(exc.code, "API 返回 HTTP " + str(exc.code) + "，连接测试未通过。")) from None
        except ApiError:
            raise
        except Exception:
            raise ApiError("无法连接 DeepSeek 官方接口，请检查网络或稍后重试。", 502) from None
        with self.lock:
            self.last_test = dict(result, fingerprint=self.fingerprint(config))
            try:
                if self.fingerprint(self.resolve({})) == self.last_test["fingerprint"]:
                    with Session(self.engine) as db:
                        row = db.get(AdminApiConfiguration, 1)
                        if row:
                            row.last_test = json.dumps(self.last_test)
                            db.commit()
            except ApiError: pass
        return result
