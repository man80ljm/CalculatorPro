"""两个 DeepSeek 账号的密钥池。

每个新的大纲读取或报告任务占用一个槽位，并选当前正在进行的调用更少的账号；
一样少时按 A、B 轮询。账号内按密钥顺序试，遇 429、401、403、网络超时或 5xx
就换下一把，每个账号最多一轮。两个账号都达到 DEEPSEEK_ACCOUNT_CONCURRENCY
（默认 5）时，新任务排队。报告比大纲先拿到空出来的槽位。

这里的并发才是打到 DeepSeek 的容量。web_app.limiter 里的次数只防刷。
日志和异常里不写密钥，也不写 Authorization。
"""
from __future__ import annotations

import os
import threading
import time

import requests

_GUARD = threading.Lock()
_LOCAL = threading.local()
_POOL: "Pool | None" = None

_INLINE_ENV = {"A": "DEEPSEEK_KEYS_A", "B": "DEEPSEEK_KEYS_B"}
_FILE_ENV = {"A": "DEEPSEEK_KEYS_A_FILE", "B": "DEEPSEEK_KEYS_B_FILE"}
_MAX_KEYS = 32
_MAX_KEY_LEN = 256


class PoolError(RuntimeError):
    pass


class QueueTimeout(PoolError):
    def __init__(self, position: int, eta_seconds: int):
        self.position = int(position)
        self.eta_seconds = int(eta_seconds)
        super().__init__(queue_label(self.position, self.eta_seconds))


class _Switch(Exception):
    """这把密钥不能用，换同一账号的下一把。"""


def queue_label(position: int, eta_seconds: int) -> str:
    return f"前面还有 {int(position)} 人，大约再等 {int(eta_seconds)} 秒"


def _env(name: str) -> str:
    return os.environ.get(name, "").strip()


def _env_int(name: str, default: int) -> int:
    raw = _env(name)
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError:
        return default
    if value < 0:
        return default
    return value


def _concurrency() -> int:
    return max(1, _env_int("DEEPSEEK_ACCOUNT_CONCURRENCY", 5))


def _estimate_seconds(kind: str) -> int:
    if kind == "syllabus":
        return _env_int("DEEPSEEK_QUEUE_SYLLABUS_SECONDS", 15)
    return _env_int("DEEPSEEK_QUEUE_REPORT_SECONDS", 30)


def _priority(kind: str) -> int:
    return 0 if kind == "report" else 1


def _read_key_file(path: str, label: str) -> list[str]:
    try:
        with open(path, encoding="utf-8-sig") as handle:
            lines = handle.readlines()
    except OSError:
        raise PoolError(f"{label} 无法读取") from None
    keys: list[str] = []
    for line in lines:
        text = line.strip()
        if not text or text.startswith("#"):
            continue
        if len(text) > _MAX_KEY_LEN or len(keys) >= _MAX_KEYS:
            raise PoolError(f"{label} 格式不正确")
        keys.append(text)
    return keys


def _parse_inline(raw: str, label: str) -> list[str]:
    if os.path.isfile(raw):
        return _read_key_file(raw, label)
    keys = []
    for part in raw.split(","):
        text = part.strip()
        if not text:
            continue
        if len(text) > _MAX_KEY_LEN or len(keys) >= _MAX_KEYS:
            raise PoolError(f"{label} 格式不正确")
        keys.append(text)
    return keys


def load_specs() -> list[tuple[str, list[str]]]:
    """返回 [(账号名, 密钥列表)]。只配了 DEEPSEEK_API_KEY 时是单账号。"""
    used_accounts = False
    specs: list[tuple[str, list[str]]] = []
    for name in ("A", "B"):
        file_path = _env(_FILE_ENV[name])
        inline = _env(_INLINE_ENV[name])
        if not file_path and not inline:
            continue
        used_accounts = True
        if file_path:
            keys = _read_key_file(file_path, _FILE_ENV[name])
        else:
            keys = _parse_inline(inline, _INLINE_ENV[name])
        if keys:
            specs.append((name, keys))
    if used_accounts:
        return specs
    single = _env("DEEPSEEK_API_KEY")
    if single:
        if len(single) > _MAX_KEY_LEN:
            raise PoolError("DEEPSEEK_API_KEY 格式不正确")
        return [("default", [single])]
    return []


def _redact(text: str, keys: list[str]) -> str:
    cleaned = "" if text is None else str(text)
    pieces: list[str] = []
    for key in keys:
        token = (key or "").strip()
        if len(token) < 4:
            continue
        pieces.append(token)
        pieces.append(f"Bearer {token}")
        if len(token) >= 8:
            pieces.extend(token[index : index + 8] for index in range(0, len(token) - 7))
    for piece in sorted(set(pieces), key=len, reverse=True):
        if piece and piece in cleaned:
            cleaned = cleaned.replace(piece, "***")
    return cleaned


def redact_text(text: str) -> str:
    try:
        keys = [key for _name, group in load_specs() for key in group]
    except PoolError:
        keys = []
    return _redact(text, keys)


class Account:
    def __init__(self, name: str, keys: list[str]):
        self.name = name
        self.keys = list(keys)
        self.inflight = 0

    def __repr__(self) -> str:
        return f"Account({self.name}, n={len(self.keys)}, inflight={self.inflight})"


class Lease:
    def __init__(self, pool: "Pool", account: Account):
        self.pool = pool
        self.account = account
        self.released = False

    @property
    def account_name(self) -> str:
        return self.account.name

    def bind(self) -> "Lease":
        bind_account(self.account)
        return self

    def unbind(self) -> None:
        if bound_account() is self.account:
            unbind_account()

    def release(self) -> None:
        self.pool.release(self)

    def __enter__(self) -> "Lease":
        return self.bind()

    def __exit__(self, exc_type, exc, tb) -> bool:
        self.unbind()
        self.release()
        return False


class _Waiter:
    def __init__(self, kind: str, on_queue):
        self.kind = kind
        self.priority = _priority(kind)
        self.seq = 0
        self.on_queue = on_queue
        self.event = threading.Event()
        self.lease: Lease | None = None
        self.error: BaseException | None = None
        self.position = 1
        self.eta = 0
        self.cancelled = False


class Ticket:
    def __init__(self, pool: "Pool", lease: Lease | None = None, waiter: _Waiter | None = None):
        self.pool = pool
        self.lease = lease
        self.waiter = waiter

    def wait(self, timeout: float | None = None) -> Lease:
        if self.lease is not None:
            return self.lease
        assert self.waiter is not None
        self.lease = self.pool.wait_waiter(self.waiter, timeout)
        return self.lease

    def cancel(self) -> None:
        if self.lease is not None or self.waiter is None:
            return
        self.pool.cancel_waiter(self.waiter)


class Pool:
    def __init__(self):
        self.lock = threading.Lock()
        self.accounts: list[Account] = []
        self.waiters: list[_Waiter] = []
        self._signature: tuple = ()
        self._rr = 0
        self._seq = 0

    def reload(self) -> None:
        specs = load_specs()
        signature = tuple((name, tuple(keys)) for name, keys in specs)
        if signature == self._signature:
            return
        busy = any(account.inflight for account in self.accounts) or bool(self.waiters)
        if not busy:
            self.accounts = [Account(name, keys) for name, keys in specs]
            self._signature = signature
            self._rr = 0
            return
        by_name = {account.name: account for account in self.accounts}
        rebuilt: list[Account] = []
        for name, keys in specs:
            current = by_name.get(name)
            if current is None:
                current = Account(name, keys)
            else:
                current.keys = list(keys)
            rebuilt.append(current)
        self.accounts = rebuilt
        self._signature = signature

    def _someone_ahead(self, kind: str) -> bool:
        priority = _priority(kind)
        return any(waiter.priority <= priority and not waiter.cancelled for waiter in self.waiters)

    def _try_reserve(self) -> Account | None:
        limit = _concurrency()
        available = [account for account in self.accounts if account.keys and account.inflight < limit]
        if not available:
            return None
        least = min(account.inflight for account in available)
        tied = [account for account in available if account.inflight == least]
        if len(tied) == 1:
            chosen = tied[0]
        else:
            chosen = tied[0]
            count = len(self.accounts)
            for offset in range(count):
                candidate = self.accounts[(self._rr + offset) % count]
                if candidate in tied:
                    chosen = candidate
                    self._rr = (self.accounts.index(candidate) + 1) % count
                    break
        chosen.inflight += 1
        return chosen

    def _grant(self) -> list[_Waiter]:
        granted: list[_Waiter] = []
        ordered = sorted(
            (waiter for waiter in self.waiters if not waiter.cancelled),
            key=lambda waiter: (waiter.priority, waiter.seq),
        )
        for waiter in ordered:
            account = self._try_reserve()
            if account is None:
                break
            self.waiters.remove(waiter)
            waiter.lease = Lease(self, account)
            granted.append(waiter)
        return granted

    def _mark_positions(self) -> list[tuple]:
        ordered = sorted(
            (waiter for waiter in self.waiters if not waiter.cancelled),
            key=lambda waiter: (waiter.priority, waiter.seq),
        )
        callbacks = []
        for index, waiter in enumerate(ordered, start=1):
            eta = index * _estimate_seconds(waiter.kind)
            waiter.position = index
            waiter.eta = eta
            if waiter.on_queue is not None:
                callbacks.append((waiter.on_queue, index, eta))
        return callbacks

    def _emit(self, callbacks: list[tuple]) -> None:
        for callback, position, eta in callbacks:
            callback(position, eta)

    def reserve(self, kind: str, on_queue=None) -> Ticket:
        """立刻占一个槽，或排进队列。回调里不能再进这把锁。"""
        with self.lock:
            self.reload()
            if not self.accounts:
                raise PoolError("未配置密钥")
            account = None if self._someone_ahead(kind) else self._try_reserve()
            granted: list[_Waiter] = []
            waiter = None
            if account is None:
                self._seq += 1
                waiter = _Waiter(kind, on_queue)
                waiter.seq = self._seq
                self.waiters.append(waiter)
                granted = self._grant()
            callbacks = self._mark_positions()
            self._emit(callbacks)
            for item in granted:
                item.event.set()
            if account is not None:
                return Ticket(self, lease=Lease(self, account))
            assert waiter is not None
            return Ticket(self, waiter=waiter)

    def wait_waiter(self, waiter: _Waiter, timeout: float | None) -> Lease:
        deadline = None if timeout is None else time.monotonic() + float(timeout)
        while True:
            if waiter.lease is not None:
                return waiter.lease
            if waiter.error is not None:
                raise waiter.error
            if deadline is None:
                waiter.event.wait()
                continue
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            waiter.event.wait(remaining)
        with self.lock:
            if waiter.lease is not None:
                return waiter.lease
            if waiter.error is not None:
                raise waiter.error
            position = waiter.position or 1
            eta = waiter.eta
            if waiter in self.waiters:
                self.waiters.remove(waiter)
            waiter.cancelled = True
            self._emit(self._mark_positions())
        raise QueueTimeout(position, eta)

    def cancel_waiter(self, waiter: _Waiter) -> None:
        with self.lock:
            if waiter.lease is not None or waiter.cancelled:
                return
            if waiter in self.waiters:
                self.waiters.remove(waiter)
            waiter.cancelled = True
            self._emit(self._mark_positions())

    def release(self, lease: Lease) -> None:
        if lease.released:
            return
        lease.released = True
        with self.lock:
            if lease.account.inflight > 0:
                lease.account.inflight -= 1
            granted = self._grant()
            callbacks = self._mark_positions()
            self._emit(callbacks)
            for waiter in granted:
                waiter.event.set()

    def abort(self) -> None:
        with self.lock:
            waiters = list(self.waiters)
            self.waiters.clear()
            for waiter in waiters:
                waiter.cancelled = True
                waiter.error = PoolError("排队已取消")
                waiter.event.set()
            for account in self.accounts:
                account.inflight = 0


def get_pool() -> Pool:
    global _POOL
    with _GUARD:
        if _POOL is None:
            _POOL = Pool()
        return _POOL


def reset_pool() -> None:
    global _POOL
    with _GUARD:
        old = _POOL
        _POOL = Pool()
    if old is not None:
        old.abort()
    unbind_account()


def bind_account(account: Account) -> None:
    _LOCAL.account = account


def unbind_account() -> None:
    _LOCAL.account = None


def bound_account() -> Account | None:
    return getattr(_LOCAL, "account", None)


def configured() -> bool:
    return bool(load_specs())


def primary_key() -> str:
    specs = load_specs()
    if not specs:
        return ""
    return specs[0][1][0]


def reserve(kind: str, on_queue=None) -> Ticket:
    return get_pool().reserve(kind, on_queue=on_queue)


def acquire(kind: str = "report", on_queue=None, wait_timeout: float | None = None) -> Lease:
    return reserve(kind, on_queue=on_queue).wait(wait_timeout)


def account_loads() -> dict[str, int]:
    pool = get_pool()
    with pool.lock:
        pool.reload()
        return {account.name: account.inflight for account in pool.accounts}


def queue_depth() -> int:
    """还在排队、尚未拿到槽位的任务数。不含已经在跑的。"""
    pool = get_pool()
    with pool.lock:
        return sum(1 for waiter in pool.waiters if not waiter.cancelled)


def _all_keys() -> list[str]:
    try:
        return [key for _name, group in load_specs() for key in group]
    except PoolError:
        return []


def _exchange(key: str, messages: list, max_tokens: int, timeout: int, temperature: float, kind: str) -> dict:
    """走 core_app.ai_report.requests.post，网页测试替换的就是这个函数。不记录请求头。"""
    from core_app import ai_report

    model = ai_report.deepseek_model()
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": int(max_tokens),
        "stream": False,
    }
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    started = time.perf_counter()

    def fail() -> None:
        ai_report.log_deepseek_call(kind, model, {}, time.perf_counter() - started, False)

    try:
        response = ai_report.requests.post(
            ai_report.deepseek_chat_url(),
            headers=headers,
            json=payload,
            timeout=timeout,
        )
    except requests.Timeout:
        fail()
        raise _Switch("网络超时") from None
    except requests.ConnectionError:
        fail()
        raise _Switch("网络错误") from None
    except requests.RequestException:
        fail()
        raise _Switch("网络错误") from None
    except Exception as exc:
        fail()
        raise PoolError(_redact(str(exc), _all_keys()) or "DeepSeek 调用失败") from None

    status = getattr(response, "status_code", None)
    try:
        status = 200 if status is None else int(status)
    except (TypeError, ValueError):
        status = 200
    if status == 429:
        fail()
        raise _Switch("限流")
    if status in (401, 403):
        fail()
        raise _Switch("鉴权失败")
    if status >= 500:
        fail()
        raise _Switch("服务暂时不可用")
    if status >= 400:
        fail()
        raise PoolError("DeepSeek 请求失败")
    try:
        body = response.json()
        content = body["choices"][0]["message"]["content"]
        if not isinstance(content, str):
            raise TypeError("content")
        usage = body.get("usage") or {}
        if not isinstance(usage, dict):
            usage = {}
    except PoolError:
        raise
    except Exception:
        fail()
        raise PoolError("DeepSeek 返回无法读取") from None
    elapsed = round(time.perf_counter() - started, 3)
    ai_report.log_deepseek_call(kind, model, usage, elapsed, True)
    return {
        "content": content,
        "model": model,
        "elapsed_seconds": elapsed,
        "prompt_tokens": usage.get("prompt_tokens"),
        "completion_tokens": usage.get("completion_tokens"),
        "total_tokens": usage.get("total_tokens"),
    }


def failover(account: Account, messages: list, max_tokens: int, timeout: int, temperature: float, kind: str) -> dict:
    keys = [key for key in account.keys if key]
    if not keys:
        raise PoolError("未配置密钥")
    last = "暂时不可用"
    for key in keys:
        try:
            return _exchange(key, messages, max_tokens, timeout, temperature, kind)
        except _Switch as exc:
            last = _redact(str(exc), _all_keys()) or last
            continue
    raise PoolError(f"DeepSeek 暂时不可用（{last}）")


def report_text(account: Account, messages: list, max_tokens: int, timeout: int = 120) -> str:
    result = failover(account, messages, max_tokens, timeout, 0.7, "report")
    return result["content"].strip()


def chat_completions(
    kind: str,
    messages: list,
    max_tokens: int = 8000,
    timeout: int = 120,
    temperature: float = 0.1,
) -> dict:
    ticket = reserve(kind)
    lease = None
    try:
        lease = ticket.wait(timeout)
        return failover(lease.account, messages, max_tokens, timeout, temperature, kind)
    finally:
        if lease is not None:
            lease.release()
        else:
            ticket.cancel()
