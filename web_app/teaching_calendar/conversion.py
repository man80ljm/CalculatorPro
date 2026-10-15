"""使用 LibreOffice 转换模板；每次转换隔离配置目录并限制进程资源。"""
import os
import shutil
import signal
import subprocess
import threading
from pathlib import Path
from tempfile import TemporaryDirectory

from .errors import CalendarError

_slots = threading.BoundedSemaphore(1)
FILTERS = {"xlsx": "xlsx:Calc MS Excel 2007 XML", "xls": "xls:MS Excel 97"}


def soffice_path():
    configured = os.environ.get("LIBREOFFICE_BIN", "").strip()
    candidates = [configured] if configured else [shutil.which("soffice")]
    if os.name == "nt" and not configured:
        candidates.extend(str(Path(os.environ.get(name, "C:/Program Files")) /
                              "LibreOffice/program/soffice.com") for name in ("ProgramFiles", "ProgramFiles(x86)"))
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(Path(candidate).resolve())
    raise CalendarError("服务器未安装或未配置 LibreOffice，暂时无法转换 .xls，请联系管理员。", 503,
                        code="converter_unavailable")


def _stop(process):
    if os.name == "nt":
        subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       creationflags=subprocess.CREATE_NO_WINDOW, timeout=10, check=False)
    else:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    if process.poll() is None:
        process.kill()
    process.communicate()


def convert_file(source: Path, output_format: str, directory: Path) -> Path:
    if output_format not in FILTERS:
        raise CalendarError("仅支持 .xls 和 .xlsx 格式。")
    executable = soffice_path()
    if not _slots.acquire(timeout=10):
        raise CalendarError("文件转换服务正忙，请稍后重试。", 503, code="converter_busy")
    try:
        with TemporaryDirectory(prefix="convert-", dir=directory) as working:
            root = Path(working)
            output = root / "output"
            output.mkdir()
            profile = root / "profile"
            # 不与桌面 Excel/LibreOffice 或其他请求共用用户配置及锁文件。
            command = [executable, f"-env:UserInstallation={profile.resolve().as_uri()}",
                       "--headless", "--nologo", "--nodefault", "--norestore",
                       "--convert-to", FILTERS[output_format], "--outdir", str(output), str(source.resolve())]
            options = {"creationflags": subprocess.CREATE_NO_WINDOW} if os.name == "nt" else {"start_new_session": True}
            try:
                process = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                                           stderr=subprocess.PIPE, **options)
                try:
                    process.communicate(timeout=60)
                except subprocess.TimeoutExpired:
                    _stop(process)
                    raise CalendarError("LibreOffice 文件转换超时，请稍后重试。", 504,
                                        code="conversion_timeout") from None
                generated = output / f"{source.stem}.{output_format}"
                if process.returncode != 0 or not generated.is_file() or generated.stat().st_size == 0:
                    raise CalendarError("LibreOffice 无法转换这份文件，请检查模板是否损坏或加密。", 422,
                                        code="conversion_failed")
                destination = directory / f"{source.stem}_converted.{output_format}"
                shutil.copyfile(generated, destination)
                return destination
            except CalendarError:
                raise
            except (OSError, subprocess.SubprocessError):
                raise CalendarError("LibreOffice 文件转换失败，请稍后重试或联系管理员。", 503,
                                    code="conversion_failed") from None
    finally:
        _slots.release()
