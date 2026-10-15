"""共用现有日历的转换并发锁；每次 Word 转 PDF 使用独立目录和配置。"""
import os
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory
from web_app.teaching_calendar.conversion import soffice_path, _slots, _stop
from .errors import LessonPlanError


def convert_pdf(source, directory):
    executable=soffice_path()
    if not _slots.acquire(timeout=10):
        raise LessonPlanError("页数检查服务正忙，请稍后重新检查。",503)
    try:
        with TemporaryDirectory(prefix="plan-pdf-",dir=directory) as temporary:
            root=Path(temporary); output=root/"output"; output.mkdir()
            command=[executable,f"-env:UserInstallation={(root/'profile').resolve().as_uri()}","--headless","--nologo","--nodefault","--norestore",
                     "--convert-to","pdf:writer_pdf_Export","--outdir",str(output),str(Path(source).resolve())]
            options={"creationflags":subprocess.CREATE_NO_WINDOW} if os.name=="nt" else {"start_new_session":True}
            process=subprocess.Popen(command,stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.PIPE,**options)
            try: process.communicate(timeout=180)
            except subprocess.TimeoutExpired:
                _stop(process); raise LessonPlanError("页数检查超时，可稍后重新检查。",504) from None
            generated=output/(Path(source).stem+".pdf")
            if process.returncode or not generated.is_file(): raise LessonPlanError("无法完成页数检查，请稍后重试。",502)
            destination=Path(directory)/(Path(source).stem+".pdf")
            destination.write_bytes(generated.read_bytes()); return destination
    finally: _slots.release()
