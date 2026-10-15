"""统一由网站的 API 错误处理器返回给页面的业务异常。"""
from web_app.service import ServiceError


class CalendarError(ServiceError):
    def __init__(self, message: str, status: int = 400, *, code="calendar_error", **extra):
        super().__init__(message, status=status, code=code, **extra)
