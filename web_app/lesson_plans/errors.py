from web_app.service import ServiceError

class LessonPlanError(ServiceError):
    def __init__(self, message, status=400, *, code="lesson_plan_error", **extra):
        super().__init__(message, status=status, code=code, **extra)
