class AIHandler:
    """DeepSeek 调用占位实现"""

    def __init__(self, api_key: str):
        self.api_key = api_key

    def generate(self, prompt: str) -> str:
        raise NotImplementedError("AI 生成逻辑尚未实现")
