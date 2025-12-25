"""
配置管理模块（Ollama 专用）
"""
import os
from typing import Optional
import dotenv

dotenv.load_dotenv()

class Settings:
    def __init__(self):
        # Ollama 配置
        self.ollama_base_url: str = os.getenv(
            "OLLAMA_BASE_URL", "http://localhost:11434"
        )
        self.default_model: str = os.getenv(
            "DEFAULT_MODEL", "deepseek-R1:latest"
        )
        self.default_temperature: float = float(
            os.getenv("DEFAULT_TEMPERATURE", "0.7")
        )

        # 服务配置
        self.host: str = os.getenv("HOST", "0.0.0.0")
        self.port: int = int(os.getenv("PORT", "8000"))
        self.debug: bool = os.getenv("DEBUG", "false").lower() == "true"

    def validate(self) -> bool:
        """Ollama 本地模式无需校验 Key"""
        return True

    def get_model_config(self, model_name: Optional[str] = None) -> dict:
        return {
            "model": model_name or self.default_model,
            "temperature": self.default_temperature,
            "base_url": self.ollama_base_url,
        }

settings = Settings()
