from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    DATABASE_URL: str = "sqlite:///./shortener.db"
    BASE_URL: str = "http://localhost:8000"
    SHORT_CODE_LENGTH: int = 6
    API_KEY: str = "changeme"
    RATE_LIMIT: int = 100  # requests per minute

    model_config = {"env_file": ".env"}


settings = Settings()
