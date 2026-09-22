from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    DATABASE_URL: str = "postgresql+psycopg2://user:password@localhost:5432/auth_db"
    SECRET_KEY: str = "cambia_esta_clave_por_una_segura"
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 30

    @property
    def is_using_default_secret(self) -> bool:
        return self.SECRET_KEY == "cambia_esta_clave_por_una_segura"

    class Config:
        env_file = ".env"

settings = Settings()
