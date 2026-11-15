from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    google_service_account_file: str | None = None
    google_service_account_json: str | None = None
    google_oauth_client_id: str | None = None
    google_oauth_client_secret: str | None = None
    google_oauth_refresh_token: str | None = None
    google_drive_folder_id: str | None = None
    google_make_public: bool = False
    drive_stub: bool = False

    max_upload_mb: int = 100
    db_path: str = "slidoteka.db"
    cors_origins: str = "https://localhost:3000"
    admin_token: str | None = None

    @property
    def max_bytes(self) -> int:
        return self.max_upload_mb * 1024 * 1024

    @property
    def origins(self) -> list[str]:
        return [o.strip() for o in self.cors_origins.split(",") if o.strip()]


@lru_cache
def get_settings() -> Settings:
    return Settings()
