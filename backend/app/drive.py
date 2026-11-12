import json
from pathlib import Path
from typing import AsyncIterator, Protocol

import anyio
import httpx
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials as OAuthCredentials

from .config import Settings

SCOPES = ["https://www.googleapis.com/auth/drive"]
TOKEN_URI = "https://oauth2.googleapis.com/token"
UPLOAD_URL = "https://www.googleapis.com/upload/drive/v3/files"
API_URL = "https://www.googleapis.com/drive/v3/files"


class DriveError(Exception):
    """Любая ошибка при общении с Google Drive."""


class Drive(Protocol):
    async def upload(self, path: Path, name: str, description: str, mime: str) -> dict: ...


class StubDrive:
    """Заглушка: ничего не отправляет. Нужна, чтобы фронтенд работал до настройки Google."""

    async def upload(self, path: Path, name: str, description: str, mime: str) -> dict:
        fake_id = f"stub-{path.stem}"
        return {"id": fake_id, "name": name, "webViewLink": f"https://example.com/stub/{fake_id}"}


async def _iter_file(path: Path, chunk: int = 1024 * 1024) -> AsyncIterator[bytes]:
    async with await anyio.open_file(path, "rb") as f:
        while data := await f.read(chunk):
            yield data


class DriveClient:
    def __init__(self, creds, folder_id: str | None, make_public: bool):
        self._creds = creds
        self._folder_id = folder_id
        self._make_public = make_public

    @classmethod
    def from_settings(cls, s: Settings) -> "DriveClient":
        # Вариант 1: OAuth от имени вашего Google-аккаунта (подходит для обычного Gmail)
        if s.google_oauth_refresh_token:
            if not (s.google_oauth_client_id and s.google_oauth_client_secret):
                raise DriveError("Для OAuth нужны GOOGLE_OAUTH_CLIENT_ID и GOOGLE_OAUTH_CLIENT_SECRET")
            creds = OAuthCredentials(
                token=None,
                refresh_token=s.google_oauth_refresh_token,
                client_id=s.google_oauth_client_id,
                client_secret=s.google_oauth_client_secret,
                token_uri=TOKEN_URI,
                scopes=SCOPES,
            )
        # Вариант 2: сервисный аккаунт (нужен Общий диск, т.е. Google Workspace)
        elif s.google_service_account_json:
            creds = service_account.Credentials.from_service_account_info(
                json.loads(s.google_service_account_json), scopes=SCOPES
            )
        elif s.google_service_account_file:
            creds = service_account.Credentials.from_service_account_file(
                s.google_service_account_file, scopes=SCOPES
            )
        else:
            raise DriveError("Не заданы данные доступа к Google (OAuth или сервисный аккаунт)")
        return cls(creds, s.google_drive_folder_id, s.google_make_public)

    def _token_sync(self) -> str:
        if not self._creds.valid:
            self._creds.refresh(Request())  # блокирующий вызов -> в поток
        return self._creds.token

    async def _headers(self) -> dict[str, str]:
        token = await anyio.to_thread.run_sync(self._token_sync)
        return {"Authorization": f"Bearer {token}"}

    async def upload(self, path: Path, name: str, description: str, mime: str) -> dict:
        size = path.stat().st_size
        metadata = {"name": name, "description": description}
        if self._folder_id:
            metadata["parents"] = [self._folder_id]
        timeout = httpx.Timeout(600.0, connect=10.0)
        try:
            auth = await self._headers()
            async with httpx.AsyncClient(timeout=timeout) as client:
                init = await client.post(
                    UPLOAD_URL,
                    params={
                        "uploadType": "resumable",
                        "supportsAllDrives": "true",
                        "fields": "id,name,webViewLink",
                    },
                    headers={
                        **auth,
                        "Content-Type": "application/json; charset=UTF-8",
                        "X-Upload-Content-Type": mime,
                        "X-Upload-Content-Length": str(size),
                    },
                    content=json.dumps(metadata),
                )
                init.raise_for_status()
                session_uri = init.headers["Location"]

                put = await client.put(
                    session_uri,
                    headers={"Content-Type": mime, "Content-Length": str(size)},
                    content=_iter_file(path),
                )
                put.raise_for_status()
                info = put.json()

                if self._make_public:
                    perm = await client.post(
                        f"{API_URL}/{info['id']}/permissions",
                        params={"supportsAllDrives": "true"},
                        headers=auth,
                        json={"role": "reader", "type": "anyone"},
                    )
                    perm.raise_for_status()
                return info
        except httpx.HTTPStatusError as e:
            raise DriveError(f"Drive HTTP {e.response.status_code}: {e.response.text[:300]}") from e
        except (httpx.HTTPError, KeyError) as e:
            raise DriveError(f"Drive network/protocol error: {e!r}") from e
        except Exception as e:  # ошибки google-auth (битый/просроченный ключ)
            raise DriveError(f"Drive auth error: {e!r}") from e
