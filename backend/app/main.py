import hashlib
import re
import sqlite3
import zipfile
from contextlib import asynccontextmanager
from functools import lru_cache
from pathlib import Path
from tempfile import NamedTemporaryFile

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware

from . import db
from .config import Settings, get_settings
from .drive import Drive, DriveClient, DriveError, StubDrive

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
ALLOWED_MIME = {PPTX_MIME, "application/octet-stream"}  # браузеры/Office иногда шлют octet-stream
CHUNK = 1024 * 1024


@lru_cache
def get_drive() -> Drive:
    s = get_settings()
    return StubDrive() if s.drive_stub else DriveClient.from_settings(s)


@asynccontextmanager
async def lifespan(app: FastAPI):
    db.init_db(get_settings().db_path)
    yield


app = FastAPI(title="Slidoteka Export API", version="0.1.0", lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=get_settings().origins,  # конкретные origin, без "*"
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


def _safe_name(title: str) -> str:
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", title).strip()[:150] or "presentation"
    return f"{name}.pptx"


def _is_valid_pptx(path: Path) -> bool:
    if not zipfile.is_zipfile(path):
        return False
    try:
        with zipfile.ZipFile(path) as z:
            return "ppt/presentation.xml" in z.namelist()
    except zipfile.BadZipFile:
        return False


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/api/exports")
def list_exports(
    limit: int = 100,
    x_admin_token: str | None = Header(default=None),
    s: Settings = Depends(get_settings),
):
    """Журнал загрузок для страницы администратора."""
    if s.admin_token and x_admin_token != s.admin_token:
        raise HTTPException(401, "Нужен заголовок X-Admin-Token")
    return db.list_exports(s.db_path, max(1, min(limit, 500)))


@app.post("/api/exports", status_code=201)
async def create_export(
    request: Request,
    title: str = Form(..., min_length=1, max_length=255),
    description: str = Form("", max_length=2000),
    tags: str = Form("", max_length=500),
    file: UploadFile = File(...),
    s: Settings = Depends(get_settings),
    drive: Drive = Depends(get_drive),
):
    # 1. дешёвые проверки до чтения тела
    try:
        cl = int(request.headers.get("content-length", "0"))
    except ValueError:
        cl = 0
    if cl > s.max_bytes + 64 * 1024:
        raise HTTPException(413, f"Файл больше {s.max_upload_mb} МБ")

    filename = file.filename or ""
    if not filename.lower().endswith(".pptx"):
        raise HTTPException(415, "Допустимы только файлы .pptx")
    if (file.content_type or "") not in ALLOWED_MIME:
        raise HTTPException(415, f"Неподдерживаемый тип: {file.content_type}")

    tags_str = ",".join(list(dict.fromkeys(t.strip() for t in tags.split(",") if t.strip()))[:20])

    tmp_path: Path | None = None
    try:
        # 2. потоковая запись + SHA-256 + контроль размера
        sha = hashlib.sha256()
        size = 0
        with NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
            tmp_path = Path(tmp.name)
            while chunk := await file.read(CHUNK):
                size += len(chunk)
                if size > s.max_bytes:
                    raise HTTPException(413, f"Файл больше {s.max_upload_mb} МБ")
                sha.update(chunk)
                tmp.write(chunk)
        if size == 0:
            raise HTTPException(400, "Пустой файл")
        digest = sha.hexdigest()

        # 3. проверка содержимого
        if not _is_valid_pptx(tmp_path):
            raise HTTPException(400, "Файл повреждён или не является презентацией PowerPoint")

        # 4. дубликаты
        dup = db.find_success_by_hash(s.db_path, digest)
        if dup:
            raise HTTPException(
                409,
                detail={"message": "Такая презентация уже загружена", "drive_url": dup["drive_url"]},
            )

        # 5. журнал: pending
        rec_id = db.create_pending(
            s.db_path, title=title, description=description, tags=tags_str,
            filename=filename, sha256=digest, size=size,
        )

        # 6. загрузка в Drive
        drive_description = f"{description}\n\nТеги: {tags_str}".strip()
        try:
            info = await drive.upload(tmp_path, _safe_name(title), drive_description, PPTX_MIME)
        except DriveError as e:
            db.mark_failed(s.db_path, rec_id, str(e))
            raise HTTPException(502, "Не удалось сохранить файл в Google Drive") from e

        try:
            db.mark_success(s.db_path, rec_id, info["id"], info["webViewLink"])
        except sqlite3.IntegrityError:  # параллельная загрузка того же файла
            db.mark_failed(s.db_path, rec_id, "duplicate (race)")
            raise HTTPException(409, "Такая презентация уже загружена")

        return {
            "id": rec_id,
            "title": title,
            "status": "success",
            "sha256": digest,
            "drive_file_id": info["id"],
            "drive_url": info["webViewLink"],
        }
    finally:
        # 7. временный файл удаляется при любом исходе
        await file.close()
        if tmp_path:
            tmp_path.unlink(missing_ok=True)
