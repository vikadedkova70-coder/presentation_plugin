import io
import os
import tempfile
import zipfile

_dir = tempfile.mkdtemp()
os.environ["DB_PATH"] = os.path.join(_dir, "test.db")
os.environ["MAX_UPLOAD_MB"] = "1"
os.environ["DRIVE_STUB"] = "true"
os.environ["ADMIN_TOKEN"] = ""

import pytest
from fastapi.testclient import TestClient

from app.drive import DriveError
from app.main import app, get_drive


class FakeDrive:
    def __init__(self):
        self.fail = False
        self.calls = 0

    async def upload(self, path, name, description, mime):
        self.calls += 1
        assert path.exists()  # файл существует на момент загрузки
        if self.fail:
            raise DriveError("boom")
        return {"id": f"id{self.calls}", "webViewLink": f"https://drive.example/{self.calls}"}


@pytest.fixture()
def fake_drive():
    d = FakeDrive()
    app.dependency_overrides[get_drive] = lambda: d
    yield d
    app.dependency_overrides.clear()


@pytest.fixture()
def client(fake_drive):
    with TestClient(app) as c:
        yield c


def make_pptx(marker: str = "a", extra_bytes: int = 0) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
        z.writestr("ppt/presentation.xml", f"<p>{marker}</p>")
        if extra_bytes:
            z.writestr("ppt/media/big.bin", os.urandom(extra_bytes))
    return buf.getvalue()
