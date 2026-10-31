"""Проверка доступа к Google Drive без запуска сервера.

Запуск (из корня проекта):  python scripts/check_drive.py
Показывает, под каким аккаунтом работает сервер, и видна ли папка GOOGLE_DRIVE_FOLDER_ID.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import httpx
from google.auth.transport.requests import Request

from app.config import get_settings
from app.drive import DriveClient

s = get_settings()
client = DriveClient.from_settings(s)
client._creds.refresh(Request())
h = {"Authorization": f"Bearer {client._creds.token}"}

about = httpx.get("https://www.googleapis.com/drive/v3/about", params={"fields": "user"}, headers=h)
about.raise_for_status()
user = about.json()["user"]
print("Авторизация OK. Аккаунт:", user.get("emailAddress"))

if s.google_drive_folder_id:
    r = httpx.get(
        f"https://www.googleapis.com/drive/v3/files/{s.google_drive_folder_id}",
        params={"fields": "id,name,mimeType", "supportsAllDrives": "true"},
        headers=h,
    )
    if r.status_code == 200:
        print("Папка найдена:", r.json()["name"])
    else:
        print("ПАПКА НЕ НАЙДЕНА:", r.status_code, r.text[:200])
else:
    print("GOOGLE_DRIVE_FOLDER_ID не задан — файлы пойдут в корень «Моего диска».")
