# Slidoteka Export API (backend)

FastAPI-сервис: принимает `.pptx` из плагина PowerPoint, проверяет, считает SHA-256,
загружает в Google Drive (замена API Слайдотеки), пишет журнал в SQLite.

## Быстрый старт (локально)
```bash
python3.12 -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
cp .env.example .env          # DRIVE_STUB=true — Google не нужен
uvicorn app.main:app --reload --port 8000
pytest -q
```
Swagger: http://localhost:8000/docs

## API
| Метод | Путь | Описание |
|---|---|---|
| GET | `/health` | проверка живости |
| POST | `/api/exports` | multipart: `title`*, `description`, `tags` (через запятую), `file`* (.pptx) |
| GET | `/api/exports?limit=100` | журнал загрузок (заголовок `X-Admin-Token`, если задан `ADMIN_TOKEN`) |

Ответ `201`: `{"id","title","status","sha256","drive_file_id","drive_url"}`

| Код | Причина |
|---|---|
| 400 | пустой файл / не похож на PowerPoint |
| 409 | дубликат по SHA-256; в `detail.drive_url` ссылка на существующий файл |
| 413 | файл больше `MAX_UPLOAD_MB` |
| 415 | не `.pptx` или неверный MIME |
| 422 | не заполнено название |
| 502 | ошибка Google Drive |

## Google Drive
1. Google Cloud → проект → включить Drive API → Service Account → JSON-ключ.
2. Создать **Общий диск**, добавить email сервисного аккаунта с ролью «Диспетчер контента».
3. `GOOGLE_DRIVE_FOLDER_ID` — ID папки из URL; `DRIVE_STUB=false`.
4. На хостинге: весь JSON в переменную `GOOGLE_SERVICE_ACCOUNT_JSON`.

## Деплой (Docker)
```bash
docker build -t slidoteka-api . && docker run -p 8000:8000 --env-file .env slidoteka-api
```
На хостингах с эфемерным диском SQLite обнуляется при перезапуске — подключите постоянный том к `DB_PATH`.

## Деплой на VPS (Docker Compose + Caddy, HTTPS автоматически)
```bash
git clone <repo> /opt/slidoteka && cd /opt/slidoteka/backend
cp .env.example .env && nano .env     # DOMAIN, CORS_ORIGINS=https://<DOMAIN>, ADMIN_TOKEN, DRIVE_STUB...
docker compose up -d --build
curl https://<DOMAIN>/health
```
Файлы плагина (папка `dist` после `npm run build`) копируются в `addin-dist/`.
Обновление: `git pull && docker compose up -d --build`. Логи: `docker compose logs -f api`.
