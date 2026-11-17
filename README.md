# Запуск проекта «Слайдотека» на локальном ноутбуке

Репозиторий: https://github.com/vikadedkova70-coder/presentation_plugin

> **Статус проекта (по коду репозитория):**
> - Backend — есть в репозитории (`backend/`), запускается локально по инструкции ниже.
> - Плагин PowerPoint — **в репозитории исходников нет** (папка `backend/addin-dist/` пустая,
>   только `.gitkeep`). Плагин существует и работает **только локально** у автора,
>   в продакшен не задеплоен. Поэтому инструкция по запуску плагина ниже — общая
>   для Office Add-in, точные команды зависят от исходников, которых в репо пока нет.

---

## 1. Что есть в репозитории (проверено)

```
backend/
  app/main.py            # FastAPI: POST /api/exports, GET /api/exports, GET /health
  app/config.py, db.py, drive.py
  scripts/get_refresh_token.py   # OAuth для Google Drive
  scripts/check_drive.py         # проверка связи с Drive
  tests/test_exports.py
  .env.example, Dockerfile, docker-compose.yml, Caddyfile
  addin-dist/            # ПУСТАЯ — сюда кладутся собранные файлы плагина
docs/api-contract.md     # контракт API
.github/workflows/ci.yml
```

## 2. Запуск backend'а на ноутбуке (проверено по backend/README.md)

Требования: **Python 3.12**.

```bash
git clone https://github.com/vikadedkova70-coder/presentation_plugin.git
cd presentation_plugin/backend

# виртуальное окружение
python3.12 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

# зависимости
pip install -r requirements-dev.txt

# конфигурация — по умолчанию DRIVE_STUB=true, Google НЕ нужен
cp .env.example .env

# запуск
uvicorn app.main:app --reload --port 8000
```

Проверка:

```bash
curl http://localhost:8000/health
# {"status":"ok"}
```

Swagger UI: http://localhost:8000/docs

Тесты:

```bash
pytest -q
```

### Что при этом работает

- `POST /api/exports` — принимает `.pptx` (multipart: `title`, `description`, `tags`, `file`),
  проверяет, что файл — валидный PowerPoint (ZIP с `ppt/presentation.xml`),
  считает SHA-256, отсекает дубликаты (код 409), пишет в SQLite.
- В режиме `DRIVE_STUB=true` Google Drive не вызывается, в ответе —
  фейковая ссылка `https://example.com/stub/...`.
- Лимит размера: `MAX_UPLOAD_MB=100` (по умолчанию).

### Подключение настоящего Google Drive (необязательно)

**Вариант A — OAuth (обычный Gmail):**

```bash
python scripts/get_refresh_token.py
```

Скрипт выдаёт `GOOGLE_OAUTH_CLIENT_ID`, `GOOGLE_OAUTH_CLIENT_SECRET`,
`GOOGLE_OAUTH_REFRESH_TOKEN` — впишите их в `.env`, поставьте `DRIVE_STUB=false`.

**Вариант B — сервисный аккаунт (нужен Google Workspace):**

1. Google Cloud → проект → включить Drive API → Service Account → JSON-ключ.
2. Создать **Общий диск**, добавить email сервисного аккаунта с ролью «Диспетчер контента».
3. В `.env`:
   ```
   GOOGLE_SERVICE_ACCOUNT_FILE=./secrets/sa.json
   GOOGLE_DRIVE_FOLDER_ID=<ID папки из URL>
   DRIVE_STUB=false
   ```

Проверка связи: `python scripts/check_drive.py`.

## 3. Плагин PowerPoint — текущий статус

**Исходников плагина в репозитории нет.** Известно только:

- Плагин собирается в папку `dist` (`npm run build`), а готовые файлы
  копируются в `backend/addin-dist/` — оттуда их раздаёт сервер
  (упомянуто в backend/README.md в разделе деплоя).
- Плагин работает **только локально** (подключён к `http://localhost:8000`),
  на сервер не задеплоен — поэтому в `docs/api-contract.md` адрес сервера
  указан как «локально: `http://localhost:8000`», а «на сервере: будет добавлен после деплоя».

### Как запустить плагин локально (общий порядок для Office Add-in)

Точных команд из репозитория нет, стандартный цикл такой:

```bash
# в папке с исходниками плагина (у автора, не в репо)
npm install
npm run build
# скопировать dist/* в backend/addin-dist/
```

Установка в PowerPoint (Windows):

1. Положить XML-манифест плагина в сетевую папку.
2. PowerPoint → **Файл → Параметры → Центр управления безопасностью →
   Параметры центра управления безопасностью → Доверенные каталоги надстроек** → добавить папку.
3. Перезапустить PowerPoint → **Вставка → Надстройки → Мои надстройки**.

Манифест должен указывать на локальный адрес сервера плагина
(например `https://localhost:3000/taskpane.html`), а в коде плагина —
адрес backend'а `http://localhost:8000`.

## 4. API, которое использует плагин (проверено по docs/api-contract.md и коду)

| Метод | Путь | Описание |
|---|---|---|
| GET | `/health` | `{"status":"ok"}` |
| POST | `/api/exports` | multipart: `title`* (1–255), `description` (до 2000), `tags` (через запятую, ≤20), `file`* (.pptx, ≤100 МБ) |
| GET | `/api/exports?limit=100` | журнал; заголовок `X-Admin-Token`, если задан `ADMIN_TOKEN` |

Успех — 201: `{"id","title","status","sha256","drive_file_id","drive_url"}`

| Код | Причина |
|---|---|
| 400 | пустой файл / не похож на PowerPoint |
| 409 | дубликат по SHA-256; ссылка в `detail.drive_url` |
| 413 | больше `MAX_UPLOAD_MB` |
| 415 | не `.pptx` или неверный MIME |
| 422 | не заполнено название |
| 502 | ошибка Google Drive |

## 5. Деплой на сервер (когда понадобится)

```bash
git clone <repo> /opt/slidoteka && cd /opt/slidoteka/backend
cp .env.example .env && nano .env   # DOMAIN, CORS_ORIGINS=https://<DOMAIN>, ADMIN_TOKEN, DRIVE_STUB...
docker compose up -d --build
curl https://<DOMAIN>/health
```

Обновление: `git pull && docker compose up -d --build`.
Логи: `docker compose logs -f api`.

> На хостингах с эфемерным диском SQLite обнуляется — подключите постоянный том к `DB_PATH`.
