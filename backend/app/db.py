"""Слой SQLite. Схему БД ведёт Team Lead (Михаил); здесь рабочая минимальная версия."""
import sqlite3
from contextlib import contextmanager
from typing import Iterator

SCHEMA = """
CREATE TABLE IF NOT EXISTS exports (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    title         TEXT NOT NULL,
    description   TEXT,
    tags          TEXT,
    filename      TEXT,
    sha256        TEXT NOT NULL,
    size_bytes    INTEGER,
    status        TEXT NOT NULL CHECK (status IN ('pending','success','failed')),
    drive_file_id TEXT,
    drive_url     TEXT,
    error         TEXT,
    created_at    TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);
CREATE INDEX IF NOT EXISTS ix_exports_created ON exports(created_at DESC);
-- закрывает гонку: два одновременных одинаковых файла не станут оба 'success'
CREATE UNIQUE INDEX IF NOT EXISTS ux_exports_sha_success
    ON exports(sha256) WHERE status = 'success';
"""


@contextmanager
def connect(db_path: str) -> Iterator[sqlite3.Connection]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db(db_path: str) -> None:
    with connect(db_path) as c:
        c.executescript(SCHEMA)


def find_success_by_hash(db_path: str, sha256: str) -> sqlite3.Row | None:
    with connect(db_path) as c:
        return c.execute(
            "SELECT * FROM exports WHERE sha256=? AND status='success' LIMIT 1", (sha256,)
        ).fetchone()


def create_pending(db_path: str, *, title, description, tags, filename, sha256, size) -> int:
    with connect(db_path) as c:
        cur = c.execute(
            "INSERT INTO exports(title,description,tags,filename,sha256,size_bytes,status)"
            " VALUES (?,?,?,?,?,?, 'pending')",
            (title, description, tags, filename, sha256, size),
        )
        return cur.lastrowid


def mark_success(db_path: str, rec_id: int, drive_file_id: str, drive_url: str) -> None:
    with connect(db_path) as c:
        c.execute(
            "UPDATE exports SET status='success', drive_file_id=?, drive_url=? WHERE id=?",
            (drive_file_id, drive_url, rec_id),
        )


def mark_failed(db_path: str, rec_id: int, error: str) -> None:
    with connect(db_path) as c:
        c.execute("UPDATE exports SET status='failed', error=? WHERE id=?", (error[:1000], rec_id))


def list_exports(db_path: str, limit: int = 100) -> list[dict]:
    with connect(db_path) as c:
        rows = c.execute(
            "SELECT id,title,description,tags,filename,sha256,size_bytes,status,"
            "drive_url,error,created_at FROM exports ORDER BY id DESC LIMIT ?",
            (limit,),
        ).fetchall()
        return [dict(r) for r in rows]
