from __future__ import annotations

import os
import sqlite3
from pathlib import Path

from sqlalchemy import create_engine, insert, select

from app import extinguishers, metadata, service_logs


BASE_DIR = Path(__file__).resolve().parent
SQLITE_PATH = BASE_DIR / "database.db"


def normalize_database_url(raw_url: str) -> str:
    if raw_url.startswith("postgres://"):
        return raw_url.replace("postgres://", "postgresql+psycopg://", 1)
    if raw_url.startswith("postgresql://"):
        return raw_url.replace("postgresql://", "postgresql+psycopg://", 1)
    return raw_url


def main() -> None:
    target_url = os.environ.get("TARGET_DATABASE_URL") or os.environ.get("DATABASE_URL")
    if not target_url:
        raise SystemExit("TARGET_DATABASE_URL veya DATABASE_URL gerekli.")

    sqlite_conn = sqlite3.connect(SQLITE_PATH)
    sqlite_conn.row_factory = sqlite3.Row

    target_engine = create_engine(normalize_database_url(target_url), future=True)
    metadata.create_all(target_engine)

    extinguisher_rows = sqlite_conn.execute(
        "SELECT * FROM extinguishers ORDER BY id"
    ).fetchall()
    service_log_rows = sqlite_conn.execute(
        "SELECT * FROM service_logs ORDER BY id"
    ).fetchall()

    with target_engine.begin() as connection:
        existing_public_ids = {
            row[0]
            for row in connection.execute(select(extinguishers.c.public_id)).all()
        }

        for row in extinguisher_rows:
            data = dict(row)
            if data["public_id"] in existing_public_ids:
                continue
            connection.execute(insert(extinguishers).values(**data))

        existing_log_ids = {
            row[0] for row in connection.execute(select(service_logs.c.id)).all()
        }
        for row in service_log_rows:
            data = dict(row)
            if data["id"] in existing_log_ids:
                continue
            connection.execute(insert(service_logs).values(**data))

    sqlite_conn.close()
    print("Veriler PostgreSQL'e tasindi.")


if __name__ == "__main__":
    main()
