import os
import asyncpg
from dotenv import load_dotenv

load_dotenv()

DB_CONFIG = {
    "database": os.environ.get("DB_NAME", "ecommerce"),
    "user": os.environ.get("DB_USER", "postgres"),
    "password": os.environ.get("DB_PASSWORD", ""),
    "host": os.environ.get("DB_HOST", "foodall.co.kr"),
    "port": int(os.environ.get("DB_PORT", "5432")),
}

async def query(sql: str, params: tuple = None) -> list[dict]:
    conn = await asyncpg.connect(**DB_CONFIG)
    try:
        rows = await conn.fetch(sql, *params) if params else await conn.fetch(sql)
        return [dict(row) for row in rows]
    finally:
        await conn.close()

async def execute(sql: str, params: tuple = None):
    conn = await asyncpg.connect(**DB_CONFIG)
    try:
        if params:
            await conn.execute(sql, *params)
        else:
            await conn.execute(sql)
    finally:
        await conn.close()
