import os
import pg8000.dbapi
from dotenv import load_dotenv

load_dotenv()

DB_CONFIG = {
    "host": os.environ.get("DB_HOST", "foodall.co.kr"),
    "database": os.environ.get("DB_NAME", "ecommerce"),
    "user": os.environ.get("DB_USER", "postgres"),
    "password": os.environ.get("DB_PASSWORD", ""),
    "port": int(os.environ.get("DB_PORT", "5432")),
    "ssl_context": False,  # SSL 비활성화
    "timeout": 10,         # 10초 타임아웃
}

def query(sql: str, params: tuple = None) -> list[dict]:
    conn = pg8000.dbapi.connect(**DB_CONFIG)
    try:
        cur = conn.cursor()
        cur.execute(sql, params or ())
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, row)) for row in cur.fetchall()]
    finally:
        conn.close()

def execute(sql: str, params: tuple = None):
    conn = pg8000.dbapi.connect(**DB_CONFIG)
    try:
        cur = conn.cursor()
        cur.execute(sql, params or ())
        conn.commit()
    finally:
        conn.close()
