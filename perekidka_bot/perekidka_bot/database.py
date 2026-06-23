import aiosqlite
import asyncio
from datetime import datetime

DB_PATH = "perekidka.db"

async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id     INTEGER PRIMARY KEY,
                username    TEXT,
                role        TEXT,
                created_at  TEXT
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                from_apteka TEXT NOT NULL,
                to_apteka   TEXT NOT NULL,
                tovar       TEXT NOT NULL,
                miqdor      TEXT NOT NULL,
                status      TEXT DEFAULT 'yangi',
                created_by  INTEGER,
                shofer_id   INTEGER,
                created_at  TEXT,
                updated_at  TEXT
            )
        """)
        await db.commit()

async def get_user(user_id: int):
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM users WHERE user_id = ?", (user_id,)
        ) as cursor:
            return await cursor.fetchone()

async def save_user(user_id: int, username: str, role: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            INSERT INTO users (user_id, username, role, created_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET role=excluded.role
        """, (user_id, username, role, datetime.now().isoformat()))
        await db.commit()

async def create_order(from_apteka, to_apteka, tovar, miqdor, created_by):
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute("""
            INSERT INTO orders (from_apteka, to_apteka, tovar, miqdor, created_by, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (from_apteka, to_apteka, tovar, miqdor, created_by,
              datetime.now().isoformat(), datetime.now().isoformat()))
        await db.commit()
        return cursor.lastrowid

async def get_orders(status=None):
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        if status:
            async with db.execute(
                "SELECT * FROM orders WHERE status=? ORDER BY created_at DESC", (status,)
            ) as cursor:
                return await cursor.fetchall()
        else:
            async with db.execute(
                "SELECT * FROM orders ORDER BY created_at DESC LIMIT 50"
            ) as cursor:
                return await cursor.fetchall()

async def update_order_status(order_id: int, status: str, shofer_id=None):
    async with aiosqlite.connect(DB_PATH) as db:
        if shofer_id:
            await db.execute(
                "UPDATE orders SET status=?, shofer_id=?, updated_at=? WHERE id=?",
                (status, shofer_id, datetime.now().isoformat(), order_id)
            )
        else:
            await db.execute(
                "UPDATE orders SET status=?, updated_at=? WHERE id=?",
                (status, datetime.now().isoformat(), order_id)
            )
        await db.commit()
