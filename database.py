from dotenv import load_dotenv
import sqlite3
import os

load_dotenv()
DB_PATH = os.getenv("DB_PATH", "workorderbot.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS work_orders (
            id TEXT PRIMARY KEY,
            property_code TEXT,
            property_name TEXT,
            unit_number TEXT,
            description TEXT,
            brief_desc TEXT,
            category TEXT,
            priority TEXT,
            status TEXT,
            created_date TEXT,
            scheduled_date TEXT,
            completed_date TEXT,
            employee TEXT,
            actual_start TEXT,
            actual_finish TEXT,
            actual_hours REAL,
            tech_notes TEXT,
            photos TEXT,
            source TEXT,
            days_open INTEGER,
            ai_priority INTEGER,
            ai_reason TEXT
        )
    """)
    conn.commit()
    conn.close()


if __name__ == "__main__":
    init_db()
    print("Database initialized.")
