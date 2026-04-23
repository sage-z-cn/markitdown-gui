import sqlite3
import os
import sys

if getattr(sys, 'frozen', False):
    _base_dir = os.path.dirname(sys.executable)
else:
    _base_dir = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(_base_dir, "sqlite.db")


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS conversions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT NOT NULL,
            output_file TEXT NOT NULL,
            status TEXT NOT NULL,
            error_message TEXT,
            created_at DATETIME DEFAULT (datetime('now', 'localtime'))
        )
    """)
    conn.commit()
    conn.close()


def add_record(source_file, output_file, status, error_message=None):
    conn = get_connection()
    cursor = conn.execute(
        "INSERT INTO conversions (source_file, output_file, status, error_message) VALUES (?, ?, ?, ?)",
        (source_file, output_file, status, error_message),
    )
    record_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return record_id


def get_all_records():
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM conversions ORDER BY created_at DESC"
    ).fetchall()
    conn.close()
    return [dict(row) for row in rows]


def delete_record(record_id):
    conn = get_connection()
    conn.execute("DELETE FROM conversions WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()


def update_output_file(record_id, new_output_file):
    conn = get_connection()
    conn.execute(
        "UPDATE conversions SET output_file = ? WHERE id = ?",
        (new_output_file, record_id),
    )
    conn.commit()
    conn.close()


def get_record_by_id(record_id):
    conn = get_connection()
    row = conn.execute(
        "SELECT * FROM conversions WHERE id = ?", (record_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None
