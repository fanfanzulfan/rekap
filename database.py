import sqlite3
import os

DB_PATH = "database/lembur.db"

def create_database():

    os.makedirs("database", exist_ok=True)

    conn = sqlite3.connect(DB_PATH)

    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS employees(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nik TEXT UNIQUE,
        nama TEXT,
        jabatan TEXT,
        satuan_kerja TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS lembur(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_id INTEGER,
        tanggal TEXT,
        jam_mulai TEXT,
        jam_selesai TEXT,
        durasi TEXT,
        jenis_pekerjaan TEXT,
        keterangan TEXT,
        requester TEXT,
        approver TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS evidence(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        lembur_id INTEGER,
        file_path TEXT
    )
    """)

    conn.commit()
    conn.close()