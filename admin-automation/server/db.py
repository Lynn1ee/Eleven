"""SQLite 数据库连接、建表、迁移"""
import os
import sqlite3
import threading

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
DB_PATH = os.path.join(DATA_DIR, "app.db")

_local = threading.local()
_init_lock = threading.Lock()


def get_db():
    conn = getattr(_local, 'connection', None)
    if conn is None:
        os.makedirs(DATA_DIR, exist_ok=True)
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA busy_timeout=5000")
        with _init_lock:
            _init_schema(conn)
        _local.connection = conn
    return conn


def _init_schema(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            last_login_at TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id),
            token TEXT UNIQUE NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            expires_at TIMESTAMP NOT NULL
        );

        CREATE TABLE IF NOT EXISTS user_smtp (
            user_id INTEGER PRIMARY KEY REFERENCES users(id),
            smtp_host TEXT NOT NULL DEFAULT 'smtp.qq.com',
            smtp_port INTEGER NOT NULL DEFAULT 465,
            smtp_user TEXT NOT NULL,
            smtp_pass_encrypted TEXT NOT NULL,
            from_name TEXT DEFAULT '元气工作站',
            to_addr TEXT DEFAULT '',
            cc_addr TEXT DEFAULT '',
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS app_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            data_key TEXT NOT NULL,
            data_value TEXT NOT NULL,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, data_key)
        );

        CREATE TABLE IF NOT EXISTS email_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id),
            subject TEXT NOT NULL,
            to_addr TEXT NOT NULL,
            cc_addr TEXT DEFAULT '',
            from_addr TEXT DEFAULT '',
            status TEXT NOT NULL,
            error_message TEXT,
            sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS email_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id),
            name TEXT NOT NULL,
            data_source TEXT NOT NULL DEFAULT '',
            to_addrs TEXT DEFAULT '',
            cc_addrs TEXT DEFAULT '',
            send_type TEXT NOT NULL DEFAULT 'manual',
            subject TEXT DEFAULT '',
            body TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS invoice_cycles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_name TEXT NOT NULL,
            year INTEGER NOT NULL,
            month_start INTEGER NOT NULL,
            month_end INTEGER NOT NULL,
            summary_month INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS performance_scores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            score REAL NOT NULL,
            is_night_shift INTEGER NOT NULL DEFAULT 0,
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(name, year, month)
        );

        CREATE TABLE IF NOT EXISTS score_status_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            note TEXT NOT NULL DEFAULT '',
            UNIQUE(name, year, month)
        );

        CREATE TABLE IF NOT EXISTS ranking_exclusions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            UNIQUE(name, year, month)
        );

        CREATE TABLE IF NOT EXISTS staff_ranks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            rank_level TEXT NOT NULL DEFAULT '客服专员',
            is_excluded INTEGER NOT NULL DEFAULT 0,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS invoice_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cycle_id INTEGER NOT NULL REFERENCES invoice_cycles(id),
            name TEXT NOT NULL,
            bank_name TEXT DEFAULT '',
            month1_amount REAL DEFAULT 0,
            month2_amount REAL DEFAULT 0,
            total_amount REAL DEFAULT 0,
            pdf1_filename TEXT DEFAULT '',
            pdf2_filename TEXT DEFAULT '',
            invoice_no1 TEXT DEFAULT '',
            invoice_no2 TEXT DEFAULT '',
            invoice_date1 TEXT DEFAULT '',
            invoice_date2 TEXT DEFAULT '',
            verification1_filename TEXT DEFAULT '',
            verification2_filename TEXT DEFAULT '',
            bill_image_filename TEXT DEFAULT '',
            invoice_list_filename TEXT DEFAULT '',
            status TEXT NOT NULL DEFAULT 'draft',
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    # 迁移：添加发票日期字段
    for col in ("invoice_date1", "invoice_date2"):
        try:
            conn.execute(f"ALTER TABLE invoice_records ADD COLUMN {col} TEXT DEFAULT ''")
        except sqlite3.OperationalError:
            pass

    # 迁移：ranking_exclusions 加 reason 字段
    try:
        conn.execute("ALTER TABLE ranking_exclusions ADD COLUMN reason TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass

    # 迁移：performance_scores 加 historical_tier 字段（绿色单元格 = 高级客服）
    try:
        conn.execute("ALTER TABLE performance_scores ADD COLUMN historical_tier TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass

    conn.execute("INSERT OR IGNORE INTO users (id, email) VALUES (0, '__shared__')")
    conn.commit()
