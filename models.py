import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'data.db')


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            name_kana TEXT,
            representative TEXT,
            representative_title TEXT,
            representative_kana TEXT,
            corporate_number TEXT,
            capital_amount TEXT,
            corporation_type INTEGER DEFAULT 1,
            postal_code TEXT,
            prefecture TEXT,
            city TEXT,
            address TEXT,
            phone TEXT,
            fax TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            application_date TEXT,
            permit_type TEXT,
            governor_or_minister INTEGER DEFAULT 2,
            permit_category INTEGER,
            permit_number TEXT,
            permit_year TEXT,
            permit_month TEXT,
            permit_day TEXT,
            general_or_specific INTEGER DEFAULT 1,
            application_category INTEGER DEFAULT 1,
            validity_adjustment INTEGER DEFAULT 2,
            side_business INTEGER DEFAULT 2,
            side_business_type TEXT,
            permit_transfer_category INTEGER,
            old_permit_number TEXT,
            old_permit_year TEXT,
            old_permit_month TEXT,
            old_permit_day TEXT,
            city_code TEXT,
            business_types TEXT,
            existing_business_types TEXT,
            applicant_name TEXT,
            applicant_address TEXT,
            proxy_name TEXT,
            proxy_address TEXT,
            contact_organization TEXT,
            contact_name TEXT,
            contact_phone TEXT,
            contact_fax TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (customer_id) REFERENCES customers(id)
        );
        CREATE TABLE IF NOT EXISTS officers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            application_id INTEGER NOT NULL,
            last_name TEXT,
            first_name TEXT,
            last_name_kana TEXT,
            first_name_kana TEXT,
            role TEXT,
            full_or_part TEXT,
            sort_order INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (application_id) REFERENCES applications(id)
        );
    ''')
        # マイグレーション: 既存テーブルに新カラムを追加
    migrations = [
        "ALTER TABLE customers ADD COLUMN representative_title TEXT",
    ]
    for sql in migrations:
        try:
            conn.execute(sql)
        except sqlite3.OperationalError:
            pass  # カラムが既に存在する場合は無視
    conn.commit()
    conn.close()
