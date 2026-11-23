import sqlite3
import os
from datetime import datetime
from werkzeug.security import generate_password_hash

# On hosting, point DB_PATH at a persistent mounted volume (for example /data/hawaii_travel.db).
# Locally, keep the database beside this file.
DB_PATH = os.environ.get(
    "DB_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "hawaii_travel.db")
)
os.makedirs(os.path.dirname(os.path.abspath(DB_PATH)), exist_ok=True)

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn

def log_audit(conn, user_id, username, action, module, record_id, details):
    conn.execute(
        """INSERT INTO audit_trail (user_id, username, action, module, record_id, details, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (user_id, username, action, module, str(record_id), details, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    )

def ensure_column(cur, table, col_name, col_def):
    cols = [r[1] for r in cur.execute(f"PRAGMA table_info({table})").fetchall()]
    if col_name not in cols:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_def}")

def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL,
        branch TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        is_active INTEGER DEFAULT 1,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    ensure_column(cur, "users", "email", "TEXT")
    cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_email_unique ON users(email) WHERE email IS NOT NULL AND email != ''")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS password_reset_tokens (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        token_hash TEXT NOT NULL UNIQUE,
        expires_at TEXT NOT NULL,
        used_at TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS audit_trail (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        username TEXT,
        action TEXT,
        module TEXT,
        record_id TEXT,
        details TEXT,
        created_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        category TEXT NOT NULL,
        phone TEXT,
        email TEXT,
        branch TEXT DEFAULT 'DAR',
        credit_limit REAL DEFAULT 0,
        opening_debt_tzs REAL DEFAULT 0,
        opening_debt_usd REAL DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    ensure_column(cur, "customers", "opening_debt_tzs", "REAL DEFAULT 0")
    ensure_column(cur, "customers", "opening_debt_usd", "REAL DEFAULT 0")
    ensure_column(cur, "customers", "credit_balance_tzs", "REAL DEFAULT 0")
    ensure_column(cur, "customers", "credit_balance_usd", "REAL DEFAULT 0")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS tickets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ticket_date TEXT NOT NULL,
        ticket_number TEXT UNIQUE NOT NULL,
        passenger_name TEXT NOT NULL,
        airline TEXT NOT NULL,
        route TEXT,
        client_type TEXT NOT NULL,
        customer_id INTEGER,
        branch TEXT NOT NULL,
        currency TEXT NOT NULL,
        net_fare REAL NOT NULL DEFAULT 0,
        taxes REAL NOT NULL DEFAULT 0,
        service_charge REAL NOT NULL DEFAULT 0,
        commission REAL NOT NULL DEFAULT 0,
        total_selling REAL NOT NULL DEFAULT 0,
        payment_method TEXT NOT NULL,
        payment_status TEXT NOT NULL,
        amount_paid REAL NOT NULL DEFAULT 0,
        status TEXT NOT NULL DEFAULT 'ISSUED',
        bsp_reconciled INTEGER DEFAULT 0,
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE SET NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS customer_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        payment_date TEXT NOT NULL,
        customer_id INTEGER NOT NULL,
        ticket_id INTEGER,
        branch TEXT NOT NULL,
        currency TEXT NOT NULL,
        amount REAL NOT NULL,
        payment_method TEXT NOT NULL,
        reference TEXT,
        notes TEXT,
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE CASCADE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS customer_payment_allocations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        payment_id INTEGER NOT NULL,
        allocation_type TEXT NOT NULL,
        ticket_id INTEGER,
        amount REAL NOT NULL,
        FOREIGN KEY (payment_id) REFERENCES customer_payments(id) ON DELETE CASCADE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS bank_transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tx_date TEXT NOT NULL,
        currency TEXT NOT NULL,
        branch TEXT NOT NULL,
        tx_type TEXT NOT NULL,
        category TEXT NOT NULL,
        reference TEXT,
        description TEXT NOT NULL,
        amount_in REAL NOT NULL DEFAULT 0,
        amount_out REAL NOT NULL DEFAULT 0,
        is_reconciled INTEGER DEFAULT 0,
        statement_ref TEXT,
        source_module TEXT,
        source_id INTEGER,
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    ensure_column(cur, "bank_transactions", "source_module", "TEXT")
    ensure_column(cur, "bank_transactions", "source_id", "INTEGER")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS cash_transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tx_date TEXT NOT NULL,
        book_type TEXT NOT NULL,
        currency TEXT NOT NULL,
        tx_type TEXT NOT NULL,
        reference TEXT,
        description TEXT NOT NULL,
        amount_in REAL NOT NULL DEFAULT 0,
        amount_out REAL NOT NULL DEFAULT 0,
        source_module TEXT,
        source_id INTEGER,
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    ensure_column(cur, "cash_transactions", "source_module", "TEXT")
    ensure_column(cur, "cash_transactions", "source_id", "INTEGER")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS bank_advances (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        advance_date TEXT NOT NULL,
        branch TEXT NOT NULL,
        currency TEXT NOT NULL,
        amount REAL NOT NULL,
        taken_by TEXT NOT NULL,
        purpose TEXT NOT NULL,
        reference TEXT,
        bank_transaction_id INTEGER,
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS bank_advance_repayments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        advance_id INTEGER NOT NULL,
        repayment_date TEXT NOT NULL,
        amount REAL NOT NULL,
        received_into TEXT NOT NULL,
        reference TEXT,
        notes TEXT,
        bank_transaction_id INTEGER,
        cash_transaction_id INTEGER,
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (advance_id) REFERENCES bank_advances(id) ON DELETE CASCADE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS director_transfers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tx_date TEXT NOT NULL,
        direction TEXT NOT NULL,
        transfer_category TEXT NOT NULL,
        branch TEXT NOT NULL,
        account_source TEXT NOT NULL,
        currency TEXT NOT NULL,
        amount REAL NOT NULL,
        reference TEXT,
        description TEXT NOT NULL,
        status TEXT DEFAULT 'VERIFIED',
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS bsp_billing (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        billing_period TEXT NOT NULL,
        tx_date TEXT NOT NULL,
        ticket_number TEXT NOT NULL,
        airline TEXT,
        tx_type TEXT DEFAULT 'SALE',
        currency TEXT NOT NULL,
        bsp_amount REAL NOT NULL,
        match_status TEXT DEFAULT 'UNMATCHED',
        notes TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS bsp_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        payment_date TEXT NOT NULL,
        billing_period TEXT NOT NULL,
        currency TEXT NOT NULL,
        amount_paid REAL NOT NULL,
        bank_reference TEXT,
        notes TEXT,
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        expense_date TEXT NOT NULL,
        branch TEXT NOT NULL,
        category TEXT NOT NULL,
        description TEXT NOT NULL,
        currency TEXT NOT NULL,
        amount REAL NOT NULL,
        payment_source TEXT NOT NULL,
        status TEXT DEFAULT 'PENDING',
        approved_by TEXT,
        created_by TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""CREATE TABLE IF NOT EXISTS daily_closings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        closing_date TEXT NOT NULL, branch TEXT NOT NULL, currency TEXT NOT NULL,
        opening_cash REAL NOT NULL DEFAULT 0, expected_cash REAL NOT NULL DEFAULT 0,
        counted_cash REAL NOT NULL DEFAULT 0, variance REAL NOT NULL DEFAULT 0,
        notes TEXT, status TEXT NOT NULL DEFAULT 'SUBMITTED',
        submitted_by TEXT, reviewed_by TEXT, reviewed_at TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(closing_date, branch, currency)
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS app_settings (
        setting_key TEXT PRIMARY KEY,
        setting_value TEXT NOT NULL
    )""")
    cur.execute("INSERT OR IGNORE INTO app_settings (setting_key, setting_value) VALUES ('non_iata_fee_tzs', '10000')")
    cur.execute("INSERT OR IGNORE INTO app_settings (setting_key, setting_value) VALUES ('non_iata_fee_usd', '10')")

    cur.execute("""CREATE TABLE IF NOT EXISTS system_alerts (
        id INTEGER PRIMARY KEY AUTOINCREMENT, alert_type TEXT NOT NULL,
        message TEXT NOT NULL, severity TEXT DEFAULT 'WARNING',
        is_resolved INTEGER DEFAULT 0, created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        resolved_by TEXT, resolved_at TEXT
    )""")

    # Do not create public, predictable demo accounts automatically.
    # Create the first administrator explicitly with bootstrap_admin.py.
    cur.execute("SELECT COUNT(*) FROM customers")
    if cur.fetchone()[0] == 0:
        default_customers = [
            ("Walk-in Customer (General)", "WALKING", "N/A", "", "DAR", 0, 0, 0),
            ("Zanzibar Beach Resorts Ltd", "NON-IATA", "+255777000111", "accounts@zbr.co.tz", "UNGUJA", 15000000, 0, 0),
            ("Sub-Agent Kilimanjaro Tours", "IATA", "+255713000222", "ops@kilitours.co.tz", "DAR", 25000000, 0, 0),
        ]
        cur.executemany(
            "INSERT INTO customers (name, category, phone, email, branch, credit_limit, opening_debt_tzs, opening_debt_usd) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            default_customers
        )

    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully at:", DB_PATH)
