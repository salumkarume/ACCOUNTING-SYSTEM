"""Create the first admin account using environment variables; never hard-code credentials."""
import os
from database import get_db, init_db
from werkzeug.security import generate_password_hash

username = os.environ.get("ADMIN_USERNAME", "").strip()
full_name = os.environ.get("ADMIN_FULL_NAME", "").strip()
email = os.environ.get("ADMIN_EMAIL", "").strip().lower()
password = os.environ.get("ADMIN_PASSWORD", "")

if not username or not full_name or not email or len(password) < 9:
    raise SystemExit("Set ADMIN_USERNAME, ADMIN_FULL_NAME, ADMIN_EMAIL and ADMIN_PASSWORD (at least 9 characters).")

init_db()
conn = get_db()
try:
    if conn.execute("SELECT 1 FROM users LIMIT 1").fetchone():
        raise SystemExit("A user already exists. Do not run bootstrap again; use the in-app user management.")
    conn.execute(
        "INSERT INTO users (username, full_name, email, role, branch, password_hash) VALUES (?, ?, ?, 'admin', 'ALL', ?)",
        (username, full_name, email, generate_password_hash(password))
    )
    conn.commit()
    print(f"Administrator '{username}' created successfully.")
finally:
    conn.close()
