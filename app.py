import os
import csv
import io
import sqlite3
import shutil
import secrets
import hashlib
import smtplib
from email.message import EmailMessage
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path
from flask import send_file
from datetime import date, datetime, timedelta
from functools import wraps
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, Response, jsonify
)
from werkzeug.security import generate_password_hash, check_password_hash
from database import get_db, init_db, log_audit

app = Flask(__name__)
# A stable SECRET_KEY must be configured in the hosting dashboard.
# The random local fallback avoids shipping a known production secret.
configured_secret = os.environ.get("SECRET_KEY")
if os.environ.get("APP_ENV") == "production" and not configured_secret:
    raise RuntimeError("SECRET_KEY must be configured before running in production.")
app.secret_key = configured_secret or secrets.token_hex(32)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "1" if os.environ.get("APP_ENV") == "production" else "0") == "1",
    MAX_CONTENT_LENGTH=10 * 1024 * 1024,  # 10 MB upload limit
)

@app.after_request
def add_security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    if request.is_secure or app.config["SESSION_COOKIE_SECURE"]:
        response.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
    return response

init_db()

def valid_password(password):
    return len(password) >= 9

def send_password_reset_email(recipient, reset_url):
    """Send a reset link only when SMTP credentials are configured by the host."""
    host = os.environ.get("SMTP_HOST")
    port = int(os.environ.get("SMTP_PORT", "587"))
    username = os.environ.get("SMTP_USERNAME")
    password = os.environ.get("SMTP_PASSWORD")
    sender = os.environ.get("SMTP_FROM", username or "")
    if not all([host, username, password, sender]):
        return False
    msg = EmailMessage()
    msg["Subject"] = "HAWAII ERP - Password reset"
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content(
        "Tumepokea ombi la kubadili password ya HAWAII ERP.\n\n"
        f"Fungua link hii ndani ya dakika 30: {reset_url}\n\n"
        "Kama hukuomba mabadiliko haya, puuza email hii."
    )
    with smtplib.SMTP(host, port, timeout=20) as server:
        server.starttls()
        server.login(username, password)
        server.send_message(msg)
    return True

def parse_num(val):
    if val is None:
        return 0.0
    s = str(val).replace(",", "").strip()
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0

def sync_ledger_entry(conn, source_module, source_id, tx_date, currency, branch, method_or_source, direction, category, ref, desc, amount, username):
    """Safely clears previous linked bank/cash rows for (source_module, source_id) and inserts fresh synchronized ledger rows."""
    conn.execute("DELETE FROM bank_transactions WHERE source_module=? AND source_id=?", (source_module, source_id))
    conn.execute("DELETE FROM cash_transactions WHERE source_module=? AND source_id=?", (source_module, source_id))

    if amount <= 0 or not method_or_source or method_or_source == "CREDIT":
        return

    src = method_or_source.upper()
    if src in ("BANK", "BANK_TZS", "BANK_USD"):
        b_curr = "TZS" if src == "BANK_TZS" else ("USD" if src == "BANK_USD" else currency)
        if direction == "IN":
            conn.execute(
                """INSERT INTO bank_transactions
                   (tx_date, currency, branch, tx_type, category, reference, description, amount_in, amount_out, source_module, source_id, created_by)
                   VALUES (?, ?, ?, 'DEPOSIT', ?, ?, ?, ?, 0, ?, ?, ?)""",
                (tx_date, b_curr, branch, category, ref, desc, amount, source_module, source_id, username)
            )
        elif direction == "OUT":
            tx_tp = "BSP_PAYMENT" if category == "BSP" else "WITHDRAWAL"
            conn.execute(
                """INSERT INTO bank_transactions
                   (tx_date, currency, branch, tx_type, category, reference, description, amount_in, amount_out, source_module, source_id, created_by)
                   VALUES (?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?)""",
                (tx_date, b_curr, branch, tx_tp, category, ref, desc, amount, source_module, source_id, username)
            )
    elif src in ("CASH", "DAR_CASH", "UNGUJA_CASH", "LIPA_NAMBA"):
        if src == "CASH":
            book = "DAR_CASH" if branch == "DAR" else "UNGUJA_CASH"
        else:
            book = src
        if direction == "IN":
            conn.execute(
                """INSERT INTO cash_transactions
                   (tx_date, book_type, currency, tx_type, reference, description, amount_in, amount_out, source_module, source_id, created_by)
                   VALUES (?, ?, ?, 'RECEIPT', ?, ?, ?, 0, ?, ?, ?)""",
                (tx_date, book, currency, ref, desc, amount, source_module, source_id, username)
            )
        elif direction == "OUT":
            conn.execute(
                """INSERT INTO cash_transactions
                   (tx_date, book_type, currency, tx_type, reference, description, amount_in, amount_out, source_module, source_id, created_by)
                   VALUES (?, ?, ?, 'PAYMENT', ?, ?, 0, ?, ?, ?, ?)""",
                (tx_date, book, currency, ref, desc, amount, source_module, source_id, username)
            )

def recompute_customer_allocations(conn, customer_id):
    """Recalculates ticket payment allocations & opening debt settlements cleanly from all customer_payments."""
    cust = conn.execute("SELECT * FROM customers WHERE id=?", (customer_id,)).fetchone()
    if not cust:
        return

    # Reset all credit/unpaid-origin tickets for this customer back to their initial ticket-time paid amount
    tickets = conn.execute("SELECT * FROM tickets WHERE customer_id=? ORDER BY ticket_date ASC, id ASC", (customer_id,)).fetchall()
    ticket_paid_map = {}
    for t in tickets:
        # If ticket was issued with immediate payment (BANK/CASH/LIPA_NAMBA), keep its direct payment; if CREDIT, base is 0
        # Wait: to be 100% accurate, if a ticket had initial payment recorded on the ticket itself vs via customer_payments:
        # Let's check: any payment recorded via customer_payments adds on top of initial ticket payment.
        # Even simpler: when ticket is created/edited, store direct paid; here we compute total paid per ticket from direct + customer_payments!
        pass

def get_customer_debt_summary(conn, customer_id):
    cust = conn.execute("SELECT * FROM customers WHERE id=?", (customer_id,)).fetchone()
    if not cust:
        return {"tzs": 0.0, "usd": 0.0}

    # Total unpaid on tickets (tickets already have amount_paid updated when payments are applied)
    t_debt_tzs = conn.execute(
        "SELECT COALESCE(SUM(total_selling - amount_paid), 0) FROM tickets WHERE customer_id=? AND currency='TZS' AND status='ISSUED'",
        (customer_id,)
    ).fetchone()[0]
    t_debt_usd = conn.execute(
        "SELECT COALESCE(SUM(total_selling - amount_paid), 0) FROM tickets WHERE customer_id=? AND currency='USD' AND status='ISSUED'",
        (customer_id,)
    ).fetchone()[0]

    # Payments that were not linked to a ticket and exceeded ticket debts reduce opening_debt
    unalloc_pay_tzs = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM customer_payments WHERE customer_id=? AND currency='TZS' AND ticket_id IS NULL",
        (customer_id,)
    ).fetchone()[0]
    unalloc_pay_usd = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM customer_payments WHERE customer_id=? AND currency='USD' AND ticket_id IS NULL",
        (customer_id,)
    ).fetchone()[0]

    # How much of unalloc_pay was already absorbed into tickets?
    # Let's track opening_debt directly: opening_debt_tzs & opening_debt_usd are reduced directly when payment is recorded!
    open_tzs = float(cust["opening_debt_tzs"] or 0.0)
    open_usd = float(cust["opening_debt_usd"] or 0.0)

    return {
        "tzs": max(0.0, t_debt_tzs + open_tzs),
        "usd": max(0.0, t_debt_usd + open_usd),
        "ticket_tzs": t_debt_tzs,
        "ticket_usd": t_debt_usd,
        "opening_tzs": open_tzs,
        "opening_usd": open_usd,
    }

def re_evaluate_bsp_for_ticket(conn, ticket_number):
    """Automatically re-checks BSP billing rows whenever a ticket is added, edited, or deleted."""
    bsp_items = conn.execute("SELECT * FROM bsp_billing WHERE ticket_number=?", (ticket_number,)).fetchall()
    sys_t = conn.execute("SELECT * FROM tickets WHERE ticket_number=? AND status='ISSUED'", (ticket_number,)).fetchone()
    if not sys_t:
        conn.execute("UPDATE bsp_billing SET match_status='MISSING_IN_SYSTEM' WHERE ticket_number=? AND match_status != 'DUPLICATE'", (ticket_number,))
        return
    sys_net = float(sys_t["net_fare"]) + float(sys_t["taxes"])
    any_matched = False
    for b in bsp_items:
        if b["match_status"] == "DUPLICATE":
            continue
        if abs(sys_net - float(b["bsp_amount"])) < 1.0:
            conn.execute("UPDATE bsp_billing SET match_status='MATCHED' WHERE id=?", (b["id"],))
            any_matched = True
        else:
            conn.execute("UPDATE bsp_billing SET match_status='AMOUNT_MISMATCH' WHERE id=?", (b["id"],))
    conn.execute("UPDATE tickets SET bsp_reconciled=? WHERE id=?", (1 if any_matched else 0, sys_t["id"]))

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

@app.context_processor
def inject_globals():
    return {
        "current_user": session.get("full_name"),
        "current_role": session.get("role"),
        "current_branch": session.get("branch"),
        "today": date.today().strftime("%Y-%m-%d")
    }


@app.before_request
def enforce_role_access():
    """Limit Boss to reports and Ticket User to ticket entry only."""
    if "user_id" not in session:
        return None
    role = session.get("role")
    endpoint = request.endpoint or ""
    common = {"logout", "change_password", "static"}
    if role == "boss":
        allowed = common | {"financial_reports", "export_csv"}
        if endpoint not in allowed:
            flash("Akaunti ya Boss imewekewa ruhusa ya kuona ripoti pekee.", "warning")
            return redirect(url_for("financial_reports"))
    elif role == "ticket_user":
        allowed = common | {"ticket_management"}
        if endpoint not in allowed:
            flash("Akaunti hii inaruhusiwa kusajili tiketi pekee.", "warning")
            return redirect(url_for("ticket_management"))
        if endpoint == "ticket_management" and request.method == "POST":
            if request.form.get("action", "create") != "create":
                flash("Mtumiaji ana ruhusa ya kusajili tiketi mpya tu.", "danger")
                return redirect(url_for("ticket_management"))

# ---------------- API ENDPOINT FOR LIVE CUSTOMER DATA ----------------
@app.route("/api/customer/<int:cid>")
@login_required
def api_customer_info(cid):
    conn = get_db()
    cust = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    if not cust:
        conn.close()
        return jsonify({"error": "Not found"}), 404
    debt = get_customer_debt_summary(conn, cid)
    tickets = conn.execute(
        """SELECT id, ticket_date, ticket_number, airline, route, currency, total_selling, amount_paid, payment_status
           FROM tickets WHERE customer_id=? AND status='ISSUED' ORDER BY ticket_date DESC, id DESC""",
        (cid,)
    ).fetchall()
    conn.close()
    return jsonify({
        "id": cust["id"],
        "name": cust["name"],
        "category": cust["category"],
        "phone": cust["phone"] or "",
        "email": cust["email"] or "",
        "branch": cust["branch"],
        "credit_limit": float(cust["credit_limit"] or 0),
        "credit_balance_tzs": float(cust["credit_balance_tzs"] or 0),
        "credit_balance_usd": float(cust["credit_balance_usd"] or 0),
        "debt_tzs": debt["tzs"],
        "debt_usd": debt["usd"],
        "opening_tzs": debt["opening_tzs"],
        "opening_usd": debt["opening_usd"],
        "tickets": [dict(t) for t in tickets]
    })

# ---------------- FIRST ADMIN SETUP / LOGIN ----------------
@app.route("/setup", methods=["GET", "POST"])
def first_admin_setup():
    conn = get_db()
    existing = conn.execute("SELECT 1 FROM users LIMIT 1").fetchone()
    conn.close()
    if existing:
        return redirect(url_for("login"))
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        if not full_name or len(username) < 3 or not email:
            flash("Jaza jina, username (angalau herufi 3) na email sahihi.", "danger")
        elif not valid_password(password):
            flash("Password iwe na angalau herufi 9.", "danger")
        elif password != confirm:
            flash("Password hazifanani.", "danger")
        else:
            conn = get_db()
            # Re-check inside the write transaction to prevent a second initial admin.
            try:
                conn.execute("BEGIN IMMEDIATE")
                if conn.execute("SELECT 1 FROM users LIMIT 1").fetchone():
                    conn.rollback()
                    return redirect(url_for("login"))
                conn.execute(
                    "INSERT INTO users (username, full_name, email, role, branch, password_hash) VALUES (?, ?, ?, 'admin', 'ALL', ?)",
                    (username, full_name, email, generate_password_hash(password))
                )
                conn.commit()
                flash("Akaunti ya Admin imetengenezwa. Sasa ingia kwa username na password yako.", "success")
                return redirect(url_for("login"))
            except sqlite3.IntegrityError:
                conn.rollback()
                flash("Username au email tayari imetumika. Jaribu nyingine.", "danger")
            finally:
                conn.close()
    return render_template("first_admin_setup.html")

# ---------------- LOGIN / LOGOUT ----------------
@app.route("/login", methods=["GET", "POST"])
def login():
    conn = get_db()
    has_users = conn.execute("SELECT 1 FROM users LIMIT 1").fetchone()
    conn.close()
    if not has_users:
        return redirect(url_for("first_admin_setup"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE username = ? AND is_active = 1", (username,)).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["full_name"] = user["full_name"]
            session["role"] = user["role"]
            session["branch"] = user["branch"]
            log_audit(conn, user["id"], user["username"], "LOGIN", "AUTH", user["id"], "User logged into system")
            conn.commit()
            conn.close()
            if user["role"] == "boss":
                return redirect(url_for("financial_reports"))
            if user["role"] == "ticket_user":
                return redirect(url_for("ticket_management"))
            return redirect(url_for("dashboard"))
        conn.close()
        flash("Username au Password si sahihi.", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/change-password", methods=["GET", "POST"])
@login_required
def change_password():
    if request.method == "POST":
        current = request.form.get("current_password", "")
        new_pw = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
        if not user or not check_password_hash(user["password_hash"], current):
            flash("Password ya sasa si sahihi.", "danger")
        elif not valid_password(new_pw):
            flash("Password mpya iwe na angalau herufi 9.", "danger")
        elif new_pw != confirm:
            flash("Password mpya na uthibitisho havifanani.", "danger")
        else:
            conn.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(new_pw), user["id"]))
            conn.execute("UPDATE password_reset_tokens SET used_at=? WHERE user_id=? AND used_at IS NULL", (datetime.utcnow().isoformat(), user["id"]))
            log_audit(conn, user["id"], user["username"], "PASSWORD_CHANGE", "AUTH", user["id"], "User changed own password")
            conn.commit()
            flash("Password imebadilishwa kikamilifu.", "success")
        conn.close()
        return redirect(url_for("change_password"))
    return render_template("change_password.html")

@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        conn = get_db()
        user = conn.execute("SELECT * FROM users WHERE lower(email)=? AND is_active=1", (email,)).fetchone()
        sent = False
        if user and email:
            raw_token = secrets.token_urlsafe(32)
            token_hash = hashlib.sha256(raw_token.encode()).hexdigest()
            expires = (datetime.utcnow() + timedelta(minutes=30)).isoformat()
            conn.execute("UPDATE password_reset_tokens SET used_at=? WHERE user_id=? AND used_at IS NULL", (datetime.utcnow().isoformat(), user["id"]))
            conn.execute("INSERT INTO password_reset_tokens (user_id, token_hash, expires_at) VALUES (?, ?, ?)", (user["id"], token_hash, expires))
            conn.commit()
            base = os.environ.get("PUBLIC_BASE_URL", request.url_root.rstrip("/"))
            reset_url = f"{base}{url_for('reset_password', token=raw_token)}"
            try:
                sent = send_password_reset_email(email, reset_url)
            except Exception:
                app.logger.exception("Password reset email could not be sent")
        conn.close()
        flash("Ikiwa email hiyo imesajiliwa na huduma ya email imewekwa, utapokea link ya kubadili password. Angalia inbox/spam; usipopokea, wasiliana na Admin.", "success")
        return redirect(url_for("forgot_password"))
    return render_template("forgot_password.html")

@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    token_hash = hashlib.sha256(token.encode()).hexdigest()
    conn = get_db()
    reset = conn.execute(
        "SELECT pr.*, u.username FROM password_reset_tokens pr JOIN users u ON u.id=pr.user_id WHERE pr.token_hash=? AND pr.used_at IS NULL AND pr.expires_at>? AND u.is_active=1",
        (token_hash, datetime.utcnow().isoformat())
    ).fetchone()
    if not reset:
        conn.close()
        flash("Link ya kubadili password si sahihi au imeisha muda. Omba link mpya.", "danger")
        return redirect(url_for("forgot_password"))
    if request.method == "POST":
        new_pw = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")
        if not valid_password(new_pw):
            flash("Password mpya iwe na angalau herufi 9.", "danger")
        elif new_pw != confirm:
            flash("Password mpya na uthibitisho havifanani.", "danger")
        else:
            conn.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(new_pw), reset["user_id"]))
            conn.execute("UPDATE password_reset_tokens SET used_at=? WHERE id=?", (datetime.utcnow().isoformat(), reset["id"]))
            log_audit(conn, reset["user_id"], reset["username"], "PASSWORD_RESET", "AUTH", reset["user_id"], "Password reset using one-time email link")
            conn.commit()
            conn.close()
            flash("Password imebadilishwa. Sasa unaweza kuingia.", "success")
            return redirect(url_for("login"))
    conn.close()
    return render_template("reset_password.html")

# ---------------- 01. DASHBOARD ----------------
@app.route("/")
@login_required
def dashboard():
    conn = get_db()
    bank_tzs = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out), 0) FROM bank_transactions WHERE currency='TZS'").fetchone()[0]
    bank_usd = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out), 0) FROM bank_transactions WHERE currency='USD'").fetchone()[0]

    dar_cash_tzs = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out), 0) FROM cash_transactions WHERE book_type='DAR_CASH' AND currency='TZS'").fetchone()[0]
    dar_cash_usd = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out), 0) FROM cash_transactions WHERE book_type='DAR_CASH' AND currency='USD'").fetchone()[0]
    znz_cash_tzs = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out), 0) FROM cash_transactions WHERE book_type='UNGUJA_CASH' AND currency='TZS'").fetchone()[0]
    znz_cash_usd = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out), 0) FROM cash_transactions WHERE book_type='UNGUJA_CASH' AND currency='USD'").fetchone()[0]
    lipa_namba_tzs = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out), 0) FROM cash_transactions WHERE book_type='LIPA_NAMBA' AND currency='TZS'").fetchone()[0]

    t_debt_tzs = conn.execute("SELECT COALESCE(SUM(total_selling - amount_paid), 0) FROM tickets WHERE currency='TZS' AND status='ISSUED'").fetchone()[0]
    t_debt_usd = conn.execute("SELECT COALESCE(SUM(total_selling - amount_paid), 0) FROM tickets WHERE currency='USD' AND status='ISSUED'").fetchone()[0]
    c_open_tzs = conn.execute("SELECT COALESCE(SUM(opening_debt_tzs), 0) FROM customers").fetchone()[0]
    c_open_usd = conn.execute("SELECT COALESCE(SUM(opening_debt_usd), 0) FROM customers").fetchone()[0]
    debt_tzs = t_debt_tzs + c_open_tzs
    debt_usd = t_debt_usd + c_open_usd

    iata_billed_tzs = conn.execute("SELECT COALESCE(SUM(net_fare + taxes), 0) FROM tickets WHERE currency='TZS' AND status='ISSUED'").fetchone()[0]
    iata_paid_tzs = conn.execute("SELECT COALESCE(SUM(amount_paid), 0) FROM bsp_payments WHERE currency='TZS'").fetchone()[0]
    iata_payable_tzs = iata_billed_tzs - iata_paid_tzs

    iata_billed_usd = conn.execute("SELECT COALESCE(SUM(net_fare + taxes), 0) FROM tickets WHERE currency='USD' AND status='ISSUED'").fetchone()[0]
    iata_paid_usd = conn.execute("SELECT COALESCE(SUM(amount_paid), 0) FROM bsp_payments WHERE currency='USD'").fetchone()[0]
    iata_payable_usd = iata_billed_usd - iata_paid_usd

    dir_in_tzs = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM director_transfers WHERE currency='TZS' AND direction='IN_TO_COMPANY'").fetchone()[0]
    dir_out_tzs = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM director_transfers WHERE currency='TZS' AND direction='OUT_FROM_COMPANY'").fetchone()[0]
    dir_bal_tzs = dir_in_tzs - dir_out_tzs

    dir_in_usd = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM director_transfers WHERE currency='USD' AND direction='IN_TO_COMPANY'").fetchone()[0]
    dir_out_usd = conn.execute("SELECT COALESCE(SUM(amount), 0) FROM director_transfers WHERE currency='USD' AND direction='OUT_FROM_COMPANY'").fetchone()[0]
    dir_bal_usd = dir_in_usd - dir_out_usd

    unexplained_count = conn.execute("SELECT COUNT(*) FROM director_transfers WHERE status='UNEXPLAINED'").fetchone()[0]
    pending_expenses = conn.execute("SELECT COUNT(*) FROM expenses WHERE status='PENDING'").fetchone()[0]
    advance_out_tzs = conn.execute("SELECT COALESCE(SUM(a.amount),0)-COALESCE((SELECT SUM(r.amount) FROM bank_advance_repayments r JOIN bank_advances aa ON aa.id=r.advance_id WHERE aa.currency='TZS'),0) FROM bank_advances a WHERE a.currency='TZS'").fetchone()[0]
    advance_out_usd = conn.execute("SELECT COALESCE(SUM(a.amount),0)-COALESCE((SELECT SUM(r.amount) FROM bank_advance_repayments r JOIN bank_advances aa ON aa.id=r.advance_id WHERE aa.currency='USD'),0) FROM bank_advances a WHERE a.currency='USD'").fetchone()[0]
    total_customers = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]

    # Reconciliation overview: cash and bank balances remain separate; advances
    # are shown separately because they have left the bank but are still accountable.
    cash_balance_tzs = dar_cash_tzs + znz_cash_tzs + lipa_namba_tzs
    cash_balance_usd = dar_cash_usd + znz_cash_usd
    total_liquid_tzs = bank_tzs + cash_balance_tzs
    total_liquid_usd = bank_usd + cash_balance_usd
    accounted_funds_tzs = total_liquid_tzs + advance_out_tzs
    accounted_funds_usd = total_liquid_usd + advance_out_usd
    bank_unreconciled_count = conn.execute("SELECT COUNT(*) FROM bank_transactions WHERE is_reconciled=0").fetchone()[0]
    bank_unreconciled_tzs = conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM bank_transactions WHERE currency='TZS' AND is_reconciled=0").fetchone()[0]
    bank_unreconciled_usd = conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM bank_transactions WHERE currency='USD' AND is_reconciled=0").fetchone()[0]
    # Agency income is shown separately for transparency. Do NOT add it again
    # to NFP: amounts already received are in cash/bank, while unpaid ticket
    # balances are already included in customer receivables.
    agency_service_tzs = conn.execute(
        "SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE currency='TZS' AND status='ISSUED'"
    ).fetchone()[0]
    agency_commission_tzs = conn.execute(
        "SELECT COALESCE(SUM(commission),0) FROM tickets WHERE currency='TZS' AND status='ISSUED'"
    ).fetchone()[0]
    agency_service_usd = conn.execute(
        "SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE currency='USD' AND status='ISSUED'"
    ).fetchone()[0]
    agency_commission_usd = conn.execute(
        "SELECT COALESCE(SUM(commission),0) FROM tickets WHERE currency='USD' AND status='ISSUED'"
    ).fetchone()[0]
    # Separate extra-profit accounts. All tickets remain in the IATA BSP ledger.
    non_iata_profit_tzs = conn.execute(
        "SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE client_type='NON-IATA' AND currency='TZS' AND status='ISSUED'"
    ).fetchone()[0]
    non_iata_profit_usd = conn.execute(
        "SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE client_type='NON-IATA' AND currency='USD' AND status='ISSUED'"
    ).fetchone()[0]
    working_profit_tzs = conn.execute(
        "SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE client_type IN ('WORKING CLIENT','WALKING') AND currency='TZS' AND status='ISSUED'"
    ).fetchone()[0]
    working_profit_usd = conn.execute(
        "SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE client_type IN ('WORKING CLIENT','WALKING') AND currency='USD' AND status='ISSUED'"
    ).fetchone()[0]

    # Profit breakdown by branch and currency. TZS and USD remain separate.
    def profit_sum(client_types, branch, currency):
        marks = ",".join("?" for _ in client_types)
        sql = f"SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE client_type IN ({marks}) AND branch=? AND currency=? AND status='ISSUED'"
        return conn.execute(sql, (*client_types, branch, currency)).fetchone()[0]

    non_types = ("NON-IATA",)
    working_types = ("WORKING CLIENT", "WALKING")
    corporate_types = ("CORPORATE", "CORPORATE CLIENT", "HAWAII CORPORATE")
    dar_non_tzs = profit_sum(non_types, "DAR", "TZS")
    dar_non_usd = profit_sum(non_types, "DAR", "USD")
    unguja_non_tzs = profit_sum(non_types, "UNGUJA", "TZS")
    unguja_non_usd = profit_sum(non_types, "UNGUJA", "USD")
    dar_work_tzs = profit_sum(working_types, "DAR", "TZS")
    dar_work_usd = profit_sum(working_types, "DAR", "USD")
    unguja_work_tzs = profit_sum(working_types, "UNGUJA", "TZS")
    unguja_work_usd = profit_sum(working_types, "UNGUJA", "USD")
    dar_corporate_tzs = profit_sum(corporate_types, "DAR", "TZS")
    dar_corporate_usd = profit_sum(corporate_types, "DAR", "USD")
    unguja_corporate_tzs = profit_sum(corporate_types, "UNGUJA", "TZS")
    unguja_corporate_usd = profit_sum(corporate_types, "UNGUJA", "USD")
    corporate_profit_tzs = conn.execute(
        "SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE REPLACE(REPLACE(UPPER(TRIM(client_type)), ' ', ''), '-', '') IN ('CORPORATE','COPORATE','CORPORATECLIENT','COPORATECLIENT','HAWAIICORPORATE','HAWAIICOPORATE','CORPORATES/FEE','COPORATES/FEE') AND currency='TZS' AND status='ISSUED'"
    ).fetchone()[0]
    corporate_profit_usd = conn.execute(
        "SELECT COALESCE(SUM(service_charge),0) FROM tickets WHERE REPLACE(REPLACE(UPPER(TRIM(client_type)), ' ', ''), '-', '') IN ('CORPORATE','COPORATE','CORPORATECLIENT','COPORATECLIENT','HAWAIICORPORATE','HAWAIICOPORATE','CORPORATES/FEE','COPORATES/FEE') AND currency='USD' AND status='ISSUED'"
    ).fetchone()[0]

    dar_total_profit_tzs = dar_non_tzs + dar_work_tzs + dar_corporate_tzs
    dar_total_profit_usd = dar_non_usd + dar_work_usd + dar_corporate_usd
    unguja_total_profit_tzs = unguja_non_tzs + unguja_work_tzs + unguja_corporate_tzs
    unguja_total_profit_usd = unguja_non_usd + unguja_work_usd + unguja_corporate_usd
    all_total_profit_tzs = dar_total_profit_tzs + unguja_total_profit_tzs
    all_total_profit_usd = dar_total_profit_usd + unguja_total_profit_usd

    # Net Financial Position: liquid funds + customer receivables + staff advances
    # + director receivable, less IATA BSP and director payable.
    director_receivable_tzs = abs(dir_bal_tzs) if dir_bal_tzs < 0 else 0
    director_payable_tzs = dir_bal_tzs if dir_bal_tzs > 0 else 0
    director_receivable_usd = abs(dir_bal_usd) if dir_bal_usd < 0 else 0
    director_payable_usd = dir_bal_usd if dir_bal_usd > 0 else 0
    net_pos_tzs = (total_liquid_tzs + debt_tzs + max(advance_out_tzs, 0)
                   + director_receivable_tzs - iata_payable_tzs - director_payable_tzs)
    net_pos_usd = (total_liquid_usd + debt_usd + max(advance_out_usd, 0)
                   + director_receivable_usd - iata_payable_usd - director_payable_usd)

    recent_tickets = conn.execute(
        """SELECT t.*, c.name AS customer_name FROM tickets t
           LEFT JOIN customers c ON t.customer_id = c.id
           ORDER BY t.id DESC LIMIT 6"""
    ).fetchall()
    recent_audits = conn.execute("SELECT * FROM audit_trail ORDER BY id DESC LIMIT 6").fetchall()
    conn.close()

    return render_template(
        "dashboard.html",
        bank_tzs=bank_tzs, bank_usd=bank_usd,
        dar_cash_tzs=dar_cash_tzs, dar_cash_usd=dar_cash_usd,
        znz_cash_tzs=znz_cash_tzs, znz_cash_usd=znz_cash_usd,
        lipa_namba_tzs=lipa_namba_tzs,
        debt_tzs=debt_tzs, debt_usd=debt_usd,
        iata_payable_tzs=iata_payable_tzs, iata_payable_usd=iata_payable_usd,
        dir_bal_tzs=dir_bal_tzs, dir_bal_usd=dir_bal_usd,
        unexplained_count=unexplained_count, pending_expenses=pending_expenses,
        advance_out_tzs=advance_out_tzs, advance_out_usd=advance_out_usd,
        cash_balance_tzs=cash_balance_tzs, cash_balance_usd=cash_balance_usd,
        accounted_funds_tzs=accounted_funds_tzs, accounted_funds_usd=accounted_funds_usd,
        bank_unreconciled_count=bank_unreconciled_count,
        bank_unreconciled_tzs=bank_unreconciled_tzs, bank_unreconciled_usd=bank_unreconciled_usd,
        total_customers=total_customers,
        net_pos_tzs=net_pos_tzs, net_pos_usd=net_pos_usd,
        agency_service_tzs=agency_service_tzs, agency_commission_tzs=agency_commission_tzs,
        agency_service_usd=agency_service_usd, agency_commission_usd=agency_commission_usd,
        non_iata_profit_tzs=non_iata_profit_tzs, non_iata_profit_usd=non_iata_profit_usd,
        working_profit_tzs=working_profit_tzs, working_profit_usd=working_profit_usd,
        corporate_profit_tzs=corporate_profit_tzs, corporate_profit_usd=corporate_profit_usd,
        dar_non_tzs=dar_non_tzs, dar_non_usd=dar_non_usd,
        unguja_non_tzs=unguja_non_tzs, unguja_non_usd=unguja_non_usd,
        dar_work_tzs=dar_work_tzs, dar_work_usd=dar_work_usd,
        unguja_work_tzs=unguja_work_tzs, unguja_work_usd=unguja_work_usd,
        dar_corporate_tzs=dar_corporate_tzs, dar_corporate_usd=dar_corporate_usd,
        unguja_corporate_tzs=unguja_corporate_tzs, unguja_corporate_usd=unguja_corporate_usd,
        dar_total_profit_tzs=dar_total_profit_tzs, dar_total_profit_usd=dar_total_profit_usd,
        unguja_total_profit_tzs=unguja_total_profit_tzs, unguja_total_profit_usd=unguja_total_profit_usd,
        all_total_profit_tzs=all_total_profit_tzs, all_total_profit_usd=all_total_profit_usd,
        director_receivable_tzs=director_receivable_tzs, director_payable_tzs=director_payable_tzs,
        director_receivable_usd=director_receivable_usd, director_payable_usd=director_payable_usd,
        recent_tickets=recent_tickets, recent_audits=recent_audits
    )

# ---------------- 02. BANK MANAGEMENT ----------------
@app.route("/bank", methods=["GET", "POST"])
@login_required
def bank_management():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        if action in ("add_tx", "edit_tx"):
            tx_date = request.form["tx_date"]
            currency = request.form["currency"]
            branch = request.form["branch"]
            tx_type = request.form["tx_type"]
            category = request.form["category"]
            reference = request.form.get("reference", "").strip()
            description = request.form["description"].strip()
            amount = parse_num(request.form.get("amount"))

            amount_in = amount if tx_type in ("DEPOSIT", "TRANSFER_IN") else 0.0
            amount_out = amount if tx_type in ("WITHDRAWAL", "BANK_CHARGE", "BSP_PAYMENT", "TRANSFER_OUT") else 0.0

            if action == "add_tx":
                cur = conn.execute(
                    """INSERT INTO bank_transactions
                       (tx_date, currency, branch, tx_type, category, reference, description, amount_in, amount_out, created_by)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (tx_date, currency, branch, tx_type, category, reference, description, amount_in, amount_out, session["username"])
                )
                log_audit(conn, session["user_id"], session["username"], "CREATE", "BANK", cur.lastrowid, f"{currency} {tx_type} {amount}: {description}")
                flash("Umefanikiwa kuhifadhi muamala wa benki.", "success")
            else:
                tx_id = request.form["tx_id"]
                conn.execute(
                    """UPDATE bank_transactions
                       SET tx_date=?, currency=?, branch=?, tx_type=?, category=?, reference=?, description=?, amount_in=?, amount_out=?
                       WHERE id=?""",
                    (tx_date, currency, branch, tx_type, category, reference, description, amount_in, amount_out, tx_id)
                )
                log_audit(conn, session["user_id"], session["username"], "EDIT", "BANK", tx_id, f"Edited bank tx #{tx_id}: {currency} {amount}")
                flash("Umefanikiwa kurekebisha muamala wa benki.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("bank_management"))

        elif action == "delete_tx":
            tx_id = request.form["tx_id"]
            conn.execute("DELETE FROM bank_transactions WHERE id=?", (tx_id,))
            log_audit(conn, session["user_id"], session["username"], "DELETE", "BANK", tx_id, f"Deleted bank tx #{tx_id}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta muamala wa benki.", "success")
            return redirect(url_for("bank_management"))

        elif action == "reconcile_toggle":
            tx_id = request.form["tx_id"]
            st_ref = request.form.get("statement_ref", "MATCHED")
            conn.execute("UPDATE bank_transactions SET is_reconciled = CASE WHEN is_reconciled=1 THEN 0 ELSE 1 END, statement_ref=? WHERE id=?", (st_ref, tx_id))
            log_audit(conn, session["user_id"], session["username"], "RECONCILE", "BANK", tx_id, f"Toggled bank reconciliation ({st_ref})")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kusasisha Bank Reconciliation.", "success")
            return redirect(url_for("bank_management"))

    currency_filter = request.args.get("currency", "ALL")
    query = "SELECT * FROM bank_transactions"
    params = []
    if currency_filter in ("TZS", "USD"):
        query += " WHERE currency = ?"
        params.append(currency_filter)
    query += " ORDER BY tx_date DESC, id DESC"
    txs = conn.execute(query, params).fetchall()

    bal_tzs = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out),0) FROM bank_transactions WHERE currency='TZS'").fetchone()[0]
    bal_usd = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out),0) FROM bank_transactions WHERE currency='USD'").fetchone()[0]
    rec_tzs = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out),0) FROM bank_transactions WHERE currency='TZS' AND is_reconciled=1").fetchone()[0]
    rec_usd = conn.execute("SELECT COALESCE(SUM(amount_in - amount_out),0) FROM bank_transactions WHERE currency='USD' AND is_reconciled=1").fetchone()[0]
    conn.close()

    return render_template(
        "bank.html", txs=txs, currency_filter=currency_filter,
        bal_tzs=bal_tzs, bal_usd=bal_usd, rec_tzs=rec_tzs, rec_usd=rec_usd
    )

# ---------------- 02A. BANK ADVANCES / MONEY OUTSTANDING ----------------
@app.route("/bank-advances", methods=["GET", "POST"])
@login_required
def bank_advances():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add_advance":
            adv_date = request.form["advance_date"]
            branch = request.form["branch"]
            currency = request.form["currency"]
            amount = parse_num(request.form.get("amount"))
            taken_by = request.form["taken_by"].strip()
            purpose = request.form["purpose"].strip()
            reference = request.form.get("reference", "").strip()
            if amount <= 0 or not taken_by:
                flash("Weka kiasi sahihi na jina la aliyechukua pesa.", "danger")
            else:
                cur = conn.execute("""INSERT INTO bank_advances
                    (advance_date, branch, currency, amount, taken_by, purpose, reference, created_by)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (adv_date, branch, currency, amount, taken_by, purpose, reference, session["username"]))
                adv_id = cur.lastrowid
                bank_cur = conn.execute("""INSERT INTO bank_transactions
                    (tx_date, currency, branch, tx_type, category, reference, description, amount_in, amount_out, source_module, source_id, created_by)
                    VALUES (?, ?, ?, 'WITHDRAWAL', 'STAFF_ADVANCE', ?, ?, 0, ?, 'BANK_ADVANCE', ?, ?)""",
                    (adv_date, currency, branch, reference, f"Pesa alizochukua {taken_by}: {purpose}", amount, adv_id, session["username"]))
                conn.execute("UPDATE bank_advances SET bank_transaction_id=? WHERE id=?", (bank_cur.lastrowid, adv_id))
                log_audit(conn, session["user_id"], session["username"], "CREATE", "BANK_ADVANCE", adv_id, f"{taken_by} took {currency} {amount}")
                flash("Pesa iliyotolewa benki imehifadhiwa na deni lake linafuatiliwa.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("bank_advances"))
        elif action == "edit_advance":
            adv_id = int(request.form.get("advance_id", "0"))
            adv = conn.execute("SELECT * FROM bank_advances WHERE id=?", (adv_id,)).fetchone()
            new_date = request.form.get("advance_date", "").strip()
            new_amount = parse_num(request.form.get("amount"))
            new_taken_by = request.form.get("taken_by", "").strip()
            new_purpose = request.form.get("purpose", "").strip()
            new_reference = request.form.get("reference", "").strip()
            repaid = conn.execute(
                "SELECT COALESCE(SUM(amount),0) FROM bank_advance_repayments WHERE advance_id=?",
                (adv_id,)
            ).fetchone()[0] if adv else 0
            if not adv:
                flash("Rekodi ya pesa iliyotolewa haijapatikana.", "danger")
            elif not new_date or not new_taken_by or new_amount <= 0:
                flash("Weka tarehe, jina na kiasi sahihi kabla ya kuhifadhi.", "danger")
            elif new_amount + 0.00001 < float(repaid):
                flash("Kiasi kipya hakiwezi kuwa chini ya marejesho yaliyokwisha rekodiwa.", "danger")
            else:
                conn.execute("""UPDATE bank_advances
                    SET advance_date=?, amount=?, taken_by=?, purpose=?, reference=? WHERE id=?""",
                    (new_date, new_amount, new_taken_by, new_purpose, new_reference, adv_id))
                # Keep the original linked bank withdrawal in sync so Bank Balance
                # and the outstanding advance always use the same corrected amount.
                if adv["bank_transaction_id"]:
                    conn.execute("""UPDATE bank_transactions SET tx_date=?, reference=?,
                        description=?, amount_out=? WHERE id=? AND source_module='BANK_ADVANCE' AND source_id=?""",
                        (new_date, new_reference, f"Pesa alizochukua {new_taken_by}: {new_purpose}",
                         new_amount, adv["bank_transaction_id"], adv_id))
                else:
                    conn.execute("""UPDATE bank_transactions SET tx_date=?, reference=?,
                        description=?, amount_out=? WHERE source_module='BANK_ADVANCE' AND source_id=?""",
                        (new_date, new_reference, f"Pesa alizochukua {new_taken_by}: {new_purpose}", new_amount, adv_id))
                log_audit(conn, session["user_id"], session["username"], "UPDATE", "BANK_ADVANCE", adv_id,
                          f"Corrected advance for {new_taken_by}: {adv['currency']} {new_amount}")
                flash("Taarifa za pesa iliyochukuliwa zimehaririwa. Bank Balance na salio la deni vimesasishwa.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("bank_advances"))
        elif action == "repay_advance":
            adv_id = int(request.form["advance_id"])
            repayment_date = request.form["repayment_date"]
            amount = parse_num(request.form.get("amount"))
            received_into = request.form["received_into"]
            reference = request.form.get("reference", "").strip()
            notes = request.form.get("notes", "").strip()
            adv = conn.execute("SELECT * FROM bank_advances WHERE id=?", (adv_id,)).fetchone()
            paid = conn.execute("SELECT COALESCE(SUM(amount),0) FROM bank_advance_repayments WHERE advance_id=?", (adv_id,)).fetchone()[0]
            outstanding = float(adv["amount"]) - float(paid) if adv else 0
            if not adv or amount <= 0 or amount > outstanding + 0.00001:
                flash("Kiasi cha marejesho si sahihi au kinazidi deni lililobaki.", "danger")
            else:
                cur = conn.execute("""INSERT INTO bank_advance_repayments
                    (advance_id, repayment_date, amount, received_into, reference, notes, created_by)
                    VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (adv_id, repayment_date, amount, received_into, reference, notes, session["username"]))
                repayment_id = cur.lastrowid
                if received_into == "BANK":
                    bcur = conn.execute("""INSERT INTO bank_transactions
                        (tx_date, currency, branch, tx_type, category, reference, description, amount_in, amount_out, source_module, source_id, created_by)
                        VALUES (?, ?, ?, 'DEPOSIT', 'ADVANCE_RETURN', ?, ?, ?, 0, 'ADVANCE_REPAYMENT', ?, ?)""",
                        (repayment_date, adv["currency"], adv["branch"], reference, f"Marejesho kutoka {adv['taken_by']}", amount, repayment_id, session["username"]))
                    conn.execute("UPDATE bank_advance_repayments SET bank_transaction_id=? WHERE id=?", (bcur.lastrowid, repayment_id))
                else:
                    book = "DAR_CASH" if adv["branch"] == "DAR" else "UNGUJA_CASH"
                    ccur = conn.execute("""INSERT INTO cash_transactions
                        (tx_date, book_type, currency, tx_type, reference, description, amount_in, amount_out, source_module, source_id, created_by)
                        VALUES (?, ?, ?, 'RECEIPT', ?, ?, ?, 0, 'ADVANCE_REPAYMENT', ?, ?)""",
                        (repayment_date, book, adv["currency"], reference, f"Marejesho kutoka {adv['taken_by']}", amount, repayment_id, session["username"]))
                    conn.execute("UPDATE bank_advance_repayments SET cash_transaction_id=? WHERE id=?", (ccur.lastrowid, repayment_id))
                log_audit(conn, session["user_id"], session["username"], "REPAYMENT", "BANK_ADVANCE", adv_id, f"{adv['taken_by']} repaid {adv['currency']} {amount}")
                flash("Marejesho yamehifadhiwa na salio limesasishwa.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("bank_advances"))

    advances = conn.execute("""SELECT a.*, COALESCE(SUM(r.amount),0) AS repaid,
        a.amount-COALESCE(SUM(r.amount),0) AS outstanding
        FROM bank_advances a LEFT JOIN bank_advance_repayments r ON r.advance_id=a.id
        GROUP BY a.id ORDER BY a.advance_date DESC, a.id DESC""").fetchall()
    outstanding = {}
    for curr in ("TZS", "USD"):
        outstanding[curr] = conn.execute("""SELECT COALESCE(SUM(a.amount),0)-COALESCE((
            SELECT SUM(r.amount) FROM bank_advance_repayments r JOIN bank_advances aa ON aa.id=r.advance_id WHERE aa.currency=?),0)
            FROM bank_advances a WHERE a.currency=?""", (curr, curr)).fetchone()[0]
    repayments = conn.execute("""SELECT r.*, a.taken_by, a.currency FROM bank_advance_repayments r
        JOIN bank_advances a ON a.id=r.advance_id ORDER BY r.repayment_date DESC, r.id DESC LIMIT 100""").fetchall()
    conn.close()
    return render_template("bank_advances.html", advances=advances, repayments=repayments, outstanding=outstanding, today=date.today().isoformat())

# ---------------- 03. CASH MANAGEMENT ----------------
@app.route("/cash", methods=["GET", "POST"])
@login_required
def cash_management():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action", "add_cash")
        if action in ("add_cash", "edit_cash"):
            tx_date = request.form["tx_date"]
            book_type = request.form["book_type"]
            currency = request.form["currency"]
            tx_type = request.form["tx_type"]
            reference = request.form.get("reference", "").strip()
            description = request.form["description"].strip()
            amount = parse_num(request.form.get("amount"))

            amount_in = amount if tx_type == "RECEIPT" else 0.0
            amount_out = amount if tx_type in ("PAYMENT", "BANK_DEPOSIT") else 0.0
            branch_map = {"DAR_CASH": "DAR", "UNGUJA_CASH": "UNGUJA", "LIPA_NAMBA": "HQ"}

            if action == "add_cash":
                cur = conn.execute(
                    """INSERT INTO cash_transactions
                       (tx_date, book_type, currency, tx_type, reference, description, amount_in, amount_out, created_by)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (tx_date, book_type, currency, tx_type, reference, description, amount_in, amount_out, session["username"])
                )
                c_id = cur.lastrowid
                if tx_type == "BANK_DEPOSIT":
                    sync_ledger_entry(
                        conn, "CASH_DEPOSIT", c_id, tx_date, currency, branch_map.get(book_type, "DAR"),
                        "BANK", "IN", "CASH_DEPOSIT", reference, f"Cash Deposit from {book_type}: {description}", amount, session["username"]
                    )
                log_audit(conn, session["user_id"], session["username"], "CREATE", "CASH", c_id, f"{book_type} {tx_type} {currency} {amount}")
                flash("Umefanikiwa kuhifadhi muamala wa Cash / Lipa Namba.", "success")
            else:
                tx_id = int(request.form["tx_id"])
                conn.execute(
                    """UPDATE cash_transactions
                       SET tx_date=?, book_type=?, currency=?, tx_type=?, reference=?, description=?, amount_in=?, amount_out=?
                       WHERE id=?""",
                    (tx_date, book_type, currency, tx_type, reference, description, amount_in, amount_out, tx_id)
                )
                if tx_type == "BANK_DEPOSIT":
                    sync_ledger_entry(
                        conn, "CASH_DEPOSIT", tx_id, tx_date, currency, branch_map.get(book_type, "DAR"),
                        "BANK", "IN", "CASH_DEPOSIT", reference, f"Cash Deposit from {book_type}: {description}", amount, session["username"]
                    )
                else:
                    conn.execute("DELETE FROM bank_transactions WHERE source_module='CASH_DEPOSIT' AND source_id=?", (tx_id,))
                log_audit(conn, session["user_id"], session["username"], "EDIT", "CASH", tx_id, f"Edited cash tx #{tx_id}: {currency} {amount}")
                flash("Umefanikiwa kurekebisha muamala wa Cash.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("cash_management"))

        elif action == "delete_cash":
            tx_id = int(request.form["tx_id"])
            conn.execute("DELETE FROM cash_transactions WHERE id=?", (tx_id,))
            conn.execute("DELETE FROM bank_transactions WHERE source_module='CASH_DEPOSIT' AND source_id=?", (tx_id,))
            log_audit(conn, session["user_id"], session["username"], "DELETE", "CASH", tx_id, f"Deleted cash tx #{tx_id}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta muamala wa Cash.", "success")
            return redirect(url_for("cash_management"))

    book_filter = request.args.get("book", "ALL")
    q = "SELECT * FROM cash_transactions"
    p = []
    if book_filter != "ALL":
        q += " WHERE book_type = ?"
        p.append(book_filter)
    q += " ORDER BY tx_date DESC, id DESC"
    txs = conn.execute(q, p).fetchall()

    balances = {
        "dar_tzs": conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE book_type='DAR_CASH' AND currency='TZS'").fetchone()[0],
        "dar_usd": conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE book_type='DAR_CASH' AND currency='USD'").fetchone()[0],
        "znz_tzs": conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE book_type='UNGUJA_CASH' AND currency='TZS'").fetchone()[0],
        "znz_usd": conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE book_type='UNGUJA_CASH' AND currency='USD'").fetchone()[0],
        "lipa_tzs": conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE book_type='LIPA_NAMBA' AND currency='TZS'").fetchone()[0],
    }
    conn.close()
    return render_template("cash.html", txs=txs, book_filter=book_filter, balances=balances)

# ---------------- 04. TICKET MANAGEMENT ----------------
@app.route("/tickets", methods=["GET", "POST"])
@login_required
def ticket_management():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action", "create")
        if action in ("create", "edit"):
            ticket_date = request.form["ticket_date"]
            ticket_number = request.form["ticket_number"].strip()
            passenger_name = request.form["passenger_name"].strip()
            airline = request.form["airline"].strip().upper()
            route = request.form.get("route", "").strip().upper()
            client_type = request.form["client_type"]
            raw_cid = request.form.get("customer_id", "").strip()
            customer_id = int(raw_cid) if raw_cid else None
            branch = request.form["branch"]
            if session.get("role") == "ticket_user" and session.get("branch") in ("DAR", "UNGUJA"):
                branch = session.get("branch")
            currency = request.form["currency"]
            net_fare = parse_num(request.form.get("net_fare"))
            taxes = parse_num(request.form.get("taxes"))
            service_charge = parse_num(request.form.get("service_charge"))
            # Respect the service charge entered on the ticket form. The admin-configured
            # fee may prefill the form, but must never overwrite a user's manual amount.
            commission = parse_num(request.form.get("commission"))
            total_selling = net_fare + taxes + service_charge
            payment_method = request.form["payment_method"]
            amount_paid = parse_num(request.form.get("amount_paid"))

            # CORPORATE tickets must be linked to a pre-registered CORPORATE account.
            selected_customer = conn.execute("SELECT id, category, branch FROM customers WHERE id=?", (customer_id,)).fetchone() if customer_id else None
            if client_type == "CORPORATE":
                if not selected_customer or str(selected_customer["category"]).strip().upper() != "CORPORATE":
                    flash("Tiketi za CORPORATE zinahitaji kwanza kusajili kampuni kwenye Customer Accounts kama CORPORATE, kisha uchague akaunti yake hapa.", "danger")
                    conn.close()
                    return redirect(url_for("ticket_management"))
            elif customer_id and not selected_customer:
                flash("Mteja uliyemchagua hajasajiliwa. Tafadhali chagua mteja aliyepo.", "danger")
                conn.close()
                return redirect(url_for("ticket_management"))

            # Keep the paid amount consistent with the recalculated selling total.
            # CREDIT tickets cannot have a direct payment; other methods are capped
            # at the new total so edits cannot leave an impossible paid balance.
            if payment_method == "CREDIT":
                amount_paid = 0.0
                payment_status = "UNPAID"
            else:
                amount_paid = min(max(0.0, amount_paid), max(0.0, total_selling))
                if total_selling > 0 and amount_paid >= total_selling - 0.005:
                    amount_paid = total_selling
                    payment_status = "PAID"
                elif amount_paid > 0:
                    payment_status = "PARTIAL"
                else:
                    payment_status = "UNPAID"

            try:
                if action == "create":
                    cur = conn.execute(
                        """INSERT INTO tickets
                           (ticket_date, ticket_number, passenger_name, airline, route, client_type, customer_id,
                            branch, currency, net_fare, taxes, service_charge, commission, total_selling,
                            payment_method, payment_status, amount_paid, created_by)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (ticket_date, ticket_number, passenger_name, airline, route, client_type, customer_id,
                         branch, currency, net_fare, taxes, service_charge, commission, total_selling,
                         payment_method, payment_status, amount_paid, session["username"])
                    )
                    t_id = cur.lastrowid
                    sync_ledger_entry(
                        conn, "TICKET", t_id, ticket_date, currency, branch, payment_method, "IN",
                        "TICKET_SALE", ticket_number, f"Ticket {ticket_number} - {passenger_name}", amount_paid, session["username"]
                    )
                    re_evaluate_bsp_for_ticket(conn, ticket_number)
                    log_audit(conn, session["user_id"], session["username"], "CREATE", "TICKETS", t_id, f"Issued ticket {ticket_number} ({currency} {total_selling})")
                    flash(f"Umefanikiwa kusajili na kuhifadhi tiketi {ticket_number}.", "success")
                else:
                    t_id = int(request.form["ticket_id"])
                    old_t = conn.execute("SELECT ticket_number FROM tickets WHERE id=?", (t_id,)).fetchone()
                    conn.execute(
                        """UPDATE tickets
                           SET ticket_date=?, ticket_number=?, passenger_name=?, airline=?, route=?, client_type=?, customer_id=?,
                               branch=?, currency=?, net_fare=?, taxes=?, service_charge=?, commission=?, total_selling=?,
                               payment_method=?, payment_status=?, amount_paid=?
                           WHERE id=?""",
                        (ticket_date, ticket_number, passenger_name, airline, route, client_type, customer_id,
                         branch, currency, net_fare, taxes, service_charge, commission, total_selling,
                         payment_method, payment_status, amount_paid, t_id)
                    )
                    sync_ledger_entry(
                        conn, "TICKET", t_id, ticket_date, currency, branch, payment_method, "IN",
                        "TICKET_SALE", ticket_number, f"Ticket {ticket_number} - {passenger_name}", amount_paid, session["username"]
                    )
                    if old_t and old_t["ticket_number"] != ticket_number:
                        re_evaluate_bsp_for_ticket(conn, old_t["ticket_number"])
                    re_evaluate_bsp_for_ticket(conn, ticket_number)
                    log_audit(conn, session["user_id"], session["username"], "EDIT", "TICKETS", t_id, f"Edited ticket {ticket_number} ({currency} {total_selling})")
                    flash(f"Umefanikiwa kurekebisha tiketi {ticket_number}.", "success")
                conn.commit()
            except Exception as e:
                conn.rollback()
                flash(f"Hitilafu: {e}", "danger")
            conn.close()
            return redirect(url_for("ticket_management"))

        elif action == "delete":
            t_id = int(request.form["ticket_id"])
            old_t = conn.execute("SELECT ticket_number FROM tickets WHERE id=?", (t_id,)).fetchone()
            conn.execute("DELETE FROM tickets WHERE id=?", (t_id,))
            conn.execute("DELETE FROM bank_transactions WHERE source_module='TICKET' AND source_id=?", (t_id,))
            conn.execute("DELETE FROM cash_transactions WHERE source_module='TICKET' AND source_id=?", (t_id,))
            if old_t:
                re_evaluate_bsp_for_ticket(conn, old_t["ticket_number"])
            log_audit(conn, session["user_id"], session["username"], "DELETE", "TICKETS", t_id, f"Deleted ticket #{t_id}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta tiketi na miamala yake.", "success")
            return redirect(url_for("ticket_management"))

        elif action == "status_change":
            t_id = int(request.form["ticket_id"])
            new_status = request.form["new_status"]
            old_t = conn.execute("SELECT ticket_number FROM tickets WHERE id=?", (t_id,)).fetchone()
            conn.execute("UPDATE tickets SET status = ? WHERE id = ?", (new_status, t_id))
            if new_status != "ISSUED":
                conn.execute("DELETE FROM bank_transactions WHERE source_module='TICKET' AND source_id=?", (t_id,))
                conn.execute("DELETE FROM cash_transactions WHERE source_module='TICKET' AND source_id=?", (t_id,))
            if old_t:
                re_evaluate_bsp_for_ticket(conn, old_t["ticket_number"])
            log_audit(conn, session["user_id"], session["username"], "UPDATE", "TICKETS", t_id, f"Changed ticket status to {new_status}")
            conn.commit()
            conn.close()
            flash(f"Umefanikiwa kubadili hali ya tiketi kuwa {new_status}.", "success")
            return redirect(url_for("ticket_management"))

    # Date-range ticket register: users can choose any start/end dates from calendar controls.
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    branch_filter = request.args.get("branch", "ALL").strip().upper()
    currency_filter = request.args.get("currency", "ALL").strip().upper()
    search_term = request.args.get("q", "").strip()
    where = []
    params = []
    if date_from:
        where.append("t.ticket_date >= ?")
        params.append(date_from)
    if date_to:
        where.append("t.ticket_date <= ?")
        params.append(date_to)
    if branch_filter in ("DAR", "UNGUJA"):
        where.append("t.branch = ?")
        params.append(branch_filter)
    if currency_filter in ("TZS", "USD"):
        where.append("t.currency = ?")
        params.append(currency_filter)
    if search_term:
        like = f"%{search_term}%"
        where.append("(t.ticket_number LIKE ? OR t.passenger_name LIKE ? OR c.name LIKE ? OR t.airline LIKE ?)")
        params.extend([like, like, like, like])
    where_sql = ("WHERE " + " AND ".join(where)) if where else ""
    if session.get("role") == "ticket_user":
        tickets = []
    else:
        tickets = conn.execute(
            f"""SELECT t.*, c.name AS customer_name
                FROM tickets t LEFT JOIN customers c ON t.customer_id = c.id
                {where_sql}
                ORDER BY t.ticket_date ASC, t.id ASC""",
            params
        ).fetchall()
    totals = {}
    for curr in ("TZS", "USD"):
        totals[curr] = {
            "count": sum(1 for t in tickets if t["currency"] == curr),
            "selling": sum(float(t["total_selling"] or 0) for t in tickets if t["currency"] == curr),
            "paid": sum(float(t["amount_paid"] or 0) for t in tickets if t["currency"] == curr),
        }
    customers = conn.execute("SELECT * FROM customers ORDER BY name").fetchall()
    fee_tzs_row = conn.execute("SELECT setting_value FROM app_settings WHERE setting_key='non_iata_fee_tzs'").fetchone()
    fee_usd_row = conn.execute("SELECT setting_value FROM app_settings WHERE setting_key='non_iata_fee_usd'").fetchone()
    fee_tzs = float(fee_tzs_row[0]) if fee_tzs_row else 10000.0
    fee_usd = float(fee_usd_row[0]) if fee_usd_row else 10.0
    conn.close()
    return render_template("tickets.html", tickets=tickets, customers=customers,
                           date_from=date_from, date_to=date_to, branch_filter=branch_filter,
                           totals=totals, today=date.today().isoformat(),
                           non_iata_fee_tzs=fee_tzs, non_iata_fee_usd=fee_usd,
                           ticket_entry_only=(session.get("role") == "ticket_user"))



# ---------------- BRANCH TICKET EXCEL EXPORTS ----------------
def _ticket_export_filters(request):
    date_from = request.args.get("date_from", "").strip()
    date_to = request.args.get("date_to", "").strip()
    where, params = [], []
    if date_from:
        where.append("t.ticket_date >= ?")
        params.append(date_from)
    if date_to:
        where.append("t.ticket_date <= ?")
        params.append(date_to)
    where.append("t.status = 'ISSUED'")
    return date_from, date_to, where, params


def _excel_response(wb, filename):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def _excel_title(ws, last_col, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    cell = ws.cell(1, 1, title)
    cell.font = Font(bold=True, size=15, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="17365D")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    cell = ws.cell(2, 1, subtitle)
    cell.font = Font(italic=True, color="404040", size=10)
    cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 26


def _make_branch_ticket_excel(branch, date_from, date_to, tickets, all_branch_tickets):
    """Create a printable branch register and currency-safe profit summary."""
    wb = Workbook()
    ws = wb.active
    ws.title = "DAR+ZANZIBAR REGISTER" if branch == "DAR" else "ZANZIBAR REGISTER"

    headers = [
        "DATE", "CLIENT NAME / PASSENGER", "TICKET NUMBER", "AIRLINE", "ROUTE",
        "BSP TOTAL AMOUNT\nTZS", "BSP TOTAL AMOUNT\nUSD",
        "CUSTOMER TOTAL", "PAID",
        "HAWAII WORKING CLIENT S/FEE USD", "HAWAII WORKING CLIENT S/FEE TZS",
        "HAWAII CORPORATE S/FEE USD", "HAWAII CORPORATE S/FEE TZS",
        "NON-IATA S/FEE USD", "NON-IATA S/FEE TZS", "BRANCH",
    ]
    register_scope = "DAR + ZANZIBAR" if branch == "DAR" else "ZANZIBAR"
    _excel_title(
        ws, len(headers), f"HAWAII TRAVEL AGENCY - {register_scope} TICKET REGISTER",
        f"Issued tickets | Included branches: {register_scope} | Period: {date_from or 'All dates'} to {date_to or 'All dates'} | BSP amount excludes service fee"
    )
    header_row = 4
    for col, heading in enumerate(headers, 1):
        c = ws.cell(header_row, col, heading)
        c.font = Font(bold=True, color="000000")
        c.fill = PatternFill("solid", fgColor="A9D18E" if col <= 9 else ("F4B183" if col <= 11 else "FFE699"))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def val(row, key, default=None):
        try:
            return row[key]
        except (KeyError, IndexError, TypeError):
            return default

    def category(row):
        # Normalize legacy/manual spellings so older CORPORATE records are not
        # silently dropped from the USD/TZS fee columns or Profit Summary.
        raw = str(val(row, "client_type", "") or "").strip().upper()
        compact = "".join(ch for ch in raw if ch.isalnum())
        if compact in ("WORKINGCLIENT", "WALKING", "HAWAIIWORKINGCLIENT"):
            return "WORKING CLIENT"
        # COPORATE is a known historical misspelling found in older entries.
        if compact in (
            "CORPORATE", "COPORATE", "CORPORATECLIENT", "COPORATECLIENT",
            "HAWAIICORPORATE", "HAWAIICOPORATE", "CORPORATESFEE", "COPORATESFEE",
        ):
            return "CORPORATE"
        if compact in ("NONIATA", "NONIATAAGENT", "AGENT"):
            return "NON-IATA"
        return raw

    def normalized_currency(row):
        raw = str(val(row, "currency", "") or "").strip().upper().replace(" ", "")
        if raw in ("TZS", "TSH", "T.SH", "TANZANIANSHILLING", "TANZANIANSHILLINGS"):
            return "TZS"
        if raw in ("USD", "US$", "$", "US DOLLAR", "US DOLLARS"):
            return "USD"
        return raw

    def normalized_branch(row):
        raw = str(val(row, "branch", "") or "").strip().upper()
        if raw in ("UNGUJA", "ZANZIBAR", "ZNZ"):
            return "UNGUJA"
        return raw

    def summarize(rows):
        out = {"WORKING CLIENT TZS": 0.0, "WORKING CLIENT USD": 0.0,
               "CORPORATE TZS": 0.0, "CORPORATE USD": 0.0,
               "NON-IATA TZS": 0.0, "NON-IATA USD": 0.0}
        for item in rows:
            curr = normalized_currency(item)
            typ = category(item)
            if curr in ("TZS", "USD") and typ in ("WORKING CLIENT", "CORPORATE", "NON-IATA"):
                out[f"{typ} {curr}"] += float(val(item, "service_charge", 0) or 0)
        out["TOTAL PROFIT TZS"] = out["WORKING CLIENT TZS"] + out["CORPORATE TZS"] + out["NON-IATA TZS"]
        out["TOTAL PROFIT USD"] = out["WORKING CLIENT USD"] + out["CORPORATE USD"] + out["NON-IATA USD"]
        return out

    for row_no, t in enumerate(tickets, header_row + 1):
        currency = normalized_currency(t)
        typ = category(t)
        bsp_amount = float(val(t, "net_fare", 0) or 0) + float(val(t, "taxes", 0) or 0)
        bsp_tzs = bsp_amount if currency == "TZS" else None
        bsp_usd = bsp_amount if currency == "USD" else None
        fee = float(val(t, "service_charge", 0) or 0)
        working_usd = fee if typ == "WORKING CLIENT" and currency == "USD" else None
        working_tzs = fee if typ == "WORKING CLIENT" and currency == "TZS" else None
        corporate_usd = fee if typ == "CORPORATE" and currency == "USD" else None
        corporate_tzs = fee if typ == "CORPORATE" and currency == "TZS" else None
        agent_usd = fee if typ == "NON-IATA" and currency == "USD" else None
        agent_tzs = fee if typ == "NON-IATA" and currency == "TZS" else None
        passenger = val(t, "passenger_name", "") or ""
        customer = val(t, "customer_name", "") or "Walk-in"
        values = [
            val(t, "ticket_date", ""), f"{customer} / {passenger}",
            val(t, "ticket_number", ""), val(t, "airline", ""), val(t, "route", ""),
            bsp_tzs, bsp_usd,
            float(val(t, "total_selling", 0) or 0),
            float(val(t, "amount_paid", 0) or 0),
            working_usd, working_tzs, corporate_usd, corporate_tzs, agent_usd, agent_tzs,
            "ZANZIBAR" if normalized_branch(t) == "UNGUJA" else normalized_branch(t)
        ]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row_no, col, value)
            cell.alignment = Alignment(vertical="center")
            if col >= 6:
                cell.number_format = '#,##0.00;[Red](#,##0.00);-'

    last_detail_row = header_row + len(tickets)
    if not tickets:
        # Make an empty branch explicit instead of presenting a blank register.
        ws.merge_cells(start_row=header_row + 1, start_column=1, end_row=header_row + 1, end_column=len(headers))
        empty_cell = ws.cell(header_row + 1, 1, f"HAKUNA TIKETI ZA {register_scope} KWENYE KIPINDI HIKI.")
        empty_cell.font = Font(bold=True, italic=True, color="9C0006")
        empty_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[header_row + 1].height = 32
        last_detail_row = header_row + 1
    total_row = last_detail_row + 1
    ws.cell(total_row, 1, f"{register_scope} TOTAL")
    for col in range(6, 16):
        letter = get_column_letter(col)
        ws.cell(total_row, col, f"=SUM({letter}{header_row+1}:{letter}{max(last_detail_row, header_row+1)})" if tickets else 0)
    for cell in ws[total_row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAD3")
        cell.number_format = '#,##0.00;[Red](#,##0.00);-'

    summary_row = total_row + 3
    # Two-level headers: category on the first line, USD/TZS directly below.
    # Every category receives two equal-width columns so all sheets keep the
    # same clean, aligned layout without wrapping one currency onto its own line.
    summary_groups = [
        (2, 3, "HAWAII WORKING CLIENT S/FEE"),
        (4, 5, "HAWAII CORPORATE S/FEE"),
        (6, 7, "NON-IATA S/FEE"),
        (8, 9, "TOTAL PROFIT"),
    ]
    summary_keys = [
        "WORKING CLIENT USD", "WORKING CLIENT TZS",
        "CORPORATE USD", "CORPORATE TZS",
        "NON-IATA USD", "NON-IATA TZS",
        "TOTAL PROFIT USD", "TOTAL PROFIT TZS",
    ]

    def write_summary_row(row, label, summary, highlight=False):
        ws.cell(row, 1, label)
        for col, key in enumerate(summary_keys, 2):
            ws.cell(row, col, summary[key])
        for cell in ws[row][:9]:
            cell.font = Font(bold=True, size=11)
            cell.fill = PatternFill("solid", fgColor="C6E0B4" if highlight else "E2F0D9")
            cell.alignment = Alignment(horizontal="center" if cell.column > 1 else "left", vertical="center")
            if cell.column > 1:
                cell.number_format = '#,##0.00;[Red](#,##0.00);-'

    ws.merge_cells(start_row=summary_row, start_column=1, end_row=summary_row + 1, end_column=1)
    title_cell = ws.cell(summary_row, 1, "PROFIT SUMMARY")
    title_cell.font = Font(bold=True, color="FFFFFF", size=11)
    title_cell.fill = PatternFill("solid", fgColor="4472C4")
    title_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for first_col, last_col, heading in summary_groups:
        ws.merge_cells(start_row=summary_row, start_column=first_col, end_row=summary_row, end_column=last_col)
        c = ws.cell(summary_row, first_col, heading)
        c.font = Font(bold=True, color="FFFFFF", size=11)
        c.fill = PatternFill("solid", fgColor="4472C4")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, shrink_to_fit=True)
        for col, currency in ((first_col, "USD"), (last_col, "TZS")):
            sub = ws.cell(summary_row + 1, col, currency)
            sub.font = Font(bold=True, color="FFFFFF", size=11)
            sub.fill = PatternFill("solid", fgColor="4472C4")
            sub.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[summary_row].height = 28
    ws.row_dimensions[summary_row + 1].height = 22
    branch_rows = [r for r in all_branch_tickets if normalized_branch(r) == branch]
    branch_display = "ZANZIBAR" if branch == "UNGUJA" else branch
    write_summary_row(summary_row + 2, f"{branch_display} PROFIT", summarize(branch_rows))
    if branch == "DAR":
        unguja_rows = [r for r in all_branch_tickets if normalized_branch(r) == "UNGUJA"]
        write_summary_row(summary_row + 3, "ZANZIBAR PROFIT", summarize(unguja_rows))
        write_summary_row(summary_row + 4, "TOTAL COMPANY PROFIT (DAR + ZANZIBAR)",
                          summarize(all_branch_tickets), highlight=True)

    from openpyxl.styles import Border, Side
    thin = Side(style="thin", color="B7B7B7")
    final_row = summary_row + (4 if branch == "DAR" else 2)
    for row in ws.iter_rows(min_row=header_row, max_row=final_row, max_col=max(len(headers), 9)):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    widths = [14, 34, 23, 13, 20, 22, 22, 20, 16, 27, 27, 25, 25, 20, 20, 14]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    # Summary uses identical widths for each USD/TZS pair on both branch exports.
    ws.column_dimensions["A"].width = 30
    for col in range(2, 10):
        ws.column_dimensions[get_column_letter(col)].width = 22
    ws.row_dimensions[header_row].height = 48
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
    ws.auto_filter.ref = f"A{header_row}:P{max(last_detail_row, header_row)}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:4"
    ws.print_area = f"A1:P{final_row}"
    ws.sheet_view.zoomScale = 85
    return wb


def _add_all_branch_tickets_sheet(wb, date_from, date_to, tickets):
    """Add a transparent detail listing so company-wide totals can be traced to tickets."""
    ws = wb.create_sheet("ALL BRANCH TICKETS")
    headers = [
        "DATE", "BRANCH", "CLIENT NAME / PASSENGER", "TICKET NUMBER", "AIRLINE", "ROUTE",
        "CURRENCY", "BSP TOTAL AMOUNT", "CUSTOMER TOTAL", "PAID", "BALANCE",
        "CLIENT TYPE", "SERVICE FEE", "COMMISSION", "STATUS"
    ]
    _excel_title(ws, len(headers), "HAWAII TRAVEL AGENCY - ALL BRANCH TICKET DETAILS",
                 f"Issued tickets | Period: {date_from or 'All dates'} to {date_to or 'All dates'}")
    for col, heading in enumerate(headers, 1):
        c = ws.cell(4, col, heading)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="4472C4")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def val(row, key, default=None):
        try:
            return row[key]
        except (KeyError, IndexError, TypeError):
            return default

    for row_no, t in enumerate(tickets, 5):
        currency = str(val(t, "currency", "") or "").upper()
        bsp = float(val(t, "net_fare", 0) or 0) + float(val(t, "taxes", 0) or 0)
        total = float(val(t, "total_selling", 0) or 0)
        paid = float(val(t, "amount_paid", 0) or 0)
        values = [
            val(t, "ticket_date", ""), str(val(t, "branch", "") or "").strip().upper(),
            f"{val(t, 'customer_name', '') or 'Walk-in'} / {val(t, 'passenger_name', '') or ''}",
            val(t, "ticket_number", ""), val(t, "airline", ""), val(t, "route", ""), currency,
            bsp, total, paid, total - paid, val(t, "client_type", ""),
            float(val(t, "service_charge", 0) or 0), float(val(t, "commission", 0) or 0),
            val(t, "status", "")
        ]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row_no, col, value)
            cell.alignment = Alignment(vertical="center")
            if col in (8, 9, 10, 11, 13, 14):
                cell.number_format = '#,##0.00;[Red](#,##0.00);-'
    if not tickets:
        ws.merge_cells("A5:O5")
        ws["A5"] = "Hakuna tiketi zilizotolewa kwenye kipindi kilichochaguliwa."
        ws["A5"].alignment = Alignment(horizontal="center")
    widths = [14, 14, 34, 22, 14, 20, 12, 20, 20, 18, 18, 20, 18, 18, 14]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:O{max(5, 4 + len(tickets))}"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:4"
    ws.print_area = f"A1:O{max(5, 4 + len(tickets))}"


def _export_branch_excel(branch):
    date_from, date_to, where, params = _ticket_export_filters(request)
    # Normalize status in SQL so older records with different capitalization or
    # accidental surrounding spaces are not silently omitted from the export.
    where = [clause for clause in where if clause != "t.status = 'ISSUED'"]
    where.append("TRIM(UPPER(COALESCE(t.status, ''))) = 'ISSUED'")
    where_sql = " AND ".join(where) if where else "1=1"
    conn = get_db()
    try:
        all_tickets = conn.execute(
            "SELECT t.*, c.name AS customer_name FROM tickets t "
            "LEFT JOIN customers c ON t.customer_id=c.id "
            f"WHERE {where_sql} ORDER BY t.ticket_date ASC, t.id ASC", params
        ).fetchall()
    finally:
        conn.close()
    # Business rule: DAR export is the single-sheet consolidated company
    # register (DAR + ZANZIBAR); the ZANZIBAR export contains ZANZIBAR only.
    branch_tickets = [
        t for t in all_tickets
        if str(t["branch"] or "").strip().upper() == branch
    ]
    detail_tickets = all_tickets if branch == "DAR" else branch_tickets
    wb = _make_branch_ticket_excel(branch, date_from, date_to, detail_tickets, all_tickets)
    return _excel_response(wb, f"hawaii_{branch.lower()}_ticket_register_{date.today().isoformat()}.xlsx")


@app.route("/tickets/export/dar/excel")
@login_required
def export_dar_tickets_excel():
    return _export_branch_excel("DAR")


@app.route("/tickets/export/unguja/excel")
@login_required
def export_unguja_tickets_excel():
    return _export_branch_excel("UNGUJA")


# ---------------- 05. BSP RECONCILIATION ----------------
@app.route("/bsp", methods=["GET", "POST"])
@login_required
def bsp_reconciliation():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        if action == "import_csv":
            billing_period = request.form["billing_period"].strip()
            file = request.files.get("bsp_file")
            if file and file.filename:
                stream = io.StringIO(file.stream.read().decode("utf-8-sig"), newline=None)
                reader = csv.DictReader(stream)
                count = 0
                for row in reader:
                    t_num = row.get("ticket_number", "").strip()
                    if not t_num:
                        continue
                    t_date = row.get("tx_date", date.today().strftime("%Y-%m-%d")).strip()
                    airline = row.get("airline", "").strip()
                    tx_type = row.get("tx_type", "SALE").strip().upper()
                    currency = row.get("currency", "USD").strip().upper()
                    bsp_amt = parse_num(row.get("amount"))

                    dup = conn.execute("SELECT COUNT(*) FROM bsp_billing WHERE ticket_number=? AND billing_period=?", (t_num, billing_period)).fetchone()[0]
                    sys_t = conn.execute("SELECT * FROM tickets WHERE ticket_number=?", (t_num,)).fetchone()
                    if dup > 0:
                        m_status = "DUPLICATE"
                    elif not sys_t:
                        m_status = "MISSING_IN_SYSTEM"
                    else:
                        sys_net = sys_t["net_fare"] + sys_t["taxes"]
                        if abs(sys_net - bsp_amt) < 1.0:
                            m_status = "MATCHED"
                            conn.execute("UPDATE tickets SET bsp_reconciled=1 WHERE id=?", (sys_t["id"],))
                        else:
                            m_status = "AMOUNT_MISMATCH"

                    conn.execute(
                        """INSERT INTO bsp_billing (billing_period, tx_date, ticket_number, airline, tx_type, currency, bsp_amount, match_status)
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                        (billing_period, t_date, t_num, airline, tx_type, currency, bsp_amt, m_status)
                    )
                    count += 1
                log_audit(conn, session["user_id"], session["username"], "IMPORT", "BSP", billing_period, f"Imported {count} BSP rows")
                conn.commit()
                flash(f"Umefanikiwa kupakia miamala {count} ya IATA BSP.", "success")
            conn.close()
            return redirect(url_for("bsp_reconciliation"))

        elif action in ("manual_bsp", "edit_bsp"):
            billing_period = request.form["billing_period"].strip()
            t_date = request.form["tx_date"]
            t_num = request.form["ticket_number"].strip()
            airline = request.form.get("airline", "").strip().upper()
            tx_type = request.form["tx_type"]
            currency = request.form["currency"]
            bsp_amt = parse_num(request.form.get("bsp_amount"))

            sys_t = conn.execute("SELECT * FROM tickets WHERE ticket_number=?", (t_num,)).fetchone()
            if not sys_t:
                m_status = "MISSING_IN_SYSTEM"
            else:
                sys_net = sys_t["net_fare"] + sys_t["taxes"]
                if abs(sys_net - bsp_amt) < 1.0:
                    m_status = "MATCHED"
                    conn.execute("UPDATE tickets SET bsp_reconciled=1 WHERE id=?", (sys_t["id"],))
                else:
                    m_status = "AMOUNT_MISMATCH"

            if action == "manual_bsp":
                dup = conn.execute("SELECT COUNT(*) FROM bsp_billing WHERE ticket_number=? AND billing_period=?", (t_num, billing_period)).fetchone()[0]
                if dup > 0:
                    m_status = "DUPLICATE"
                conn.execute(
                    """INSERT INTO bsp_billing (billing_period, tx_date, ticket_number, airline, tx_type, currency, bsp_amount, match_status)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (billing_period, t_date, t_num, airline, tx_type, currency, bsp_amt, m_status)
                )
                log_audit(conn, session["user_id"], session["username"], "CREATE", "BSP", t_num, f"Added BSP item {t_num} ({m_status})")
                flash(f"Umefanikiwa kuhifadhi tiketi {t_num} kwenye BSP ({m_status}).", "success")
            else:
                bsp_id = request.form["bsp_id"]
                conn.execute(
                    """UPDATE bsp_billing SET billing_period=?, tx_date=?, ticket_number=?, airline=?, tx_type=?, currency=?, bsp_amount=?, match_status=?
                       WHERE id=?""",
                    (billing_period, t_date, t_num, airline, tx_type, currency, bsp_amt, m_status, bsp_id)
                )
                re_evaluate_bsp_for_ticket(conn, t_num)
                log_audit(conn, session["user_id"], session["username"], "EDIT", "BSP", bsp_id, f"Edited BSP item #{bsp_id}")
                flash("Umefanikiwa kurekebisha rekodi ya BSP.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("bsp_reconciliation"))

        elif action == "delete_bsp":
            bsp_id = request.form["bsp_id"]
            row = conn.execute("SELECT ticket_number FROM bsp_billing WHERE id=?", (bsp_id,)).fetchone()
            conn.execute("DELETE FROM bsp_billing WHERE id=?", (bsp_id,))
            if row:
                re_evaluate_bsp_for_ticket(conn, row["ticket_number"])
            log_audit(conn, session["user_id"], session["username"], "DELETE", "BSP", bsp_id, f"Deleted BSP item #{bsp_id}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta rekodi ya BSP.", "success")
            return redirect(url_for("bsp_reconciliation"))

        elif action == "pay_bsp":
            pay_date = request.form["payment_date"]
            period = request.form["billing_period"]
            currency = request.form["currency"]
            amount_paid = parse_num(request.form.get("amount_paid"))
            ref = request.form.get("bank_reference", "")
            notes = request.form.get("notes", "")

            cur = conn.execute(
                """INSERT INTO bsp_payments (payment_date, billing_period, currency, amount_paid, bank_reference, notes, created_by)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (pay_date, period, currency, amount_paid, ref, notes, session["username"])
            )
            p_id = cur.lastrowid
            sync_ledger_entry(
                conn, "BSP_PAYMENT", p_id, pay_date, currency, "HQ", "BANK", "OUT",
                "BSP", ref, f"IATA BSP Settlement ({period})", amount_paid, session["username"]
            )
            log_audit(conn, session["user_id"], session["username"], "PAYMENT", "BSP", p_id, f"Paid BSP {currency} {amount_paid}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kuhifadhi malipo ya IATA BSP.", "success")
            return redirect(url_for("bsp_reconciliation"))

        elif action == "delete_bsp_payment":
            p_id = int(request.form["payment_id"])
            conn.execute("DELETE FROM bsp_payments WHERE id=?", (p_id,))
            conn.execute("DELETE FROM bank_transactions WHERE source_module='BSP_PAYMENT' AND source_id=?", (p_id,))
            log_audit(conn, session["user_id"], session["username"], "DELETE", "BSP_PAY", p_id, f"Deleted BSP payment #{p_id}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta malipo ya BSP.", "success")
            return redirect(url_for("bsp_reconciliation"))

    bsp_rows = conn.execute("SELECT * FROM bsp_billing ORDER BY id DESC").fetchall()
    bsp_payments = conn.execute("SELECT * FROM bsp_payments ORDER BY id DESC").fetchall()
    unreconciled_tickets = conn.execute("SELECT * FROM tickets WHERE bsp_reconciled=0 AND status='ISSUED'").fetchall()

    sys_iata_usd = conn.execute("SELECT COALESCE(SUM(net_fare+taxes),0) FROM tickets WHERE currency='USD' AND status='ISSUED'").fetchone()[0]
    paid_bsp_usd = conn.execute("SELECT COALESCE(SUM(amount_paid),0) FROM bsp_payments WHERE currency='USD'").fetchone()[0]
    sys_iata_tzs = conn.execute("SELECT COALESCE(SUM(net_fare+taxes),0) FROM tickets WHERE currency='TZS' AND status='ISSUED'").fetchone()[0]
    paid_bsp_tzs = conn.execute("SELECT COALESCE(SUM(amount_paid),0) FROM bsp_payments WHERE currency='TZS'").fetchone()[0]

    conn.close()
    return render_template(
        "bsp.html", bsp_rows=bsp_rows, bsp_payments=bsp_payments,
        unreconciled_tickets=unreconciled_tickets,
        out_usd=sys_iata_usd - paid_bsp_usd, out_tzs=sys_iata_tzs - paid_bsp_tzs
    )

def apply_customer_payment_allocations(conn, payment_id, customer_id, ticket_id, currency, amount):
    """Allocate a payment to ticket debt, opening debt, then customer credit."""
    rem = float(amount)
    cust = conn.execute("SELECT * FROM customers WHERE id=?", (customer_id,)).fetchone()
    if not cust:
        raise ValueError("Mteja hajapatikana.")

    def allocate_ticket(t):
        nonlocal rem
        due = max(0.0, float(t["total_selling"] or 0) - float(t["amount_paid"] or 0))
        part = min(rem, due)
        if part > 0:
            paid = float(t["amount_paid"] or 0) + part
            status = "PAID" if paid >= float(t["total_selling"] or 0) else "PARTIAL"
            conn.execute("UPDATE tickets SET amount_paid=?, payment_status=? WHERE id=?", (paid, status, t["id"]))
            conn.execute("INSERT INTO customer_payment_allocations (payment_id, allocation_type, ticket_id, amount) VALUES (?, 'TICKET', ?, ?)", (payment_id, t["id"], part))
            rem -= part

    if ticket_id:
        t = conn.execute("SELECT * FROM tickets WHERE id=? AND customer_id=? AND currency=? AND status='ISSUED'", (ticket_id, customer_id, currency)).fetchone()
        if not t:
            raise ValueError("Tiketi uliyochagua haihusiani na mteja au sarafu hii.")
        allocate_ticket(t)
    else:
        unpaid = conn.execute("SELECT * FROM tickets WHERE customer_id=? AND currency=? AND payment_status!='PAID' AND status='ISSUED' ORDER BY ticket_date, id", (customer_id, currency)).fetchall()
        for t in unpaid:
            if rem <= 0: break
            allocate_ticket(t)

    debt_col = "opening_debt_tzs" if currency == "TZS" else "opening_debt_usd"
    credit_col = "credit_balance_tzs" if currency == "TZS" else "credit_balance_usd"
    opening = float(cust[debt_col] or 0)
    part = min(rem, opening)
    if part > 0:
        conn.execute(f"UPDATE customers SET {debt_col}=? WHERE id=?", (opening-part, customer_id))
        conn.execute("INSERT INTO customer_payment_allocations (payment_id, allocation_type, amount) VALUES (?, 'OPENING_DEBT', ?)", (payment_id, part))
        rem -= part
    if rem > 0:
        conn.execute(f"UPDATE customers SET {credit_col}=COALESCE({credit_col},0)+? WHERE id=?", (rem, customer_id))
        conn.execute("INSERT INTO customer_payment_allocations (payment_id, allocation_type, amount) VALUES (?, 'CREDIT', ?)", (payment_id, rem))

def reverse_customer_payment_allocations(conn, payment_id, customer_id, currency):
    """Reverse exact allocations made by a payment; returns False for legacy records."""
    rows = conn.execute("SELECT * FROM customer_payment_allocations WHERE payment_id=? ORDER BY id DESC", (payment_id,)).fetchall()
    if not rows:
        return False
    debt_col = "opening_debt_tzs" if currency == "TZS" else "opening_debt_usd"
    credit_col = "credit_balance_tzs" if currency == "TZS" else "credit_balance_usd"
    for row in rows:
        amount = float(row["amount"] or 0)
        if row["allocation_type"] == "TICKET" and row["ticket_id"]:
            t = conn.execute("SELECT * FROM tickets WHERE id=?", (row["ticket_id"],)).fetchone()
            if t:
                paid = max(0.0, float(t["amount_paid"] or 0)-amount)
                status = "PAID" if paid >= float(t["total_selling"] or 0) and paid > 0 else ("PARTIAL" if paid > 0 else "UNPAID")
                conn.execute("UPDATE tickets SET amount_paid=?, payment_status=? WHERE id=?", (paid, status, t["id"]))
        elif row["allocation_type"] == "OPENING_DEBT":
            conn.execute(f"UPDATE customers SET {debt_col}=COALESCE({debt_col},0)+? WHERE id=?", (amount, customer_id))
        elif row["allocation_type"] == "CREDIT":
            conn.execute(f"UPDATE customers SET {credit_col}=MAX(0, COALESCE({credit_col},0)-?) WHERE id=?", (amount, customer_id))
    conn.execute("DELETE FROM customer_payment_allocations WHERE payment_id=?", (payment_id,))
    return True


# ---------------- 06. CUSTOMER ACCOUNTS ----------------
@app.route("/customers", methods=["GET", "POST"])
@login_required
def customer_accounts():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        if action in ("add_customer", "edit_customer"):
            name = request.form["name"].strip()
            category = request.form["category"]
            phone = request.form.get("phone", "").strip()
            email = request.form.get("email", "").strip()
            branch = request.form["branch"]
            credit_limit = parse_num(request.form.get("credit_limit"))
            opening_debt_tzs = parse_num(request.form.get("opening_debt_tzs"))
            opening_debt_usd = parse_num(request.form.get("opening_debt_usd"))

            if action == "add_customer":
                cur = conn.execute(
                    """INSERT INTO customers (name, category, phone, email, branch, credit_limit, opening_debt_tzs, opening_debt_usd)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (name, category, phone, email, branch, credit_limit, opening_debt_tzs, opening_debt_usd)
                )
                log_audit(conn, session["user_id"], session["username"], "CREATE", "CUSTOMERS", cur.lastrowid, f"Added customer {name} (Open Debt: TZS {opening_debt_tzs}, USD {opening_debt_usd})")
                flash(f"Umefanikiwa kusajili mteja '{name}'.", "success")
            else:
                cid = request.form["customer_id"]
                conn.execute(
                    """UPDATE customers SET name=?, category=?, phone=?, email=?, branch=?, credit_limit=?, opening_debt_tzs=?, opening_debt_usd=?
                       WHERE id=?""",
                    (name, category, phone, email, branch, credit_limit, opening_debt_tzs, opening_debt_usd, cid)
                )
                log_audit(conn, session["user_id"], session["username"], "EDIT", "CUSTOMERS", cid, f"Edited customer {name}")
                flash("Umefanikiwa kurekebisha taarifa za mteja.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("customer_accounts"))

        elif action == "delete_customer":
            cid = request.form["customer_id"]
            conn.execute("DELETE FROM customers WHERE id=?", (cid,))
            log_audit(conn, session["user_id"], session["username"], "DELETE", "CUSTOMERS", cid, f"Deleted customer #{cid}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta mteja.", "success")
            return redirect(url_for("customer_accounts", scroll_y=request.form.get("scroll_y", "0")))

        elif action == "record_payment":
            pay_date = request.form["payment_date"]
            customer_id = int(request.form["customer_id"])
            raw_tid = request.form.get("ticket_id", "").strip()
            ticket_id = int(raw_tid) if raw_tid else None
            branch = request.form["branch"]
            currency = request.form["currency"]
            amount = parse_num(request.form.get("amount"))
            method = request.form["payment_method"]
            ref = request.form.get("reference", "").strip()
            notes = request.form.get("notes", "").strip()

            if amount <= 0:
                flash("Weka kiasi sahihi cha malipo (zaidi ya 0).", "danger")
                conn.close()
                return redirect(url_for("customer_accounts"))

            cust = conn.execute("SELECT * FROM customers WHERE id=?", (customer_id,)).fetchone()
            cust_name = cust["name"] if cust else f"Customer #{customer_id}"

            cur = conn.execute(
                """INSERT INTO customer_payments (payment_date, customer_id, ticket_id, branch, currency, amount, payment_method, reference, notes, created_by)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (pay_date, customer_id, ticket_id, branch, currency, amount, method, ref, notes, session["username"])
            )
            pay_id = cur.lastrowid

            try:
                apply_customer_payment_allocations(conn, pay_id, customer_id, ticket_id, currency, amount)
            except ValueError as exc:
                conn.execute("DELETE FROM customer_payments WHERE id=?", (pay_id,))
                conn.rollback()
                conn.close()
                flash(str(exc), "danger")
                return redirect(url_for("customer_accounts"))

            sync_ledger_entry(
                conn, "CUSTOMER_PAYMENT", pay_id, pay_date, currency, branch, method, "IN",
                "CUSTOMER_DEBT_PAYMENT", ref, f"Malipo ya Deni: {cust_name}", amount, session["username"]
            )

            log_audit(conn, session["user_id"], session["username"], "PAYMENT", "CUSTOMERS", pay_id, f"{cust_name} paid {currency} {amount} via {method}")
            conn.commit()
            conn.close()
            flash(f"Umefanikiwa kuhifadhi malipo ya {cust_name} ({currency} {amount:,.2f}).", "success")
            return redirect(url_for("customer_accounts"))

        elif action == "edit_payment":
            pid = int(request.form["payment_id"])
            pay = conn.execute("SELECT * FROM customer_payments WHERE id=?", (pid,)).fetchone()
            if not pay:
                conn.close(); flash("Malipo hayakupatikana.", "danger")
                return redirect(url_for("customer_accounts"))
            cid = int(pay["customer_id"]); old_currency = pay["currency"]
            # Reverse the original allocation before applying the edited values.
            had_allocations = reverse_customer_payment_allocations(conn, pid, cid, old_currency)
            if not had_allocations:
                # Legacy payment: reverse using the same historical FIFO fallback used by Delete.
                old_amt = float(pay["amount"] or 0); old_tid = pay["ticket_id"]
                if old_tid:
                    old_ticket = conn.execute("SELECT * FROM tickets WHERE id=?", (old_tid,)).fetchone()
                    if old_ticket:
                        paid_now = max(0.0, float(old_ticket["amount_paid"] or 0)-old_amt)
                        st_now = "PAID" if paid_now >= float(old_ticket["total_selling"] or 0) and paid_now > 0 else ("PARTIAL" if paid_now > 0 else "UNPAID")
                        conn.execute("UPDATE tickets SET amount_paid=?, payment_status=? WHERE id=?", (paid_now, st_now, old_tid))
                else:
                    rem_old = old_amt
                    old_paid_tickets = conn.execute("SELECT * FROM tickets WHERE customer_id=? AND currency=? AND amount_paid>0 AND status='ISSUED' ORDER BY ticket_date DESC, id DESC", (cid, old_currency)).fetchall()
                    for old_ticket in old_paid_tickets:
                        if rem_old <= 0: break
                        reversal = min(rem_old, float(old_ticket["amount_paid"] or 0))
                        paid_now = float(old_ticket["amount_paid"] or 0)-reversal
                        st_now = "PAID" if paid_now >= float(old_ticket["total_selling"] or 0) and paid_now > 0 else ("PARTIAL" if paid_now > 0 else "UNPAID")
                        conn.execute("UPDATE tickets SET amount_paid=?, payment_status=? WHERE id=?", (paid_now, st_now, old_ticket["id"]))
                        rem_old -= reversal
                    if rem_old > 0:
                        old_debt_col = "opening_debt_tzs" if old_currency == "TZS" else "opening_debt_usd"
                        conn.execute(f"UPDATE customers SET {old_debt_col}=COALESCE({old_debt_col},0)+? WHERE id=?", (rem_old, cid))
            pay_date = request.form["payment_date"]
            currency = request.form["currency"]
            amount = parse_num(request.form.get("amount"))
            branch = request.form["branch"]
            method = request.form["payment_method"]
            ref = request.form.get("reference", "").strip()
            notes = request.form.get("notes", "").strip()
            ticket_raw = request.form.get("ticket_id", "").strip()
            ticket_id = int(ticket_raw) if ticket_raw else None
            if amount <= 0:
                conn.rollback(); conn.close(); flash("Kiasi lazima kiwe zaidi ya sifuri.", "danger")
                return redirect(url_for("customer_accounts"))
            conn.execute("UPDATE customer_payments SET payment_date=?, ticket_id=?, branch=?, currency=?, amount=?, payment_method=?, reference=?, notes=? WHERE id=?", (pay_date, ticket_id, branch, currency, amount, method, ref, notes, pid))
            try:
                apply_customer_payment_allocations(conn, pid, cid, ticket_id, currency, amount)
            except ValueError as exc:
                conn.rollback(); conn.close(); flash(str(exc), "danger")
                return redirect(url_for("customer_accounts"))
            sync_ledger_entry(conn, "CUSTOMER_PAYMENT", pid, pay_date, currency, branch, method, "IN", "CUSTOMER_DEBT_PAYMENT", ref, "Malipo ya Deni (Yamehaririwa)", amount, session["username"])
            log_audit(conn, session["user_id"], session["username"], "EDIT", "CUST_PAY", pid, f"Edited customer payment #{pid}: {currency} {amount}")
            conn.commit(); conn.close(); flash("Malipo yamehaririwa na salio limekokotolewa upya.", "success")
            return redirect(url_for("customer_accounts", scroll_y=request.form.get("scroll_y", "0")))

        elif action == "delete_payment":
            pid = int(request.form["payment_id"])
            pay = conn.execute("SELECT * FROM customer_payments WHERE id=?", (pid,)).fetchone()
            if pay:
                cid = pay["customer_id"]
                curr = pay["currency"]
                amt = float(pay["amount"])
                tid = pay["ticket_id"]

                had_allocations = reverse_customer_payment_allocations(conn, pid, cid, curr)
                if not had_allocations:
                    # Backward-compatible reversal for payments recorded before allocation tracking.
                    if tid:
                        t = conn.execute("SELECT * FROM tickets WHERE id=?", (tid,)).fetchone()
                        if t:
                            new_paid = max(0.0, float(t["amount_paid"] or 0) - amt)
                            st = "PAID" if new_paid >= float(t["total_selling"] or 0) and new_paid > 0 else ("PARTIAL" if new_paid > 0 else "UNPAID")
                            conn.execute("UPDATE tickets SET amount_paid=?, payment_status=? WHERE id=?", (new_paid, st, tid))
                    else:
                        rem = amt
                        paid_ts = conn.execute("SELECT * FROM tickets WHERE customer_id=? AND currency=? AND amount_paid>0 AND status='ISSUED' ORDER BY ticket_date DESC, id DESC", (cid, curr)).fetchall()
                        for t in paid_ts:
                            if rem <= 0: break
                            rev = min(rem, float(t["amount_paid"] or 0))
                            new_paid = float(t["amount_paid"] or 0) - rev
                            st = "PAID" if new_paid >= float(t["total_selling"] or 0) and new_paid > 0 else ("PARTIAL" if new_paid > 0 else "UNPAID")
                            conn.execute("UPDATE tickets SET amount_paid=?, payment_status=? WHERE id=?", (new_paid, st, t["id"]))
                            rem -= rev
                        if rem > 0:
                            debt_col = "opening_debt_tzs" if curr == "TZS" else "opening_debt_usd"
                            conn.execute(f"UPDATE customers SET {debt_col}=COALESCE({debt_col},0)+? WHERE id=?", (rem, cid))

                conn.execute("DELETE FROM customer_payments WHERE id=?", (pid,))
                conn.execute("DELETE FROM bank_transactions WHERE source_module='CUSTOMER_PAYMENT' AND source_id=?", (pid,))
                conn.execute("DELETE FROM cash_transactions WHERE source_module='CUSTOMER_PAYMENT' AND source_id=?", (pid,))
                log_audit(conn, session["user_id"], session["username"], "DELETE", "CUST_PAY", pid, f"Deleted customer payment #{pid} and restored debt")
                conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta malipo ya mteja.", "success")
            return redirect(url_for("customer_accounts", scroll_y=request.form.get("scroll_y", "0")))

    customers = conn.execute(
        """SELECT c.*,
           (COALESCE((SELECT SUM(total_selling - amount_paid) FROM tickets t WHERE t.customer_id=c.id AND t.currency='TZS' AND t.status='ISSUED'), 0) + COALESCE(c.opening_debt_tzs, 0)) AS debt_tzs,
           (COALESCE((SELECT SUM(total_selling - amount_paid) FROM tickets t WHERE t.customer_id=c.id AND t.currency='USD' AND t.status='ISSUED'), 0) + COALESCE(c.opening_debt_usd, 0)) AS debt_usd,
           COALESCE(c.credit_balance_tzs,0) AS credit_tzs, COALESCE(c.credit_balance_usd,0) AS credit_usd
           FROM customers c ORDER BY c.id DESC"""
    ).fetchall()
    unpaid_tickets = conn.execute(
        """SELECT t.*, c.name AS customer_name FROM tickets t
           LEFT JOIN customers c ON t.customer_id = c.id
           WHERE t.payment_status != 'PAID' AND t.status='ISSUED'
           ORDER BY t.ticket_date DESC"""
    ).fetchall()
    payments = conn.execute(
        """SELECT p.*, c.name AS customer_name, t.ticket_number
           FROM customer_payments p
           JOIN customers c ON p.customer_id = c.id
           LEFT JOIN tickets t ON p.ticket_id = t.id
           ORDER BY p.payment_date DESC, p.id DESC LIMIT 50"""
    ).fetchall()
    all_customer_tickets = conn.execute(
        """SELECT id, customer_id, ticket_number, passenger_name, currency
           FROM tickets WHERE status='ISSUED' AND customer_id IS NOT NULL
           ORDER BY ticket_date DESC, id DESC"""
    ).fetchall()
    conn.close()
    return render_template("customers.html", customers=customers, unpaid_tickets=unpaid_tickets, payments=payments, all_customer_tickets=all_customer_tickets)

@app.route("/customers/<int:cid>/statement")
@login_required
def customer_statement(cid):
    conn = get_db()
    customer = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    if not customer:
        conn.close()
        flash("Mteja hajapatikana.", "danger")
        return redirect(url_for("customer_accounts"))
    tickets = conn.execute(
        """SELECT ticket_date, ticket_number, airline, route, currency, total_selling, amount_paid,
                  (total_selling - amount_paid) AS balance, payment_status
           FROM tickets WHERE customer_id=? AND status='ISSUED'
           ORDER BY ticket_date, id""", (cid,)
    ).fetchall()
    payments = conn.execute(
        """SELECT p.payment_date, p.currency, p.amount, p.payment_method, p.reference,
                  t.ticket_number
           FROM customer_payments p LEFT JOIN tickets t ON p.ticket_id=t.id
           WHERE p.customer_id=? ORDER BY p.payment_date, p.id""", (cid,)
    ).fetchall()
    debt = get_customer_debt_summary(conn, cid)
    conn.close()
    return render_template("customer_statement.html", customer=customer, tickets=tickets,
                           payments=payments, debt=debt, generated_on=date.today().isoformat())

# ---------------- 07. DIRECTOR & TRANSFER CONTROL ----------------
@app.route("/director", methods=["GET", "POST"])
@login_required
def director_control():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        if action in ("add_transfer", "edit_transfer"):
            tx_date = request.form["tx_date"]
            direction = request.form["direction"]
            category = request.form["transfer_category"]
            branch = request.form["branch"]
            account_source = request.form["account_source"]
            currency = request.form["currency"]
            amount = parse_num(request.form.get("amount"))
            reference = request.form.get("reference", "").strip()
            description = request.form["description"].strip()
            status = "UNEXPLAINED" if category == "UNEXPLAINED" else "VERIFIED"

            ledger_dir = "IN" if direction == "IN_TO_COMPANY" else ("OUT" if direction == "OUT_FROM_COMPANY" else "NONE")
            desc_prefix = f"[DIRECTOR IN - {category}]" if direction == "IN_TO_COMPANY" else f"[DIRECTOR OUT - {category}]"

            if action == "add_transfer":
                cur = conn.execute(
                    """INSERT INTO director_transfers
                       (tx_date, direction, transfer_category, branch, account_source, currency, amount, reference, description, status, created_by)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (tx_date, direction, category, branch, account_source, currency, amount, reference, description, status, session["username"])
                )
                tr_id = cur.lastrowid
                if ledger_dir in ("IN", "OUT"):
                    sync_ledger_entry(
                        conn, "DIRECTOR", tr_id, tx_date, currency, branch, account_source, ledger_dir,
                        category, reference, f"{desc_prefix} {description}", amount, session["username"]
                    )
                log_audit(conn, session["user_id"], session["username"], "CREATE", "DIRECTOR", tr_id, f"{direction} {category} {currency} {amount}")
                flash("Umefanikiwa kuhifadhi muamala wa Director.", "success")
            else:
                tr_id = int(request.form["tr_id"])
                conn.execute(
                    """UPDATE director_transfers
                       SET tx_date=?, direction=?, transfer_category=?, branch=?, account_source=?, currency=?, amount=?, reference=?, description=?, status=?
                       WHERE id=?""",
                    (tx_date, direction, category, branch, account_source, currency, amount, reference, description, status, tr_id)
                )
                if ledger_dir in ("IN", "OUT"):
                    sync_ledger_entry(
                        conn, "DIRECTOR", tr_id, tx_date, currency, branch, account_source, ledger_dir,
                        category, reference, f"{desc_prefix} {description}", amount, session["username"]
                    )
                else:
                    conn.execute("DELETE FROM bank_transactions WHERE source_module='DIRECTOR' AND source_id=?", (tr_id,))
                    conn.execute("DELETE FROM cash_transactions WHERE source_module='DIRECTOR' AND source_id=?", (tr_id,))
                log_audit(conn, session["user_id"], session["username"], "EDIT", "DIRECTOR", tr_id, f"Edited director transfer #{tr_id}")
                flash("Umefanikiwa kurekebisha muamala wa Director.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("director_control"))

        elif action == "delete_transfer":
            tr_id = int(request.form["tr_id"])
            conn.execute("DELETE FROM director_transfers WHERE id=?", (tr_id,))
            conn.execute("DELETE FROM bank_transactions WHERE source_module='DIRECTOR' AND source_id=?", (tr_id,))
            conn.execute("DELETE FROM cash_transactions WHERE source_module='DIRECTOR' AND source_id=?", (tr_id,))
            log_audit(conn, session["user_id"], session["username"], "DELETE", "DIRECTOR", tr_id, f"Deleted director transfer #{tr_id}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta muamala wa Director.", "success")
            return redirect(url_for("director_control"))

        elif action == "verify_transfer":
            tr_id = request.form["tr_id"]
            new_cat = request.form["new_category"]
            conn.execute("UPDATE director_transfers SET transfer_category=?, status='VERIFIED' WHERE id=?", (new_cat, tr_id))
            log_audit(conn, session["user_id"], session["username"], "VERIFY", "DIRECTOR", tr_id, f"Classified unexplained transfer as {new_cat}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kuthibitisha muamala wa Director.", "success")
            return redirect(url_for("director_control"))

    transfers = conn.execute("SELECT * FROM director_transfers ORDER BY tx_date DESC, id DESC").fetchall()
    in_tzs = conn.execute("SELECT COALESCE(SUM(amount),0) FROM director_transfers WHERE currency='TZS' AND direction='IN_TO_COMPANY'").fetchone()[0]
    out_tzs = conn.execute("SELECT COALESCE(SUM(amount),0) FROM director_transfers WHERE currency='TZS' AND direction='OUT_FROM_COMPANY'").fetchone()[0]
    in_usd = conn.execute("SELECT COALESCE(SUM(amount),0) FROM director_transfers WHERE currency='USD' AND direction='IN_TO_COMPANY'").fetchone()[0]
    out_usd = conn.execute("SELECT COALESCE(SUM(amount),0) FROM director_transfers WHERE currency='USD' AND direction='OUT_FROM_COMPANY'").fetchone()[0]
    conn.close()

    return render_template(
        "director.html", transfers=transfers,
        in_tzs=in_tzs, out_tzs=out_tzs, net_tzs=in_tzs - out_tzs,
        in_usd=in_usd, out_usd=out_usd, net_usd=in_usd - out_usd
    )

# ---------------- 08. BRANCH CONTROL ----------------
@app.route("/branches")
@login_required
def branch_control():
    conn = get_db()
    branches = ["DAR", "UNGUJA"]
    stats = {}
    for b in branches:
        book = "DAR_CASH" if b == "DAR" else "UNGUJA_CASH"
        stats[b] = {
            "cash_tzs": conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE book_type=? AND currency='TZS'", (book,)).fetchone()[0],
            "cash_usd": conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE book_type=? AND currency='USD'", (book,)).fetchone()[0],
            "sales_tzs": conn.execute("SELECT COALESCE(SUM(total_selling),0) FROM tickets WHERE branch=? AND currency='TZS' AND status='ISSUED'", (b,)).fetchone()[0],
            "sales_usd": conn.execute("SELECT COALESCE(SUM(total_selling),0) FROM tickets WHERE branch=? AND currency='USD' AND status='ISSUED'", (b,)).fetchone()[0],
            "profit_tzs": conn.execute("SELECT COALESCE(SUM(service_charge+commission),0) FROM tickets WHERE branch=? AND currency='TZS' AND status='ISSUED'", (b,)).fetchone()[0],
            "profit_usd": conn.execute("SELECT COALESCE(SUM(service_charge+commission),0) FROM tickets WHERE branch=? AND currency='USD' AND status='ISSUED'", (b,)).fetchone()[0],
            "exp_tzs": conn.execute("SELECT COALESCE(SUM(amount),0) FROM expenses WHERE branch=? AND currency='TZS' AND status='APPROVED'", (b,)).fetchone()[0],
            "exp_usd": conn.execute("SELECT COALESCE(SUM(amount),0) FROM expenses WHERE branch=? AND currency='USD' AND status='APPROVED'", (b,)).fetchone()[0],
            "tickets_count": conn.execute("SELECT COUNT(*) FROM tickets WHERE branch=? AND status='ISSUED'", (b,)).fetchone()[0],
        }
    branch_tickets = conn.execute("SELECT * FROM tickets ORDER BY ticket_date DESC LIMIT 20").fetchall()
    conn.close()
    return render_template("branches.html", stats=stats, branch_tickets=branch_tickets)

# ---------------- 09. EXPENSES MANAGEMENT ----------------
@app.route("/expenses", methods=["GET", "POST"])
@login_required
def expenses_management():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        if action in ("add_expense", "edit_expense"):
            exp_date = request.form["expense_date"]
            branch = request.form["branch"]
            category = request.form["category"]
            description = request.form["description"].strip()
            currency = request.form["currency"]
            amount = parse_num(request.form.get("amount"))
            source = request.form["payment_source"]

            if action == "add_expense":
                cur = conn.execute(
                    """INSERT INTO expenses (expense_date, branch, category, description, currency, amount, payment_source, status, created_by)
                       VALUES (?, ?, ?, ?, ?, ?, ?, 'PENDING', ?)""",
                    (exp_date, branch, category, description, currency, amount, source, session["username"])
                )
                log_audit(conn, session["user_id"], session["username"], "CREATE", "EXPENSES", cur.lastrowid, f"Submitted expense {currency} {amount}: {description}")
                flash("Umefanikiwa kuhifadhi matumizi.", "success")
            else:
                exp_id = int(request.form["expense_id"])
                exp = conn.execute("SELECT * FROM expenses WHERE id=?", (exp_id,)).fetchone()
                conn.execute(
                    """UPDATE expenses SET expense_date=?, branch=?, category=?, description=?, currency=?, amount=?, payment_source=?
                       WHERE id=?""",
                    (exp_date, branch, category, description, currency, amount, source, exp_id)
                )
                if exp and exp["status"] == "APPROVED":
                    sync_ledger_entry(
                        conn, "EXPENSE", exp_id, exp_date, currency, branch, source, "OUT",
                        "EXPENSE", f"EXP-{exp_id}", description, amount, session["username"]
                    )
                log_audit(conn, session["user_id"], session["username"], "EDIT", "EXPENSES", exp_id, f"Edited expense #{exp_id}")
                flash("Umefanikiwa kurekebisha matumizi.", "success")
            conn.commit()
            conn.close()
            return redirect(url_for("expenses_management"))

        elif action == "delete_expense":
            exp_id = int(request.form["expense_id"])
            conn.execute("DELETE FROM expenses WHERE id=?", (exp_id,))
            conn.execute("DELETE FROM bank_transactions WHERE source_module='EXPENSE' AND source_id=?", (exp_id,))
            conn.execute("DELETE FROM cash_transactions WHERE source_module='EXPENSE' AND source_id=?", (exp_id,))
            log_audit(conn, session["user_id"], session["username"], "DELETE", "EXPENSES", exp_id, f"Deleted expense #{exp_id}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kufuta matumizi.", "success")
            return redirect(url_for("expenses_management"))

        elif action in ("APPROVE", "REJECT"):
            if session.get("role") not in ("admin", "accountant", "boss"):
                flash("Accountant, Boss, au Admin pekee ndio wanaoweza kuidhinisha matumizi.", "danger")
                conn.close()
                return redirect(url_for("expenses_management"))
            exp_id = int(request.form["expense_id"])
            exp = conn.execute("SELECT * FROM expenses WHERE id=?", (exp_id,)).fetchone()
            if exp and exp["status"] == "PENDING":
                new_st = "APPROVED" if action == "APPROVE" else "REJECTED"
                conn.execute("UPDATE expenses SET status=?, approved_by=? WHERE id=?", (new_st, session["username"], exp_id))
                if new_st == "APPROVED":
                    sync_ledger_entry(
                        conn, "EXPENSE", exp_id, exp["expense_date"], exp["currency"], exp["branch"],
                        exp["payment_source"], "OUT", "EXPENSE", f"EXP-{exp_id}", exp["description"], exp["amount"], session["username"]
                    )
                log_audit(conn, session["user_id"], session["username"], new_st, "EXPENSES", exp_id, f"Expense #{exp_id} {new_st}")
                conn.commit()
                flash(f"Umefanikiwa kubadili hali ya matumizi #{exp_id} kuwa {new_st}.", "success")
            conn.close()
            return redirect(url_for("expenses_management"))

    expenses = conn.execute("SELECT * FROM expenses ORDER BY expense_date DESC, id DESC").fetchall()
    conn.close()
    return render_template("expenses.html", expenses=expenses)

# ---------------- 10. FINANCIAL REPORTS ----------------
@app.route("/reports")
@login_required
def financial_reports():
    conn = get_db()
    start_date = request.args.get("start_date", date.today().replace(day=1).strftime("%Y-%m-%d"))
    end_date = request.args.get("end_date", date.today().strftime("%Y-%m-%d"))

    rev_tzs = conn.execute(
        "SELECT COALESCE(SUM(service_charge + commission), 0) FROM tickets WHERE status='ISSUED' AND currency='TZS' AND ticket_date BETWEEN ? AND ?",
        (start_date, end_date)
    ).fetchone()[0]
    rev_usd = conn.execute(
        "SELECT COALESCE(SUM(service_charge + commission), 0) FROM tickets WHERE status='ISSUED' AND currency='USD' AND ticket_date BETWEEN ? AND ?",
        (start_date, end_date)
    ).fetchone()[0]

    exp_tzs = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM expenses WHERE status='APPROVED' AND currency='TZS' AND expense_date BETWEEN ? AND ?",
        (start_date, end_date)
    ).fetchone()[0]
    exp_usd = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) FROM expenses WHERE status='APPROVED' AND currency='USD' AND expense_date BETWEEN ? AND ?",
        (start_date, end_date)
    ).fetchone()[0]

    bk_chg_tzs = conn.execute(
        "SELECT COALESCE(SUM(amount_out), 0) FROM bank_transactions WHERE tx_type='BANK_CHARGE' AND currency='TZS' AND tx_date BETWEEN ? AND ?",
        (start_date, end_date)
    ).fetchone()[0]
    bk_chg_usd = conn.execute(
        "SELECT COALESCE(SUM(amount_out), 0) FROM bank_transactions WHERE tx_type='BANK_CHARGE' AND currency='USD' AND tx_date BETWEEN ? AND ?",
        (start_date, end_date)
    ).fetchone()[0]

    net_profit_tzs = rev_tzs - exp_tzs - bk_chg_tzs
    net_profit_usd = rev_usd - exp_usd - bk_chg_usd

    bank_tzs = conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM bank_transactions WHERE currency='TZS' AND tx_date <= ?", (end_date,)).fetchone()[0]
    bank_usd = conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM bank_transactions WHERE currency='USD' AND tx_date <= ?", (end_date,)).fetchone()[0]
    cash_tzs = conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE currency='TZS' AND tx_date <= ?", (end_date,)).fetchone()[0]
    cash_usd = conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE currency='USD' AND tx_date <= ?", (end_date,)).fetchone()[0]
    ar_tzs = (
        conn.execute("SELECT COALESCE(SUM(total_selling-amount_paid),0) FROM tickets WHERE currency='TZS' AND status='ISSUED' AND ticket_date <= ?", (end_date,)).fetchone()[0]
        + conn.execute("SELECT COALESCE(SUM(opening_debt_tzs),0) FROM customers").fetchone()[0]
    )
    ar_usd = (
        conn.execute("SELECT COALESCE(SUM(total_selling-amount_paid),0) FROM tickets WHERE currency='USD' AND status='ISSUED' AND ticket_date <= ?", (end_date,)).fetchone()[0]
        + conn.execute("SELECT COALESCE(SUM(opening_debt_usd),0) FROM customers").fetchone()[0]
    )

    bsp_due_tzs = (
        conn.execute("SELECT COALESCE(SUM(net_fare+taxes),0) FROM tickets WHERE currency='TZS' AND status='ISSUED' AND ticket_date <= ?", (end_date,)).fetchone()[0]
        - conn.execute("SELECT COALESCE(SUM(amount_paid),0) FROM bsp_payments WHERE currency='TZS' AND payment_date <= ?", (end_date,)).fetchone()[0]
    )
    bsp_due_usd = (
        conn.execute("SELECT COALESCE(SUM(net_fare+taxes),0) FROM tickets WHERE currency='USD' AND status='ISSUED' AND ticket_date <= ?", (end_date,)).fetchone()[0]
        - conn.execute("SELECT COALESCE(SUM(amount_paid),0) FROM bsp_payments WHERE currency='USD' AND payment_date <= ?", (end_date,)).fetchone()[0]
    )

    dir_tzs = (
        conn.execute("SELECT COALESCE(SUM(amount),0) FROM director_transfers WHERE currency='TZS' AND direction='IN_TO_COMPANY' AND tx_date <= ?", (end_date,)).fetchone()[0]
        - conn.execute("SELECT COALESCE(SUM(amount),0) FROM director_transfers WHERE currency='TZS' AND direction='OUT_FROM_COMPANY' AND tx_date <= ?", (end_date,)).fetchone()[0]
    )
    dir_usd = (
        conn.execute("SELECT COALESCE(SUM(amount),0) FROM director_transfers WHERE currency='USD' AND direction='IN_TO_COMPANY' AND tx_date <= ?", (end_date,)).fetchone()[0]
        - conn.execute("SELECT COALESCE(SUM(amount),0) FROM director_transfers WHERE currency='USD' AND direction='OUT_FROM_COMPANY' AND tx_date <= ?", (end_date,)).fetchone()[0]
    )

    conn.close()
    return render_template(
        "reports.html", start_date=start_date, end_date=end_date,
        rev_tzs=rev_tzs, rev_usd=rev_usd, exp_tzs=exp_tzs, exp_usd=exp_usd,
        bk_chg_tzs=bk_chg_tzs, bk_chg_usd=bk_chg_usd,
        net_profit_tzs=net_profit_tzs, net_profit_usd=net_profit_usd,
        bank_tzs=bank_tzs, bank_usd=bank_usd, cash_tzs=cash_tzs, cash_usd=cash_usd,
        ar_tzs=ar_tzs, ar_usd=ar_usd, bsp_due_tzs=bsp_due_tzs, bsp_due_usd=bsp_due_usd,
        dir_tzs=dir_tzs, dir_usd=dir_usd
    )

@app.route("/reports/export/<report_type>")
@login_required
def export_csv(report_type):
    conn = get_db()
    si = io.StringIO()
    cw = csv.writer(si)
    if report_type == "tickets":
        rows = conn.execute("SELECT * FROM tickets ORDER BY ticket_date DESC").fetchall()
        cw.writerow(["ID", "Date", "Ticket No", "Passenger", "Airline", "Client Type", "Branch", "Currency", "Net Fare", "Taxes", "Service Charge", "Commission", "Total Selling", "Paid", "Status", "BSP Reconciled"])
        for r in rows:
            cw.writerow([r["id"], r["ticket_date"], r["ticket_number"], r["passenger_name"], r["airline"], r["client_type"], r["branch"], r["currency"], r["net_fare"], r["taxes"], r["service_charge"], r["commission"], r["total_selling"], r["amount_paid"], r["status"], r["bsp_reconciled"]])
    elif report_type == "bank":
        rows = conn.execute("SELECT * FROM bank_transactions ORDER BY tx_date DESC").fetchall()
        cw.writerow(["ID", "Date", "Currency", "Branch", "Type", "Category", "Reference", "Description", "In", "Out", "Reconciled", "Created By"])
        for r in rows:
            cw.writerow([r["id"], r["tx_date"], r["currency"], r["branch"], r["tx_type"], r["category"], r["reference"], r["description"], r["amount_in"], r["amount_out"], r["is_reconciled"], r["created_by"]])
    elif report_type == "director":
        rows = conn.execute("SELECT * FROM director_transfers ORDER BY tx_date DESC").fetchall()
        cw.writerow(["ID", "Date", "Direction", "Category", "Branch", "Source", "Currency", "Amount", "Reference", "Description", "Status", "Created By"])
        for r in rows:
            cw.writerow([r["id"], r["tx_date"], r["direction"], r["transfer_category"], r["branch"], r["account_source"], r["currency"], r["amount"], r["reference"], r["description"], r["status"], r["created_by"]])
    conn.close()
    return Response(
        si.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment;filename=hawaii_{report_type}_{date.today()}.csv"}
    )

# ---------------- 11. USER MANAGEMENT & AUDIT TRAIL ----------------
@app.route("/users", methods=["GET", "POST"])
@login_required
def user_management():
    conn = get_db()
    if request.method == "POST":
        action = request.form.get("action")
        if action in ("create_user", "edit_user"):
            if session.get("role") != "admin":
                flash("Admin pekee ndiye anayeweza kusimamia watumiaji.", "danger")
                conn.close()
                return redirect(url_for("user_management"))
            username = request.form["username"].strip()
            full_name = request.form["full_name"].strip()
            email = request.form.get("email", "").strip().lower()
            role = request.form["role"]
            if not email or "@" not in email:
                conn.close(); flash("Weka email sahihi kwa ajili ya kurejesha password.", "danger")
                return redirect(url_for("user_management"))
            branch = request.form["branch"]
            if action == "create_user":
                password = request.form["password"]
                if not valid_password(password):
                    conn.close(); flash("Password iwe na angalau herufi 9.", "danger")
                    return redirect(url_for("user_management"))
                try:
                    cur = conn.execute(
                        "INSERT INTO users (username, full_name, email, role, branch, password_hash) VALUES (?, ?, ?, ?, ?, ?)",
                        (username, full_name, email, role, branch, generate_password_hash(password))
                    )
                    log_audit(conn, session["user_id"], session["username"], "CREATE", "USERS", cur.lastrowid, f"Created user {username} ({role})")
                    conn.commit()
                    flash(f"Umefanikiwa kusajili mtumiaji {username}.", "success")
                except Exception:
                    flash("Username au email tayari imetumika.", "danger")
            else:
                uid = request.form["user_id"]
                conn.execute(
                    "UPDATE users SET username=?, full_name=?, email=?, role=?, branch=? WHERE id=?",
                    (username, full_name, email, role, branch, uid)
                )
                log_audit(conn, session["user_id"], session["username"], "EDIT", "USERS", uid, f"Edited user #{uid} ({username})")
                conn.commit()
                flash("Umefanikiwa kurekebisha taarifa za mtumiaji.", "success")
            conn.close()
            return redirect(url_for("user_management"))

        elif action == "delete_user":
            if session.get("role") != "admin":
                flash("Admin pekee ndiye anayeweza kufuta watumiaji.", "danger")
                conn.close()
                return redirect(url_for("user_management"))
            uid = request.form["user_id"]
            if str(uid) == str(session["user_id"]):
                flash("Huwezi kufuta akaunti yako mwenyewe ukiwa umeingia.", "danger")
            else:
                conn.execute("DELETE FROM users WHERE id=?", (uid,))
                log_audit(conn, session["user_id"], session["username"], "DELETE", "USERS", uid, f"Deleted user #{uid}")
                conn.commit()
                flash("Umefanikiwa kufuta mtumiaji.", "success")
            conn.close()
            return redirect(url_for("user_management"))

        elif action == "reset_password":
            if session.get("role") != "admin":
                flash("Admin pekee ndiye anayeweza kubadili nywila za wengine.", "danger")
                conn.close()
                return redirect(url_for("user_management"))
            uid = request.form["user_id"]
            new_pw = request.form["new_password"]
            if not valid_password(new_pw):
                conn.close(); flash("Password iwe na angalau herufi 9.", "danger")
                return redirect(url_for("user_management"))
            conn.execute("UPDATE users SET password_hash=? WHERE id=?", (generate_password_hash(new_pw), uid))
            conn.execute("UPDATE password_reset_tokens SET used_at=? WHERE user_id=? AND used_at IS NULL", (datetime.utcnow().isoformat(), uid))
            log_audit(conn, session["user_id"], session["username"], "PASSWORD_RESET", "USERS", uid, f"Admin reset password for user #{uid}")
            conn.commit()
            conn.close()
            flash("Umefanikiwa kubadili nywila (Password).", "success")
            return redirect(url_for("user_management"))

    users = conn.execute("SELECT id, username, full_name, email, role, branch, is_active, created_at FROM users ORDER BY id").fetchall()
    audits = conn.execute("SELECT * FROM audit_trail ORDER BY id DESC LIMIT 100").fetchall()
    conn.close()
    return render_template("users.html", users=users, audits=audits)


# ---------------- 12. DAILY CLOSING & STAFF CONTROL ----------------
@app.route("/control", methods=["GET", "POST"])
@login_required
def control_center():
    conn = get_db()
    privileged = session.get("role") in ("admin", "accountant", "boss")
    if request.method == "POST":
        action = request.form.get("action")
        if action == "submit_closing":
            branch = request.form.get("branch", session.get("branch"))
            if session.get("role") not in ("admin", "accountant", "boss") and branch != session.get("branch"):
                conn.close(); flash("Huruhusiwi kufunga hesabu ya tawi jingine.", "danger")
                return redirect(url_for("control_center"))
            day = request.form.get("closing_date")
            currency = request.form.get("currency", "TZS")
            counted = parse_num(request.form.get("counted_cash"))
            opening = parse_num(request.form.get("opening_cash"))
            notes = request.form.get("notes", "").strip()
            if counted < 0 or opening < 0 or currency not in ("TZS", "USD") or branch not in ("DAR", "UNGUJA"):
                conn.close(); flash("Taarifa za closing si sahihi.", "danger")
                return redirect(url_for("control_center"))
            book = "DAR_CASH" if branch == "DAR" else "UNGUJA_CASH"
            movement = conn.execute("SELECT COALESCE(SUM(amount_in-amount_out),0) FROM cash_transactions WHERE book_type=? AND currency=? AND tx_date=?", (book,currency,day)).fetchone()[0]
            expected = float(opening) + float(movement)
            variance = counted - expected
            conn.execute("""INSERT INTO daily_closings
                (closing_date,branch,currency,opening_cash,expected_cash,counted_cash,variance,notes,status,submitted_by)
                VALUES (?,?,?,?,?,?,?,?, 'SUBMITTED',?)
                ON CONFLICT(closing_date,branch,currency) DO UPDATE SET
                opening_cash=excluded.opening_cash, expected_cash=excluded.expected_cash,
                counted_cash=excluded.counted_cash, variance=excluded.variance,
                notes=excluded.notes,status='SUBMITTED',submitted_by=excluded.submitted_by,
                reviewed_by=NULL,reviewed_at=NULL""",
                (day,branch,currency,opening,expected,counted,variance,notes,session["username"]))
            log_audit(conn,session["user_id"],session["username"],"SUBMIT","DAILY_CLOSING",f"{day}-{branch}-{currency}",f"Expected {expected}; counted {counted}; variance {variance}")
            if abs(variance) > 0.01:
                conn.execute("INSERT INTO system_alerts(alert_type,message,severity) VALUES('CASH_VARIANCE',?,?)", (f"{branch} {currency} closing {day}: tofauti {variance:,.2f}","HIGH"))
            conn.commit(); conn.close(); flash("Daily closing imehifadhiwa.","success")
            return redirect(url_for("control_center"))
        if action == "review_closing" and privileged:
            cid = request.form.get("closing_id")
            conn.execute("UPDATE daily_closings SET status='REVIEWED',reviewed_by=?,reviewed_at=CURRENT_TIMESTAMP WHERE id=?",(session["username"],cid))
            log_audit(conn,session["user_id"],session["username"],"REVIEW","DAILY_CLOSING",cid,"Reviewed daily closing")
            conn.commit(); conn.close(); flash("Closing imehakikiwa.","success"); return redirect(url_for("control_center"))
        if action == "resolve_alert" and privileged:
            aid=request.form.get("alert_id")
            conn.execute("UPDATE system_alerts SET is_resolved=1,resolved_by=?,resolved_at=CURRENT_TIMESTAMP WHERE id=?",(session["username"],aid))
            log_audit(conn,session["user_id"],session["username"],"RESOLVE","ALERT",aid,"Resolved system alert")
            conn.commit(); conn.close(); return redirect(url_for("control_center"))

    branch_filter = None if privileged else session.get("branch")
    closings = conn.execute("SELECT * FROM daily_closings WHERE (? IS NULL OR branch=?) ORDER BY closing_date DESC,id DESC LIMIT 100",(branch_filter,branch_filter)).fetchall()
    alerts = conn.execute("SELECT * FROM system_alerts WHERE is_resolved=0 ORDER BY id DESC LIMIT 100").fetchall() if privileged else []
    staff = conn.execute("""SELECT username, SUM(ticket_count) ticket_count, SUM(ticket_value) ticket_value,
        SUM(expense_count) expense_count, SUM(expense_value) expense_value FROM (
        SELECT created_by username,COUNT(*) ticket_count,COALESCE(SUM(total_selling),0) ticket_value,0 expense_count,0 expense_value
        FROM tickets WHERE created_by IS NOT NULL GROUP BY created_by
        UNION ALL
        SELECT created_by username,0,0,COUNT(*),COALESCE(SUM(amount),0) FROM expenses WHERE created_by IS NOT NULL GROUP BY created_by
        ) GROUP BY username ORDER BY username""").fetchall()
    pending = conn.execute("SELECT COUNT(*) FROM expenses WHERE status='PENDING'").fetchone()[0]
    overdue = conn.execute("SELECT COUNT(*) FROM bank_advances a WHERE a.advance_date < date('now','-7 day') AND a.amount > COALESCE((SELECT SUM(r.amount) FROM bank_advance_repayments r WHERE r.advance_id=a.id),0)").fetchone()[0]
    unreconciled = conn.execute("SELECT COUNT(*) FROM bank_transactions WHERE is_reconciled=0").fetchone()[0]
    conn.close()
    return render_template("control.html",closings=closings,alerts=alerts,staff=staff,privileged=privileged,
        pending=pending,overdue=overdue,unreconciled=unreconciled)

# ---------------- NON-IATA SERVICE CHARGE SETTINGS ----------------
@app.route("/settings/service-charges", methods=["GET", "POST"])
@login_required
def service_charge_settings():
    if session.get("role") != "admin":
        flash("Admin pekee ndiye anayeruhusiwa kubadilisha mipangilio hii.", "danger")
        return redirect(url_for("ticket_management"))
    conn = get_db()
    if request.method == "POST":
        tzs = parse_num(request.form.get("fee_tzs"))
        usd = parse_num(request.form.get("fee_usd"))
        if tzs < 0 or usd < 0:
            flash("Service charge haiwezi kuwa chini ya sifuri.", "danger")
        else:
            conn.execute("INSERT INTO app_settings(setting_key,setting_value) VALUES('non_iata_fee_tzs',?) ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value", (str(tzs),))
            conn.execute("INSERT INTO app_settings(setting_key,setting_value) VALUES('non_iata_fee_usd',?) ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value", (str(usd),))
            log_audit(conn, session["user_id"], session["username"], "UPDATE", "SETTINGS", "NON-IATA SERVICE CHARGE", f"TZS={tzs}; USD={usd}")
            conn.commit()
            flash("NON-IATA Service Charge imehifadhiwa kikamilifu.", "success")
    tzs_row = conn.execute("SELECT setting_value FROM app_settings WHERE setting_key='non_iata_fee_tzs'").fetchone()
    usd_row = conn.execute("SELECT setting_value FROM app_settings WHERE setting_key='non_iata_fee_usd'").fetchone()
    conn.close()
    return render_template("service_charge_settings.html", fee_tzs=tzs_row[0] if tzs_row else "10000", fee_usd=usd_row[0] if usd_row else "10")


# ---------------- 13. BACKUP & RESTORE ----------------
@app.route("/backup")
@login_required
def download_backup():
    if session.get("role") not in ("admin","boss"):
        flash("Backup inaruhusiwa kwa Admin au Boss pekee.","danger"); return redirect(url_for("dashboard"))
    db_path = Path(os.environ.get("DB_PATH", Path(__file__).with_name("hawaii_travel.db")))
    if not db_path.exists():
        flash("Database haijapatikana.","danger"); return redirect(url_for("dashboard"))
    # SQLite online backup captures committed data consistently, including WAL content.
    temp_path = Path(os.environ.get("BACKUP_TMP_DIR", str(db_path.parent))) / f"backup_{date.today().isoformat()}.db"
    src = sqlite3.connect(str(db_path)); dst = sqlite3.connect(str(temp_path))
    try: src.backup(dst)
    finally: dst.close(); src.close()
    logconn=get_db(); log_audit(logconn,session["user_id"],session["username"],"BACKUP","SYSTEM","DATABASE","Downloaded database backup"); logconn.commit(); logconn.close()
    return send_file(temp_path,as_attachment=True,download_name=f"hawaii_travel_backup_{date.today().isoformat()}.db")

@app.route("/restore", methods=["POST"])
@login_required
def restore_backup():
    if session.get("role") != "admin":
        flash("Restore inaruhusiwa kwa Admin pekee.","danger"); return redirect(url_for("control_center"))
    uploaded=request.files.get("backup_file")
    if not uploaded or not uploaded.filename.lower().endswith(".db"):
        flash("Chagua SQLite backup file (.db).", "danger"); return redirect(url_for("control_center"))
    target=Path(os.environ.get("DB_PATH", Path(__file__).with_name("hawaii_travel.db")))
    candidate=target.with_name("restore_candidate.db")
    uploaded.save(candidate)
    try:
        check=sqlite3.connect(str(candidate)); result=check.execute("PRAGMA integrity_check").fetchone()[0]; check.close()
        if result != "ok": raise ValueError("Database integrity check failed")
        if target.exists(): shutil.copy2(target, target.with_name(f"hawaii_travel_before_restore_{date.today().isoformat()}.db"))
        # Remove stale SQLite sidecar files so they cannot be replayed against the restored database.
        for suffix in ("-wal", "-shm"):
            sidecar = Path(str(target) + suffix)
            if sidecar.exists(): sidecar.unlink()
        os.replace(candidate,target)
        flash("Database imerudishwa. Ingia tena kuthibitisha taarifa.","success")
    except Exception as exc:
        if candidate.exists(): candidate.unlink()
        flash(f"Restore imeshindikana: {exc}","danger")
    return redirect(url_for("login"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
