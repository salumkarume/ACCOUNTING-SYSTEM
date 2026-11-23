# HAWAII ERP V5 — Hosting preparation

## Before publishing
1. Create a persistent disk/volume on the host and mount it at `/data`.
2. Set `APP_ENV=production`.
3. Set `DB_PATH=/data/hawaii_travel.db` (must point inside the persistent volume).
4. Set `COOKIE_SECURE=1` and enable HTTPS.
5. Set `SECRET_KEY` to a unique random value (at least 32 random bytes). Never share it or commit it to Git.
6. Install dependencies from `requirements.txt` and start using the included `Procfile`/Gunicorn command.
7. Create the first administrator once, before opening the site to users:
   `ADMIN_USERNAME=... ADMIN_FULL_NAME=... ADMIN_PASSWORD=... python bootstrap_admin.py`
   Use a unique password of at least 9 characters. Do not reuse passwords from other services.
8. Confirm `/data/hawaii_travel.db` exists, log in, create a test record, restart the app, and confirm the record remains.
9. Download a database backup and store a copy outside the hosting provider. Test restore using a separate test database before relying on it.

## Important
- The old predictable demo accounts have been removed. The app will not create public default usernames/passwords.
- SQLite data is persistent only if `DB_PATH` points to a persistent mounted volume. The server's temporary filesystem may be erased on redeploy/restart.
- Use one Gunicorn worker with threads for this SQLite deployment. For many simultaneous users, migrate to PostgreSQL before scaling workers.
- Configure off-site backups. Hosting uptime is not a backup and no free host can guarantee uninterrupted 24/7 availability.
- This is a hosting-hardening preparation, not a full independent security audit. Before entering real company financial/customer data, verify role permissions, HTTPS, backups and restore procedures, and test all workflows.
- For local HTTP testing only, set `COOKIE_SECURE=0`. Keep it `1` on the public HTTPS site.
