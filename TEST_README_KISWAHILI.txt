HAWAII ERP V5.4.4 - TEST VERSION (LOCAL ONLY)

Hii ni nakala ya majaribio. Haijaunganishwa na Railway.
Database iliyowekwa ni COPY ya backup ya 2026-11-23; backup original haijabadilishwa.

Kuanza (Windows):
1. Install Python 3.11 au 3.13.
2. Fungua PowerShell kwenye folder hii.
3. Run: py -m venv .venv
4. Run: .\.venv\Scripts\Activate.ps1
5. Run: pip install -r requirements.txt
6. Run: py app.py
7. Fungua browser: http://127.0.0.1:5000

MUHIMU:
- Usitumie Railway credentials au production database kwenye test.
- DB_PATH ikibadilishwa, tumia database ya test pekee.
- Hii haijadeployiwa live. Browser test bado inahitaji kufanyika kwenye PC.
