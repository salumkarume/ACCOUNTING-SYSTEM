# Hawaii Travel Agency Company Ltd - Accounting & Reconciliation ERP (v3.0)

Maboresho makubwa ya **Data Integration & Synchronization (v3.0)**:
1. **Customer & Debt Interaction (Muunganiko Kamili wa Wateja na Madeni)**:
   - Unaposajili mteja mpya (pamoja na deni lake la awali la TZS/USD kama analo), jina lake na salio lake vinaonekana papo hapo kwenye **Customer Payments (Kurejesha Deni)**, **Ticket Management**, **Dashboard**, na **Financial Reports**.
   - Ukichagua mteja kwenye sehemu ya kurejesha deni, mfumo unavuta taarifa zake live (Deni lake TZS/USD, Tawi, Namba ya Simu, na Tiketi zake ambazo hazijalipwa).
2. **Two-Way Ledger Synchronization (Muunganiko wa Pande Mbili)**:
   - Ukirekodi, ukiedit, au ukifuta **Tiketi**, **Malipo ya Deni la Mteja**, **Matumizi (Approved Expenses)**, **Malipo ya BSP**, au **Director Transfers**, salio la **Bank**, **Cash (Dar/Zanzibar)**, na **Lipa Namba** linajirekebisha lenyewe bila kuacha rekodi hewa (orphan records).
3. **Auto BSP Re-Matching**:
   - Ukirekodi au kuedit tiketi baada ya kupakia BSP statement, mfumo unalinganisha upya tiketi na BSP automatically.


## HAWAII ERP v5 – Advanced Financial Control

- Menu 12: Daily Closing, Staff Accountability & Control Center.
- Menu 13: Backup download (Admin/Boss); database restore is available in the Control Center for Admin.
- Existing V4 modules are retained.

## Ticket Register – Date Range & Print
Open **04. Ticket Management**. In the **Date Range** section, select a start date and end date using the calendar controls, optionally choose DAR, ZANZIBAR, or all branches, then click **Tafuta Tiketi**. Click **Print / Chapisha** to print the filtered ticket register, including TZS and USD totals. Use the browser print dialog to select a printer or save as PDF.

## Password management and recovery
- Every user can change their own password from **Badili Password** after signing in.
- The login page includes **Umesahau password?**. Users enter the email address saved on their account.
- Email recovery requires SMTP environment variables (`SMTP_HOST`, `SMTP_PORT`, `SMTP_USERNAME`, `SMTP_PASSWORD`, `SMTP_FROM`) and `PUBLIC_BASE_URL` to be configured in the hosting dashboard. Reset links expire after 30 minutes and can only be used once.
- Admin should add a valid, unique recovery email for each user. Existing users can add/update their email in **Users & Audit Trail**.
- Passwords must be at least 9 characters. Admin can also reset a user's password from User Management.
- Never put SMTP credentials or the production `SECRET_KEY` in source code or commit them to a public repository.
