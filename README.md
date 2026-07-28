# Invoice Automation

Flask web app that automates hotel invoice generation on the basmaemaargroup CRM. Upload an Excel file with voucher/hotel/room data, and the bot logs into the CRM, fetches each voucher's booking page, checks if invoices already exist, and generates new ones with calculated room prices.

## Room Rate Calculation

| Room Type | Formula       |
|-----------|---------------|
| Sharing   | Rate ÷ 1     |
| Double    | Rate ÷ 2     |
| Triple    | Rate ÷ 3     |
| Quad      | Rate ÷ 4     |
| Quint     | Rate ÷ 5     |

## Excel Format

File must be `.xlsx` with these columns (row 1 = headers). Room Type is read from the CRM page directly, not from Excel:

| Voucher # | Hotel | Room Rate |
|-----------|-------|-----------|
| 102524    | TALAL MASHAD | 100 |
| 102524    | DIYAR AL HIJAZ | 135 |

## Features

- PIN-only login
- Mobile-responsive UI
- Excel upload with instant voucher preview
- Background processing (survives tab close)
- Concurrency lock (one file at a time)
- Live logs during processing
- Dashboard with results summary
- No server-side file storage (Vercel-compatible)

## Setup

```bash
pip install -r requirements.txt
```

Create a `.env` file:

```
USER_NAME=your_crm_username
PASSWORD=your_crm_password
PIN=your_app_pin
FLASK_SECRET_KEY=your_secret_key
```

## Run Locally

```bash
python3 app.py
```

Open `http://127.0.0.1:5000`

## Deploy to Vercel

1. Push to GitHub
2. Import repo in Vercel
3. Set environment variables in Vercel dashboard
4. Deploy — `vercel.json` is already configured

## How It Works

1. User logs in with PIN
2. Uploads Excel file → unique voucher numbers displayed
3. Clicks Start → background thread processes each voucher:
   - Logs into CRM
   - Fetches voucher booking page
   - Matches Excel hotel names to booking page hotels
   - If invoice already exists (checkbox checked + price > 0) → skips
   - Otherwise: computes price, checks custom sale checkbox, sets price, sets note to "Done", POSTs update
4. Results displayed on dashboard with stats

## Files

- `app.py` — Flask application
- `main.py` — CRM automation logic
- `templates/login.html` — PIN login screen
- `templates/dashboard.html` — Main dashboard
- `vercel.json` — Vercel deployment config
- `requirements.txt` — Python dependencies
