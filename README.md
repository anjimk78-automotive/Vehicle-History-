# Vehicle History Monitoring System

A Streamlit app to **enter** and **view** vehicle history records, backed by your Google Sheet:
https://docs.google.com/spreadsheets/d/1uHRV6X1xYkid9XhUYULK3xrDBFVHW7oStMaLi7oY2es/edit

## ⚠️ Security first — rotate your key

You uploaded a live Google Cloud service account key (`kmn-vehicle-history@kmn-automotive.iam.gserviceaccount.com`)
into this chat. Treat it as **exposed**:

1. Go to Google Cloud Console → **IAM & Admin → Service Accounts**.
2. Open `kmn-vehicle-history@kmn-automotive.iam.gserviceaccount.com`.
3. Under **Keys**, delete the key with ID starting `665c9b02f...` and create a **new** key.
4. Use only the *new* key below. Never paste key contents into a chat again — upload it directly
   into your deployment environment instead.

## 1. Share the Sheet with the service account

Open the Google Sheet → **Share** → add the service account's `client_email`
(`kmn-vehicle-history@kmn-automotive.iam.gserviceaccount.com`) as an **Editor**.

## 2. Install dependencies

```bash
pip install -r requirements.txt
```

## 3. Add your (new) credentials

Choose ONE:

**Option A — local file (simplest, for local use only):**
Save your new key as `service_account.json` in this same folder (already `.gitignore`d — never commit it).

**Option B — Streamlit secrets (recommended for deployment):**
Create `.streamlit/secrets.toml`:

```toml
[gcp_service_account]
type = "service_account"
project_id = "kmn-automotive"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "kmn-vehicle-history@kmn-automotive.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

If deploying to Streamlit Community Cloud, paste this same block into
**App settings → Secrets** instead of a local file.

## 4. Run

```bash
streamlit run app.py
```

## What it does

- **➕ Add Entry** — form to log a new vehicle history record (VIN, make/model/year, owner,
  mileage, event type, cost, notes, etc.) directly into the Sheet.
- **📋 View / Search History** — browse all records with filters (VIN, event type, status),
  summary metrics, and CSV export.
- **✏️ Edit / Delete** — update or remove an existing row by its row number.

The app auto-creates the header row in your Sheet on first run if it's missing, so you
don't need to pre-format anything — just make sure the tab is named `Sheet1` (or edit
`WORKSHEET_NAME` in `app.py`).
