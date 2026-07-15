# Nexus Pest Control — Portal (Python version)

Same portal (quotes, jobs, photos, payments, emails) as the original Google
Apps Script app, rebuilt as a plain Python/Flask app so all the code lives
in this one GitHub repo and can be edited normally (any editor, any AI
tool, version history via git) instead of through the Apps Script web
editor.

Data still lives in the same free Google storage as before — the same
Google Sheet ("Nexus Pest Control Portal") for jobs/settings, and the same
Google Drive folder for job photos. This app talks to them over the Google
Sheets API / Drive API instead of being natively bound to them.

Email still sends from `nexuspestcontrolservice@gmail.com`, via Gmail SMTP.

## One-time setup (you need to do these — they involve creating accounts
## and secrets, which I can't do on your behalf)

### 1. Create a Google Cloud service account
1. Go to https://console.cloud.google.com/ and create a new project (any name, e.g. "nexus-portal").
2. In **APIs & Services → Library**, enable **Google Sheets API** and **Google Drive API**.
3. In **APIs & Services → Credentials**, click **Create Credentials → Service account**. Give it any name (e.g. `nexus-portal-bot`).
4. Open the new service account → **Keys** tab → **Add Key → Create new key → JSON**. This downloads a `.json` file — keep it private, never commit it to GitHub.
5. Note the service account's email address (looks like `nexus-portal-bot@your-project.iam.gserviceaccount.com`).

### 2. Share your existing Sheet and Drive folder with the service account
1. Open your existing "Nexus Pest Control Portal" Google Sheet → **Share** → add the service account's email as **Editor**.
2. Open (or create) the Drive folder used for job photos ("Nexus Pest Control - Job Photos") → **Share** → add the service account's email as **Editor**.
3. Copy the Sheet's ID from its URL: `https://docs.google.com/spreadsheets/d/THIS_PART_IS_THE_ID/edit`
4. Copy the Drive folder's ID from its URL: `https://drive.google.com/drive/folders/THIS_PART_IS_THE_ID`

### 3. Create a Gmail App Password
1. On the `nexuspestcontrolservice@gmail.com` account, turn on 2-Step Verification if it isn't already: https://myaccount.google.com/security
2. Go to https://myaccount.google.com/apppasswords, create a new app password (name it "Nexus Portal"), and copy the 16-character code.

### 4. Deploy to Render (free)
1. Create a free account at https://render.com and connect your GitHub.
2. **New → Web Service**, pick this repo. Render will detect `render.yaml` automatically.
3. When prompted for environment variables, enter:
   - `GOOGLE_SERVICE_ACCOUNT_JSON` — paste the **entire contents** of the JSON key file from step 1 (as one line).
   - `GOOGLE_SHEET_ID` — from step 2.
   - `DRIVE_FOLDER_ID` — from step 2.
   - `GMAIL_ADDRESS` — `nexuspestcontrolservice@gmail.com`
   - `GMAIL_APP_PASSWORD` — from step 3.
4. Deploy. Render gives you a URL like `https://nexus-pest-control-portal.onrender.com`.

### 5. Point swastikexpress.ca at the new URL
Once the Render URL is confirmed working, update the `<iframe src="...">` in
`nexuspestcontrol/index.html` in the `swastikexpress.ca` repo to the new
Render URL instead of the old Apps Script `/exec` URL. (I can do this last
step for you once you've completed steps 1–4 and confirm the app works.)

## Local development
```bash
pip install -r requirements.txt
export GOOGLE_SERVICE_ACCOUNT_JSON='<paste json>'
export GOOGLE_SHEET_ID='...'
export DRIVE_FOLDER_ID='...'
export GMAIL_ADDRESS='nexuspestcontrolservice@gmail.com'
export GMAIL_APP_PASSWORD='...'
python app.py
```
Then open http://localhost:5000

## Files
- `app.py` — the entire backend (API + data access + email templates).
- `index.html` — the entire frontend (single page, no build step).
- `requirements.txt` — Python dependencies.
- `render.yaml` — one-click Render deploy config.

## Notes
- Render's free tier spins down after 15 minutes of inactivity — the first
  request after idle time takes ~30-50 seconds to wake up. Upgrading to a
  paid instance ($7/mo) removes that delay if it becomes annoying.
- Access is currently open to anyone with the link, same as the original
  Apps Script version. Locking it down (e.g. simple password) is still on
  the list whenever you're ready.
