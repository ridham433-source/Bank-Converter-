# BankConverter

Upload a bank statement PDF, get a clean Excel/CSV file back.

## How to put this online for FREE (no coding, ~10 minutes)

### Step 1 — Create a free GitHub account
Go to https://github.com and sign up (skip if you already have one).

### Step 2 — Create a new repository
1. Click the **+** icon (top right) → **New repository**
2. Name it anything, e.g. `bankconverter`
3. Set it to **Public**
4. Click **Create repository**

### Step 3 — Upload these files
1. On your new repo's page, click **Add file → Upload files**
2. Drag in both `app.py` and `requirements.txt` (in this folder)
3. Click **Commit changes**

### Step 4 — Deploy on Streamlit Community Cloud (free, no card)
1. Go to https://share.streamlit.io
2. Sign in with your GitHub account
3. Click **New app**
4. Pick your `bankconverter` repository
5. Make sure the file path says `app.py`
6. Click **Deploy**

That's it. Wait about a minute — your app will be live at a public URL like
`https://your-app-name.streamlit.app`, and you can share that link with anyone.

## If something breaks
Come back to this chat, tell me exactly what happened (a screenshot of the
error is ideal), and I'll fix the code — no cost, no limit on how many times.

## What this app does
- Accepts a bank statement PDF (works with password-protected PDFs too)
- Automatically finds the transaction table, whether the bank uses visible
  ruled lines (like Saraswat) or an invisible borderless layout (like HDFC)
- Lets you pick which columns to keep
- Gives you a preview, then an Excel or CSV download

Tested against real HDFC Bank and Saraswat Bank statements with the running
balance validated row-by-row against the statement's own stated balance —
zero mismatches on either file.
