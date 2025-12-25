Apollo Automation for 180DC
===========================

This toolkit generates personalized outreach emails with Gemini and sends them in bulk via Gmail SMTP. Use `apollo.py` to draft emails from a contact list, then `send_emails_smtp.py` to deliver them.

Prerequisites
-------------
- Python 3.10+ installed.
- Gmail account with 2FA and an App Password for SMTP (regular passwords will fail).
- Google Generative AI API key (set as `GOOGLE_API_KEY`).
- Contact list as CSV or Excel with at least these columns: `First Name`, `Last Name`, `Email`. Optional columns used for personalization: `Title`, `Company Name`, `Industry`, `Keywords`, `Website`, `Email Status`.

Setup
-----
1) Get the code locally (clone or download this folder).
2) Create a virtual environment and install dependencies:
```
python -m venv .venv
.\.venv\Scripts\activate
pip install pandas google-generativeai python-dotenv openpyxl
```
3) Add your Google API key to `.env` in the project root:
```
GOOGLE_API_KEY=your_gemini_api_key_here
```
4) Prepare your contacts spreadsheet (CSV or Excel). Keep the required columns and fill optional fields for better personalization. Avoid rows with missing emails; the generator will skip them.

Generate outreach emails (apollo.py)
------------------------------------
1) Run the generator:
```
python apollo.py
```
2) When prompted:
   - Provide the path to your contacts file (CSV/XLSX/XLS).
   - Choose whether to customize the email template. If you answer "no", the built-in 180DC IIT Kharagpur template is used. If "yes", you can supply your own description, tone, key points, and template text (placeholders like `$FIRST_NAME`, `$LAST_NAME`, `$COMPANY`, `$EMAIL_BODY` are supported).
3) The script will:
   - Load and validate contacts (warns about missing/invalid emails).
   - Generate a short company-specific subject and body per contact using Gemini (model `models/gemini-2.5-flash`).
   - Save results to `generated_emails.json` (for automation) and `generated_emails.txt` (for quick review).
4) Optional: view a sample email when prompted.

Output format (generated_emails.json)
-------------------------------------
Each entry looks like:
```
{
  "contact": "First Last",
  "email_address": "user@example.com",
  "subject": "180DC IIT Kharagpur X ExampleCo",
  "generated_email": "Respected Dr Last,\n...\nBest regards,\nParth Sethi\n..."
}
```
This file is the input for the sender script.

Send emails via Gmail SMTP (send_emails_smtp.py)
------------------------------------------------
1) Ensure `generated_emails.json` is present in the project root (created by `apollo.py`).
2) Run the sender:
```
python send_emails_smtp.py
```
3) First run will create `credentials.py`. Enter:
   - Gmail address (defaults to `parthsethi@180dc.org` if you press Enter).
   - Gmail App Password (create in Google Account > Security > App passwords).
4) The script:
   - Imports credentials from `credentials.py`.
   - Uses CC recipients hardcoded in the script (`arghyadip.pal@180dc.org`, `mishra.moulik@180dc.org`). Edit `cc_emails` in `send_emails_smtp.py` if you want to change these.
   - Reads all messages from `generated_emails.json`.
   - Sends a test email to your own address to verify SMTP login.
   - Asks for confirmation (`Do you want to send X emails?`). Type `yes` to proceed.
   - Sends each email with a 2-second delay to respect Gmail rate limits. Successes and failures are printed, and a summary is shown at the end.

Tips and troubleshooting
------------------------
- Gmail SMTP requires an App Password (not your normal password). If the test email fails, regenerate an App Password and update `credentials.py`.
- If Gemini returns an error, verify `GOOGLE_API_KEY` in `.env` and your network access. Model name is set to `models/gemini-2.5-flash` in `apollo.py`.
- Make sure your contacts file path is correct and includes the required columns; missing emails are skipped.
- Review `generated_emails.txt` before sending if you want to spot-check tone or personalization.
