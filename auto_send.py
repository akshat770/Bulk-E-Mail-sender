import smtplib
from email.message import EmailMessage
import time
import os
from dotenv import load_dotenv
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Load email config
load_dotenv()
SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")

# Google Sheets setup
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPE)
client = gspread.authorize(creds)

# Open Google Sheet by URL
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwxyz)
worksheet = sheet.sheet1  # first sheet
rows = worksheet.get_all_records()  # list of dicts

# Function to send email
def send_email(to_email, name="Friend"):
    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = to_email
    msg["Subject"] = "Interested in Research Work Under Your Supervision"
    # Plain text fallback
    msg.set_content("""\your message""")

# HTML version
    msg.add_alternative("""\
<html>
  <body>
    <p>your message</p>
  </body>
</html>
""", subtype="html")

    # Attach resumes
    for file in ["1.pdf", "2.pdf", "3.pdf", "4.pdf"]:
        with open(file, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename=file
            )

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(SMTP_USER, SMTP_PASS)
            smtp.send_message(msg)
            print(f"Sent to {to_email}")
    except Exception as e:
        print(f"Error sending to {to_email}: {e}")

# Loop over rows
for row in rows:
    email = row.get("Mail Id")
    if email:
        send_email(email)
        time.sleep(3)  # delay between sends


print(rows)
