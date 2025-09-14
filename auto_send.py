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
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1YjFo3yRsdJJ6cNxyfkcLRvtjL-4kAzv4hIg0esQY5ME/edit?usp=sharing")
worksheet = sheet.sheet1  # first sheet
rows = worksheet.get_all_records()  # list of dicts

# Function to send email
def send_email(to_email, name="Friend"):
    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = to_email
    msg["Subject"] = "Interested in Research Work Under Your Supervision"
    # Plain text fallback
    msg.set_content("""\
Respected Sir,

We hope this email finds you well. We are students of the National Institute of Technology, Srinagar, currently in our third year of B.Tech in Information Technology. We are writing to express our keen interest in securing a research internship under your guidance during the winter break (15 December 2025 – 10 February 2026).

Our academic background and technical exposure have equipped us with a solid foundation in Python, C++, and SQL, along with hands-on experience in Machine Learning libraries such as NumPy, Pandas, Scikit-learn, TensorFlow, and PyTorch. We have also studied Data Structures, Algorithms, Probability & Statistics, and Database Management Systems, which complement our interest in Artificial Intelligence and Machine Learning research.

We are highly motivated to explore real-world applications of ML and would be truly grateful for the opportunity to learn and contribute under your mentorship. Please find our resumes attached for your kind consideration.

Thank you for your time and consideration.

Sincerely,
Jatin Kumar
Ankit Maholiya
Lucky Yadav
Akshat Singh
""")

# HTML version
    msg.add_alternative("""\
<html>
  <body>
    <p>Respected Sir,</p>
    <p>We hope this email finds you well. We are students of the <b>National Institute of Technology, Srinagar</b>, currently in our third year of <b>B.Tech in Information Technology</b>. We are writing to express our keen interest in securing a research internship under your guidance during the winter break (15 December 2025 – 10 February 2026).</p>

    <p>Our academic background and technical exposure have equipped us with a solid foundation in <b>Python, C++, and SQL</b>, along with hands-on experience in Machine Learning libraries such as <b>NumPy, Pandas, Scikit-learn, TensorFlow, and PyTorch</b>. We have also studied Data Structures, Algorithms, Probability & Statistics, and Database Management Systems, which complement our interest in Artificial Intelligence and Machine Learning research.</p>

    <p>We are highly motivated to explore real-world applications of ML and would be truly grateful for the opportunity to learn and contribute under your mentorship. Please find our resumes attached for your kind consideration.</p>

    <p>Thank you for your time and consideration.</p>

    <p>Sincerely,<br>
    Jatin Kumar<br>
    Ankit Maholiya<br>
    Lucky Yadav<br>
    Akshat Singh</p>
  </body>
</html>
""", subtype="html")

    # Attach resumes
    for file in ["Akshat_resume.pdf", "Ankit_resume.pdf", "Jatin_resume.pdf", "Lucky_resume.pdf"]:
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
