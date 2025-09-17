import pandas as pd
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()

# Your details from .env
YOUR_NAME = os.getenv("YOUR_NAME")
YOUR_EMAIL = os.getenv("YOUR_EMAIL")
YOUR_PASSWORD = os.getenv("YOUR_PASSWORD")
EMAIL_SUBJECT = "Collaborate Smarter — Development That Pays Off"

# Load spreadsheet
df = pd.read_excel("testspread.xlsx")

# Load HTML template
with open("Email format.html", "r", encoding="utf-8") as f:
    TEMPLATE = f.read()

# Setup SMTP
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(YOUR_EMAIL, YOUR_PASSWORD)

for _, row in df.iterrows():
    if pd.notna(row["Email"]):
        designer_name = str(row["Designer Name"]).strip()

        # Replace placeholders
        message = TEMPLATE.replace("[designer name]", designer_name)

        # Create email (HTML format)
        msg = MIMEText(message, "html", "utf-8")
        msg["Subject"] = EMAIL_SUBJECT
        msg["From"] = YOUR_EMAIL
        msg["To"] = row["Email"]

        # Send email
        server.sendmail(YOUR_EMAIL, row["Email"], msg.as_string())
        print(f"✅ Email sent to {designer_name} -> {row['Email']}")

server.quit()
