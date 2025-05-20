import pandas as pd
from datetime import datetime, timedelta
import smtplib
import os
from email.mime.text import MIMEText
from dotenv import load_dotenv

load_dotenv()

EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASS = os.getenv('EMAIL_PASS')

# quick sanity-check:
print("EMAIL_USER:", EMAIL_USER)
print("EMAIL_PASS set?", bool(EMAIL_PASS))

reminder_data = []

def parse_excel(file_path):
    df = pd.read_excel(file_path)
    reminders = []
    for _, row in df.iterrows():
        try:
            factory = row['Factory Name']
            audit_date = pd.to_datetime(row['Audit Date'])
            schedule = row['Schedule']
            emails = str(row['Email ID']).replace(';', '\n').split('\n')
            reminder_days = int(schedule.split()[0])
            reminder_date = audit_date - timedelta(days=reminder_days)
            reminders.append({
                'factory': factory,
                'audit_date': audit_date.date(),
                'reminder_date': reminder_date.date(),
                'emails': [e.strip() for e in emails if e.strip()],
            })
        except Exception as e:
            print("Skipping row due to error:", e)
            continue
    print(reminders)       
    return reminders

def save_reminders(reminders):
    global reminder_data
    reminder_data = reminders

def get_due_reminders():
    today = datetime.today().date()
    return [r for r in reminder_data if r['reminder_date'] == today]

def send_email(subject, body, to_emails):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = EMAIL_USER
    msg['To'] = ', '.join(to_emails)
    print(msg)
    try:
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.sendmail(EMAIL_USER, to_emails, msg.as_string())
            print("Email sent to", to_emails)
    except Exception as e:
        print("Email sending failed:", e)
