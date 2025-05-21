import os
import pandas as pd
from datetime import datetime
import smtplib
from email.message import EmailMessage

def send_email(subject, body, to_emails):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = os.environ['EMAIL_USER']
    msg['To'] = ', '.join(to_emails)
    msg.set_content(body)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(os.environ['EMAIL_USER'], os.environ['EMAIL_PASS'])
        smtp.send_message(msg)

def main():
    excel_path = 'reminders.xlsx'
    if not os.path.exists(excel_path):
        print("No reminder file found.")
        return

    df = pd.read_excel(excel_path)
    today = datetime.today().date()

    for _, row in df.iterrows():
        reminder_date = row['Reminder Date'].date()
        if reminder_date == today:
            subject = f"Reminder: Upcoming Audit for {row['Factory']}"
            body = f"Dear Team,\n\nAudit for '{row['Factory']}' is on {row['Audit Date'].date()}.\nPlease prepare accordingly.\n\n— Audit Reminder System"
            to_emails = [email.strip() for email in row['Emails'].split(',')]
            send_email(subject, body, to_emails)

if __name__ == "__main__":
    main()
