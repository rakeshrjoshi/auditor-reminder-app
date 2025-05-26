from utils import parse_excel, get_due_reminders, send_email, save_reminders
import os

EXCEL_PATH = os.path.join('uploads', 'reminders.xlsx')

def main():
    if not os.path.exists(EXCEL_PATH):
        print("No Excel file found. Skipping reminders.")
        return

    # Parse and load reminders
    reminders = parse_excel(EXCEL_PATH)
    save_reminders(reminders)

    # Filter today's reminders
    due = get_due_reminders()
    for r in due:
        subject = f"Reminder: Upcoming Audit for {r['factory']}"
        body = (
            f"Dear “{r['factory']}” Quality Team,\n\n"
            f"Audit for “{r['factory']}” is on {r['audit_date']}.\n"
            f"Please prepare.\n\n"
            f"– Technical Team"
        )
        send_email(subject, body, r['emails'])
        print(f"Reminder sent to {r['emails']} for {r['factory']}")

    if not due:
        print("No reminders due today.")

if __name__ == "__main__":
    main()
