from flask import Flask, render_template, request, flash
import os
from werkzeug.utils import secure_filename
from utils import parse_excel, save_reminders, get_due_reminders, send_email
from dotenv import load_dotenv

load_dotenv()

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = os.getenv('SECRET_KEY', 'devkey')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.lower().endswith(tuple(ALLOWED_EXTENSIONS))

@app.route('/', methods=['GET', 'POST'])
def index():
    reminders = []
    if request.method == 'POST':
        file = request.files.get('file')
        if file and allowed_file(file.filename):
            path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(path)
            reminders = parse_excel(path)
            save_reminders(reminders)
            flash('Uploaded and parsed!', 'success')
        else:
            flash('Please upload a valid .xlsx file', 'danger')
    return render_template('index.html', reminders=reminders)

@app.route('/run_reminders', methods=['GET'])
def run_reminders():
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
    return {'sent': len(due)}, 200

if __name__ == "__main__":
    app.run(host='0.0.0.0')
