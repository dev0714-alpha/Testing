import os
import csv
import smtplib
import tempfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import platform
from flask import Flask, request, Response, stream_with_context, jsonify, send_from_directory
from flask_cors import CORS
from dotenv import load_dotenv
from openai import OpenAI

# Load environment
load_dotenv(override=True)

GOOGLE_API_KEY = os.getenv('GOOGLE_API_KEY') or os.getenv('OPENAI_API_KEY')

if not GOOGLE_API_KEY:
    print("WARNING: GOOGLE_API_KEY/OPENAI_API_KEY not found. AI endpoints will fail.")

client = OpenAI(api_key=GOOGLE_API_KEY, base_url="https://generativelanguage.googleapis.com/v1beta/openai/") if GOOGLE_API_KEY else None
MODEL = 'gemini-2.5-pro'

# Flask setup
app = Flask(__name__, static_folder='Frontend', static_url_path='')
CORS(app)

# CSV files
STAFF_FILE = 'staff_data.csv'
ADVISERS_FILE = 'advisers.csv'

def get_csv_data(filepath):
    """Safe CSV loading."""
    data = []
    if os.path.exists(filepath):
        with open(filepath, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            data = list(reader)
    return data

# Serve frontend
@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

# API Endpoints
@app.route('/api/departments')
def get_departments():
    data = get_csv_data(STAFF_FILE)
    departments = sorted(list(set(row.get('Department', '') for row in data)))
    return jsonify(departments)

@app.route('/api/assignees')
def get_assignees():
    dept = request.args.get('department', '')
    data = get_csv_data(STAFF_FILE)
    assignees = [{"name": r['Name'], "email": r['Email'], "cc": r.get('CC_Emails', '')} 
                 for r in data if r.get('Department','') == dept]
    return jsonify(assignees)

@app.route('/api/advisers')
def get_advisers():
    assignee_name = request.args.get('assignee', '')
    data = get_csv_data(ADVISERS_FILE)
    advisers = [{"name": r['Name'], "email": r['Email']} 
                for r in data if r.get('Assignee','') == assignee_name]
    return jsonify(advisers)

# AI Feedback Processing
@app.route('/process_feedback', methods=['POST'])
def process_feedback():
    data = request.json
    user_message = data.get('message', '')
    department = data.get('department', 'General')
    assignee_name = data.get('assignee_name', 'Manager')
    assignee_email = data.get('assignee_email', '')
    assignee_cc_str = data.get('assignee_cc', '')

    service_adviser_name = data.get('service_adviser', 'N/A')
    service_adviser_email = data.get('service_adviser_email', '')

    # Merge CCs
    cc_list = [e.strip() for e in assignee_cc_str.split(';') if e.strip()]
    if service_adviser_email and service_adviser_email not in cc_list:
        cc_list.append(service_adviser_email)
    final_cc_string = "; ".join(cc_list)

    # Optional fields
    customer_name = data.get('customer_name', 'N/A')
    contact = data.get('contact', 'N/A')
    vehicle_no = data.get('vehicle_no', 'N/A')
    model = data.get('model', 'N/A')
    km_hr = data.get('km_hr', 'N/A')
    location = data.get('location', 'N/A')
    source = data.get('source', 'N/A')
    ticket_id = data.get('ticket_id', 'N/A')
    workshop = data.get('workshop', 'N/A')

    # System prompt
    system_instruction = f"""
You are a customer complaint assistant.

Department: {department}
Staff: {assignee_name} ({assignee_email})
CC: {final_cc_string}
Ticket: {ticket_id}, Customer: {customer_name}, Contact: {contact}, Vehicle: {vehicle_no}, Model: {model}, km/Hr: {km_hr}, Location: {location}, Source: {source}, Workshop: {workshop}, Adviser: {service_adviser_name}

CLASSIFY as Complaint or Concern and draft email.
"""

    messages = [{"role": "system", "content": system_instruction},
                {"role": "user", "content": user_message}]

    def generate():
        if not client:
            yield "Error: API key not configured."
            return
        try:
            # streaming mode
            stream = client.chat.completions.create(
                model=MODEL,
                messages=messages,
                stream=True
            )
            for chunk in stream:
                content = chunk.choices[0].delta.content
                if content:
                    yield content
        except Exception as e:
            yield f"Error: {str(e)}"

    return Response(stream_with_context(generate()), mimetype='text/plain')

# Cross-platform email fallback
@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json
    to_email = data.get('to', '')
    cc_email = data.get('cc', '')
    html_body = data.get('html_body', '')
    subject = data.get('subject', 'Customer Feedback Review')

    try:
        # Use SMTP (example: Gmail) instead of Outlook COM
        smtp_server = os.getenv('SMTP_SERVER')
        smtp_port = int(os.getenv('SMTP_PORT', 587))
        smtp_user = os.getenv('SMTP_USER')
        smtp_pass = os.getenv('SMTP_PASS')

        if not smtp_server or not smtp_user or not smtp_pass:
            return jsonify({"success": False, "error": "SMTP credentials not configured"}), 500

        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = to_email
        msg['Cc'] = cc_email
        msg['Subject'] = subject
        msg.attach(MIMEText(html_body, 'html'))

        recipients = [to_email] + cc_email.split(';') if cc_email else [to_email]
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, recipients, msg.as_string())

        return jsonify({"success": True})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

# Main
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
