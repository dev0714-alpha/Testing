import os
import csv
import json
import tempfile
import platform
import subprocess
import re
from flask import Flask, request, Response, stream_with_context, jsonify
from flask_cors import CORS
from dotenv import load_dotenv
from openai import OpenAI

# Try importing pywin32 for Outlook integration
try:
    import win32com.client
    import pythoncom
    OUTLOOK_AVAILABLE = True
except ImportError:
    OUTLOOK_AVAILABLE = False
    print("Warning: 'pywin32' not installed. Outlook integration will be disabled.")

# 1. Setup Environment
load_dotenv(override=True)
# Prefer GOOGLE_API_KEY but fall back to OPENAI_API_KEY if provided
google_api_key = os.getenv('GOOGLE_API_KEY') or os.getenv('OPENAI_API_KEY')

if not google_api_key:
    print("WARNING: GOOGLE_API_KEY/OPENAI_API_KEY not found. Please check your .env file.")
    client = None
else:
    gemini_url = "https://generativelanguage.googleapis.com/v1beta/openai/"
    client = OpenAI(api_key=google_api_key, base_url=gemini_url)

MODEL = 'gemini-2.5-pro'

# Serve the frontend folder as static files so `index.html` is available at `/`
app = Flask(__name__, static_folder='Frontend', static_url_path='')
CORS(app)


@app.route('/')
def index():
    """Serve `Frontend/index.html` as the root page."""
    return app.send_static_file('index.html')

# 2. Data Loading Logic
STAFF_FILE = 'staff_data.csv'
ADVISERS_FILE = 'advisers.csv'

def get_csv_data(filepath):
    """Reads a CSV and returns a list of dictionaries."""
    data = []
    if os.path.exists(filepath):
        with open(filepath, mode='r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                data.append(row)
    return data

@app.route('/api/departments', methods=['GET'])
def get_departments():
    """Returns unique departments from staff data."""
    data = get_csv_data(STAFF_FILE)
    departments = sorted(list(set(row['Department'] for row in data)))
    return jsonify(departments)

@app.route('/api/assignees', methods=['GET'])
def get_assignees():
    """Returns assignees for a specific department."""
    dept = request.args.get('department')
    data = get_csv_data(STAFF_FILE)
    # Filter by department and return Name, Email, AND CC
    assignees = [
        {
            "name": row['Name'], 
            "email": row['Email'],
            "cc": row.get('CC_Emails', '') 
        } 
        for row in data if row['Department'] == dept
    ]
    return jsonify(assignees)

@app.route('/api/advisers', methods=['GET'])
def get_advisers():
    """Returns service advisers for a specific Assignee."""
    assignee_name = request.args.get('assignee') # CHANGED: Filter by Assignee
    data = get_csv_data(ADVISERS_FILE)
    advisers = [
        {"name": row['Name'], "email": row['Email']}
        for row in data if row['Assignee'] == assignee_name
    ]
    return jsonify(advisers)

# 3. Enhanced AI Logic
@app.route('/process_feedback', methods=['POST'])
def process_feedback():
    data = request.json
    user_message = data.get('message', '')
    department = data.get('department', 'General')
    assignee_name = data.get('assignee_name', 'Manager')
    assignee_email = data.get('assignee_email', '')
    assignee_cc_str = data.get('assignee_cc', '') 
    
    # Get Selected Service Adviser
    service_adviser_name = data.get('service_adviser', 'N/A')
    service_adviser_email = data.get('service_adviser_email', '')

    # Logic: Merge Assignee CCs + Service Adviser Email
    cc_list = [email.strip() for email in assignee_cc_str.split(';') if email.strip()]
    
    if service_adviser_email and service_adviser_email not in cc_list:
        cc_list.append(service_adviser_email)
    
    final_cc_string = "; ".join(cc_list)

    # Extract optional fields
    customer_name = data.get('customer_name', 'N/A')
    contact = data.get('contact', 'N/A')
    vehicle_no = data.get('vehicle_no', 'N/A')
    model = data.get('model', 'N/A')
    km_hr = data.get('km_hr', 'N/A')
    location = data.get('location', 'N/A')
    source = data.get('source', 'N/A')
    ticket_id = data.get('ticket_id', 'N/A')
    workshop = data.get('workshop', 'N/A')

    # Constructed System Prompt
    system_instruction = f"""
    You are a customer complaint management assistant at an automotive services company.
    
    Current Context:
    - Selected Department: {department}
    - Assigned Staff: {assignee_name} ({assignee_email})
    - CC Recipients: {final_cc_string}
    
    Ticket & Vehicle Information:
    - Ticket ID: {ticket_id}
    - Customer Name: {customer_name}
    - Contact: {contact}
    - Vehicle/Chassis No: {vehicle_no}
    - Model: {model}
    - Odometer (km/Hr): {km_hr}
    - Location: {location}
    - Source: {source}
    - Workshop: {workshop}
    - Service Adviser: {service_adviser_name}
    
    CLASSIFICATION GUIDELINES:

    **1. COMPLAINT (Actionable)**
    * **Definition:** The customer requires a tangible solution, rectification, or compensation from our end.
    * **Criteria:** The issue implies a failure in our service that necessitates immediate action (e.g., re-doing a repair, fixing a new fault caused by us, refunding a charge).
    * **Key Signals:** "Redo," "Fix this," "Broken after service," "Demanding solution."

    **2. CONCERN (Non-Actionable / Improvement Feedback)**
    * **Definition:** The customer provides feedback where no immediate solution is required/possible for that specific transaction, but the feedback is valuable for future process improvement.
    * **Criteria:**
        * Service was completed but the experience was sub-optimal (e.g., delays, long wait times, poor washing quality).
        * Issues related to work done at *other* workshops (competitors).
        * General advice on how to improve.
    * **Key Signals:** "Slow," "Late," "Dirty," "Next time," "Other shop did X."
    
    YOUR TASKS:
    1. Analyze the customer description based on the CLASSIFICATION GUIDELINES above.
    2. CLASSIFY it strictly as "Complaint" or "Concern".
    3. DRAFT AN EMAIL to the Assigned Staff member ({assignee_name}) summarizing the issue.
    
    FORMAT YOUR RESPONSE EXACTLY LIKE THIS:
    
    ### CLASSIFICATION
    [Complaint or Concern]
    
    ### DRAFT EMAIL
    **Subject:** [Create a professional subject line including Ticket ID and Source if available]
    
    Dear {assignee_name},

    [Brief opening statement regarding the customer feedback]

    [Write a professional, minima and concise email body here analyzing the issue based on the user description.]

    | Field | Detail |
    | :--- | :--- |
    | **Ticket ID** | {ticket_id} |
    | **Customer** | {customer_name} |
    | **Contact** | {contact} |
    | **Vehicle No** | {vehicle_no} |
    | **Model** | {model} |
    | **km/Hr** | {km_hr} |
    | **Location** | {location} |
    | **Source** | {source} |
    | **Workshop** | {workshop} |
    | **Adviser** | {service_adviser_name} |
    
    Regards,\n
    Automated Support System
    """

    messages = [
        {"role": "system", "content": system_instruction},
        {"role": "user", "content": user_message}
    ]

    def generate():
            # If no API client is configured, return a clear error message instead
            if client is None:
                yield "Error: OpenAI/Google API key not configured on server. Set GOOGLE_API_KEY or OPENAI_API_KEY in .env"
                return

            try:
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

# 4. Outlook Integration Endpoint
@app.route('/open_outlook', methods=['POST'])
def open_outlook():
    data = request.json
    to_email = data.get('to', '')
    cc_email = data.get('cc', '') 
    html_body = data.get('html_body', '')
    
    # 1. Automatic Subject Extraction & Body Cleanup
    subject = "Customer Feedback Review" 

    subject_pattern = re.compile(r'Subject:</strong>\s*(.*?)(?:<br>|</p>)', re.IGNORECASE | re.DOTALL)
    match = subject_pattern.search(html_body)
    
    if match:
        raw_subject = match.group(1)
        subject = re.sub(r'<[^>]+>', '', raw_subject).strip()
        html_body = html_body[match.end():].strip()

    # 2. Enhanced Email Styling (CSS)
    email_styles = """
    <style>
        body { font-family: Calibri, sans-serif; font-size: 11pt; color: #333; }
        table { border-collapse: collapse; width: 100%; max-width: 600px; margin: 15px 0; }
        th, td { border: 1px solid #d1d5db; padding: 8px 12px; text-align: left; }
        th { background-color: #f3f4f6; font-weight: bold; color: #111; }
        td { background-color: #ffffff; }
        strong { color: #000; }
        p { margin-bottom: 10px; }
    </style>
    """
    
    full_html = f"""
    <html>
    <head>{email_styles}</head>
    <body>
        {html_body}
    </body>
    </html>
    """

    # Method 1: Direct Outlook COM
    if OUTLOOK_AVAILABLE:
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0) 
            mail.To = to_email
            mail.CC = cc_email
            mail.Subject = subject
            mail.HTMLBody = full_html
            mail.Display() 
            return jsonify({"success": True})
        except Exception as e:
            print(f"Direct Outlook COM failed ({e}). Falling back to .eml file...")

    # Method 2: Generate .eml file
    try:
        with tempfile.NamedTemporaryFile(suffix=".eml", delete=False, mode='w', encoding='utf-8') as f:
            filename = f.name
            f.write(f"To: {to_email}\n")
            if cc_email:
                f.write(f"Cc: {cc_email}\n")
            f.write(f"Subject: {subject}\n")
            f.write("X-Unsent: 1\n")
            f.write("Content-Type: text/html; charset=utf-8\n\n")
            f.write(full_html)
        
        if platform.system() == 'Windows':
            os.startfile(filename)
        elif platform.system() == 'Darwin': 
            subprocess.call(('open', filename))
        else: 
            subprocess.call(('xdg-open', filename))
            
        return jsonify({"success": True, "note": "Opened via EML fallback"})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

if __name__ == '__main__':
    # production
    # production: respect PORT env var so Render (or other PaaS) can set it
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)