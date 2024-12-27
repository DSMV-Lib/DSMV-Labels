from flask import Flask, request, jsonify, render_template_string
from flask_mail import Mail, Message
from fpdf import FPDF
import requests
import os
import threading

app = Flask(__name__)

# Configure Flask-Mail
app.config['MAIL_SERVER'] = 'smtp.office365.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'Bot_donotreply@dynatecsmv.no'
app.config['MAIL_PASSWORD'] = 'NeverReply2024'
app.config['MAIL_DEFAULT_SENDER'] = 'Bot_donotreply@dynatecsmv.no'

mail = Mail(app)

class PDF(FPDF):
    def __init__(self):
        super().__init__(orientation='P', unit='mm', format='A4')
        # Add Unicode font support
        self.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
        self.add_font('DejaVu', 'B', 'DejaVuSansCondensed-Bold.ttf', uni=True)
        self.add_font('DejaVu', 'I', 'DejaVuSansCondensed-Oblique.ttf', uni=True)

    def header(self):
        self.image('logo.jpg', x=self.w - 90, y=10, w=80)

    def footer(self):
        self.set_draw_color(0, 176, 80)  # RGB for #00B050
        self.set_line_width(2)
        self.rect(5, 5, 200, self.h / 2 - 5)

def create_pallet_label(data_array, filename):
    pdf = PDF()
    pdf.set_auto_page_break(auto=False)

    def clean_text(text):
        if not text:
            return text
        # Remove problematic characters but keep Norwegian letters
        return ''.join(char for char in text if char.isprintable() or char in 'æøåÆØÅ')

    # Sort the data array by PositionID
    data_array.sort(key=lambda x: x.get('PositionID', ''))

    for data in data_array:
        pdf.add_page()

        # QR Code generation
        qr_code_data = clean_text(data.get('qr', 'DEFAULT_STORAGE_ID'))
        qr_code_url = f"https://api.qrserver.com/v1/create-qr-code/?size=150x150&data={qr_code_data}"
        qr_code_response = requests.get(qr_code_url)

        if qr_code_response.status_code == 200:
            qr_code_filename = f"qr_code_{qr_code_data}.png"
            with open(qr_code_filename, 'wb') as f:
                f.write(qr_code_response.content)

        pdf.image(qr_code_filename, x=10, y=20, w=50, h=50)
        os.remove(qr_code_filename)

        # Use DejaVu font for text with Norwegian characters
        pdf.set_font('DejaVu', 'B', 40)
        pdf.set_xy(10, 80)
        pdf.cell(0, 10, clean_text(data.get('Customer + Order', 'DEFAULT_ORDER')), ln=True)

        pdf.set_font('DejaVu', '', 25)
        pdf.set_xy(10, pdf.get_y() + 5)
        content = clean_text(data.get('Content', 'Default Content'))
        pdf.multi_cell(0, 8, content)

        remaining_space = (pdf.h / 2) - pdf.get_y() - 20
        if remaining_space > 0:
            pdf.set_y(pdf.get_y() + remaining_space / 2)

        pdf.set_font('DejaVu', '', 12)
        pdf.cell(0, 8, f"Owner: {clean_text(data.get('Owner', 'Default Owner'))}", ln=True)
        pdf.cell(0, 8, f"Created By: {clean_text(data.get('Created By', 'Default Creator'))}", ln=True)

        pdf.set_y(pdf.get_y() + 10)

        label_revision = clean_text(data.get('LabelRevision', 'DEFAULT_LABEL_REVISION'))
        position_id = clean_text(data.get('PositionID', 'DEFAULT_POSITION_ID'))
        pdf.set_xy(10, 10)
        pdf.set_font('DejaVu', 'I', 10)
        pdf.cell(0, 10, f"{label_revision} - {position_id}", 0, 1, 'L')

        mail_address = clean_text(data.get('MailAdress', 'default@mail.com'))

    pdf.output(filename, 'F')
    return mail_address, [clean_text(data['StorageID']) for data in data_array]

@app.route('/')
def home():
    return render_template_string('''
        <!doctype html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Coming Soon</title>
            <style>
                body { display: flex; justify-content: center; align-items: center; height: 100vh; background-color: #f8f9fa; font-family: Arial, sans-serif; text-align: center; }
                h1 { font-size: 3em; color: #343a40; }
                p { font-size: 1.5em; color: #6c757d; }
            </style>
        </head>
        <body>
            <div>
                <h1>Coming Soon</h1>
                <p>We are working hard to get this site up and running. Stay tuned!</p>
            </div>
        </body>
        </html>
    ''')

@app.route('/fetch-and-generate', methods=['GET'])
def fetch_and_generate():
    try:
        glide_api_url = 'https://functions.prod.internal.glideapps.com/api/apps/WALpghAXNz7kVRH9HFdx/tables/native-table-7HWStsovzmSyVBveLMUe/rows'
        headers = {
            'user-agent': 'Make/production',
            'authorization': 'Bearer 4ba1af06-5ac2-4346-a5b6-b9ded23fafa8',  # Replace with your actual token
            'x-glide-client-id': '3def4919-730b-4dac-9a18-5478c9d4ea0c'
        }

        response = requests.get(glide_api_url, headers=headers)

        if response.status_code == 200:
            data = response.json()
            data_array = []
            for row in data.get('data', {}).get('rows', []):
                entry_dict = {
                    'Customer + Order': row.get('Name', 'DEFAULT_ORDER'),
                    'StorageID': row.get('$rowID', 'DEFAULT_STORAGE_ID'),
                    'MailAdress': row.get('383W6', 'default@mail.com'),
                    'Owner': row.get('loQhD', 'Default Owner'),
                    'qr': row.get('knlbN', 'DEFAULT_STORAGE_ID'),
                    'Created By': row.get('dVWZJ', 'Default Creator'),
                    'Content': row.get('PyIlB', 'Default Content'),
                    'LabelRevision': row.get('edrDV', 'DEFAULT_LABEL_REVISION'),
                    'PositionID': row.get('LVx14', 'DEFAULT_POSITION_ID'),
                }
                data_array.append(entry_dict)

            grouped_data = {}
            all_row_ids = []
            for entry in data_array:
                name = entry.get('Customer + Order', 'Unknown')
                if name not in grouped_data:
                    grouped_data[name] = {}
                if entry['StorageID'] not in grouped_data[name]:
                    grouped_data[name][entry['StorageID']] = entry
                all_row_ids.append(entry['StorageID'])

            for name, group in grouped_data.items():
                pdf_filename = f"pallet_labels_{name}.pdf"
                mail_address, _ = create_pallet_label(list(group.values()), pdf_filename)

                # Send to common recipient
                email_thread_common = threading.Thread(target=send_email_with_attachment, args=("CommonRecipent@Dynatecsmv.no", pdf_filename, f"Pallet Labels for {name}"))
                email_thread_common.start()

                # Send to individual recipient
                email_thread_individual = threading.Thread(target=send_email_with_attachment, args=(mail_address, pdf_filename, f"Pallet Labels for {name}"))
                email_thread_individual.start()

            # Delete processed rows
            delete_response = requests.delete(glide_api_url, headers=headers, json={"rowIDs": all_row_ids})
            if delete_response.status_code != 200:
                print(f"Failed to delete rows: {delete_response.status_code} - {delete_response.text}")

            return jsonify({"message": "PDFs created, emails sent, and rows deleted successfully"}), 200
        else:
            return jsonify({"error": "Failed to fetch data from Glide API", "status_code": response.status_code}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

def send_email_with_attachment(to_email, pdf_filename, subject):
    try:
        with app.app_context():
            msg = Message(subject, recipients=[to_email])
            msg.body = "Please find the attached pallet label PDF."
            with app.open_resource(pdf_filename) as pdf:
                msg.attach(pdf_filename, "application/pdf", pdf.read())
            mail.send(msg)
    except Exception as e:
        print(f"Error sending email: {str(e)}")

if __name__ == '__main__':
    app.run(debug=True)
