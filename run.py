# run.py

import io
from flask import Flask, render_template, request, jsonify
from app.email_sender import send_emails, get_contacts_from_excel
import os
from dotenv import load_dotenv
from werkzeug.utils import secure_filename
import traceback
import pandas as pd

load_dotenv()

app = Flask(
    __name__,
    template_folder='app/templates',
    static_folder='app/static'  # <-- вот оно, ключевое
)
app.config['UPLOAD_FOLDER'] = 'app/data'

MY_ADDRESS = os.getenv("MY_ADDRESS")
PASSWORD = os.getenv("PASSWORD")


@app.route('/')
def index():
    with open('app/email_templates/message.txt', 'r', encoding='utf-8') as f:
        default_template = f.read()
    return render_template('index.html', default_template=default_template)


@app.route('/preview-excel', methods=['POST'])
def preview_excel():
    file = request.files.get('contacts_file')
    if not file:
        return "❌ Файл не загружен.", 400

    try:
        df = pd.read_excel(io.BytesIO(file.read()))

        df['name'] = df['name'].fillna('').astype(str)
        df.loc[df['name'].str.strip() == '', 'name'] = 'Коллеги'

        add_prefix = request.form.get('add_tc_prefix', 'true').lower() == 'true'

        def fix_sm(sm_value):
            if not isinstance(sm_value, str):
                return sm_value
            sm_value_strip = sm_value.strip()
            prefixes = ("ТЦ", "ТРЦ", "ТРК", "ТД", "ТК")
            if any(sm_value_strip.startswith(prefix) for prefix in prefixes):
                return sm_value_strip
            return "ТЦ " + sm_value_strip if add_prefix else sm_value_strip

        df['sm'] = df['sm'].apply(fix_sm)

        first_row = df.iloc[0].to_dict() if not df.empty else {}
        attrs = f'data-sm="{first_row.get("sm", "")}" data-city="{first_row.get("city", "")}"' if first_row else ""
        table_html = df.to_html(classes="preview-table", index=False)
        return f'<div id="first-row-data" {attrs} style="display:none;"></div>' + table_html
    except Exception as e:
        return f"<div style='color:red;'>❌ Ошибка при чтении файла: {str(e)}</div>"


@app.route('/send-emails', methods=['POST'])
def send():
    brand = request.form['brand']
    period = request.form['period']
    cc_input = request.form.get('cc_list', '')
    cc_addresses = [email.strip() for email in cc_input.split(',') if email.strip()]
    uploaded_file = request.files['contacts_file']

    if uploaded_file.filename == '':
        return render_template("status.html", status="❌ No file uploaded.")

    filename = secure_filename(uploaded_file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    uploaded_file.save(file_path)
    template_text = request.form['message_template']

    try:
        contacts = get_contacts_from_excel(file_path)
        send_emails(
            my_address=MY_ADDRESS,
            password=PASSWORD,
            contacts=contacts,
            cc_addresses=cc_addresses,
            brand=brand,
            period=period,
            template_text=template_text
        )
        status = f"✅ Emails sent successfully to {len(contacts)} contacts."
    except Exception as e:
        error_trace = traceback.format_exc()
        status = f"❌ Error occurred: {str(e)}\n\n<pre>{error_trace}</pre>"
    return render_template("status.html", status=status)


if __name__ == '__main__':
    app.run(debug=True)
