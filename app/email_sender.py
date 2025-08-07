# app/email_sender.py

import smtplib
import ssl
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template
from pathlib import Path

def get_contacts_from_excel(filepath):
    df = pd.read_excel(filepath)

    return df.to_dict(orient='records')

def read_template(template_path):
    with open(template_path, 'r', encoding='utf-8') as file:
        return Template(file.read())

def send_emails(my_address, password, contacts, cc_addresses, brand, period, template_text):
    template = Template(template_text)
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.yandex.ru', 465, context=context) as server:
        server.set_debuglevel(1)
        server.ehlo()
        server.login(my_address, password)

        for contact in contacts:
            msg = MIMEMultipart()
            name = contact.get('name')
            safe_name = name.title() if name and isinstance(name, str) and name.strip() else "Коллеги"

            text = contact.get('text', '')
            message = template.safe_substitute(
                NAME=safe_name,
                BRAND=brand,
                PERIOD=period,
                TEXT=text or ''
            )

            msg['From'] = my_address
            msg['To'] = contact['email']
            msg['Cc'] = ", ".join(cc_addresses)
            msg['Subject'] = f"{contact['sm']} (г. {contact['city']}) // {brand} // {period}"

            msg.attach(MIMEText(message, 'plain'))
            recipients = [contact['email']] + cc_addresses + [my_address]
            server.send_message(msg, from_addr=my_address, to_addrs=recipients)
            del msg
