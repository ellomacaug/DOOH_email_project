# app/email_sender.py

import smtplib
import ssl
import re
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from string import Template
from pathlib import Path


import re


def get_contacts_from_excel(filepath, template_text=None, doc=None):
    df = pd.read_excel(filepath)

    if 'email' not in df.columns:
        raise ValueError("❌ Нет обязательного столбца: email")
        
    # Process email column to handle multiple emails
    contacts = []
    for idx, row in df.iterrows():
        email_str = str(row.get('email', '')).strip()
        if not email_str:
            raise ValueError(f"❌ Строка {idx + 2} не содержит email. Удалите её или заполните.")
            
        # Split by common separators and clean up
        emails = []
        for separator in [',', ';', ' и ', ' ', 'и', 'и ', ' и ', '  ']:
            email_str = email_str.replace(separator, ' ')
        emails = [e.strip() for e in email_str.split() if e.strip()]
        
        if not emails:
            raise ValueError(f"❌ Неверный формат email в строке {idx + 2}")
            
        # Validate email format
        email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        for email in emails:
            if not re.match(email_regex, email):
                raise ValueError(f"❌ Неверный формат email: {email} в строке {idx + 2}")
        
        # Create contact with first email as primary and rest as cc
        contact = row.to_dict()
        contact['email'] = emails[0]
        contact['_cc_emails'] = emails[1:]  # Store additional emails as CC
        contacts.append(contact)
    
    # Convert back to DataFrame for further processing
    df = pd.DataFrame(contacts)
    
    # Fill empty values and process other columns
    for col in df.columns:
        if col != '_cc_emails':  # Skip the temporary _cc_emails column
            df[col] = df[col].fillna('').astype(str).str.strip()

    if 'name' in df.columns:
        df.loc[df['name'] == '', 'name'] = 'Коллеги'

    if 'mall' in df.columns:
        df['mall'] = df['mall'].fillna('').astype(str).str.strip()

    if template_text:
        required_map = {
            "RIM": "rim",
            "LINK": "link",
            "MIN": "min",
            "SEC": "sec"
        }
        placeholders = set(re.findall(r"\$\{(\w+)\}", template_text))

        needed_columns = [required_map[p] for p in placeholders if p in required_map]
        missing_cols = [col for col in needed_columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"❌ Нет необходимого столбца(ов): {', '.join(missing_cols)}")

        empty_required = []
        for ph, col in required_map.items():
            if ph in placeholders and col in df.columns:
                if df[col].fillna('').astype(str).str.strip().eq('').any():
                    empty_required.append(col)
        if empty_required:
            raise ValueError(
                "❌ В обязательных столбцах есть пустые значения: "
                + ", ".join(empty_required)
                + ". Заполните их или удалите строки."
            )

        if "DOC" in placeholders and not (doc and str(doc).strip()):
            raise ValueError("❌ Нет необходимого поля doc (ссылка)")

    return df.to_dict(orient='records')


def read_template(template_path):
    with open(template_path, 'r', encoding='utf-8') as file:
        return Template(file.read())


def send_emails(my_address, password, contacts, cc_addresses, brand, period, doc, template_text, display_name):
    template = Template(template_text)
    context = ssl.create_default_context()
    cc_addresses = cc_addresses or []

    with smtplib.SMTP_SSL('smtp.yandex.ru', 465, context=context) as server:
        server.set_debuglevel(1)
        server.ehlo()
        server.login(my_address, password)

        for contact in contacts:
            msg = MIMEMultipart()

            # Clean up mall name by removing any existing double quotes for both subject and body
            mall_name = contact['mall'].replace('"', '')
            
            message = template.safe_substitute(
                NAME=contact['name'],
                BRAND=brand,
                PERIOD=period,
                MALL=mall_name,
                RIM=contact.get('rim', ''),
                LINK=contact.get('link', ''),
                MIN=contact.get('min', ''),
                SEC=contact.get('sec', ''),
                DOC=doc or ""
            )

            # Get CC emails from both the contact and the additional CC addresses
            contact_cc = contact.get('_cc_emails', [])
            all_cc = list(set(cc_addresses + contact_cc))  # Remove duplicates
            
            msg['From'] = formataddr((display_name, my_address))
            msg['To'] = contact['email']
            if all_cc:
                msg['Cc'] = ", ".join(all_cc)
            
            # Clean up mall name by removing any existing double quotes
            mall_name = contact['mall'].replace('"', '')
            msg['Subject'] = f"{mall_name} (г. {contact['city']}) // {brand} // {period}"

            msg.attach(MIMEText(message, 'plain'))
            recipients = [contact['email']] + all_cc
            server.send_message(msg, from_addr=my_address, to_addrs=recipients)
            del msg

