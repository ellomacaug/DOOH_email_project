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


def get_contacts_from_excel(filepath, template_text=None, doc=None):
    # Read and keep only relevant columns (email, name, mall, city, rim)
    df = pd.read_excel(filepath)
    cols = [c for c in ['email', 'name', 'mall', 'city', 'rim'] if c in df.columns]
    if 'email' not in cols:
        raise ValueError("❌ Нет обязательного столбца: email")
    
    df = df[cols].fillna('').astype(str).apply(lambda x: x.str.strip())
    
    # Split multi-email cells into primary + cc list, preserve other columns
    def split_emails(email_str):
        s = str(email_str)
        # normalize separators to spaces
        for sep in [',', ';', '/', '|', ' и ']:
            s = s.replace(sep, ' ')
        parts = [p.strip() for p in s.split() if p.strip()]
        return parts
    
    email_regex = r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$'
    contacts = []
    for idx, row in df.iterrows():
        parts = split_emails(row['email'])
        if not parts:
            raise ValueError(f"❌ Строка {idx + 2} не содержит email. Удалите её или заполните.")
        # validate all found emails
        for e in parts:
            if not re.match(email_regex, e):
                raise ValueError(f"❌ Неверный формат email: {e} в строке {idx + 2}")
        primary = parts[0]
        cc = parts[1:]
        contact = {
            'email': primary,
            'name': row.get('name', ''),
            'mall': row.get('mall', ''),
            'city': row.get('city', ''),
            'rim': row.get('rim', ''),
            '_cc_emails': cc
        }
        contacts.append(contact)
    
    df = pd.DataFrame(contacts)
    # Normalize text fields
    for col in ['email', 'name', 'mall', 'city', 'rim']:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str).str.strip()
    
    if 'name' in df.columns:
        df.loc[df['name'] == '', 'name'] = 'Коллеги'
    
    # Group by city,mall,email,name and concatenate rim with newline
    if 'rim' in df.columns:
        agg_map = {'rim': lambda x: '\n'.join(filter(lambda v: v != '', map(str, x)))}
        # preserve aggregated cc emails (unique)
        if '_cc_emails' in df.columns:
            agg_map['_cc_emails'] = lambda lists: list({email for lst in lists for email in (lst if isinstance(lst, list) else [lst]) if email})
        df = df.groupby(['city', 'mall', 'email', 'name'], as_index=False).agg(agg_map)

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

            msg['Subject'] = f"{mall_name} (г. {contact['city']}) // {brand} // {period}"

            msg.attach(MIMEText(message, 'plain'))
            recipients = [contact['email']] + all_cc
            server.send_message(msg, from_addr=my_address, to_addrs=recipients)
            del msg

