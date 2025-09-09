import io
import os
import json
import pytest
import pandas as pd
from email.message import Message

from app.email_sender import get_contacts_from_excel, send_emails


def _write_xlsx(tmp_path, rows, name="contacts.xlsx"):
    df = pd.DataFrame(rows)
    path = tmp_path / name
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return path


def test_get_contacts_from_excel_splits_and_groups(tmp_path):
    # Arrange
    path = _write_xlsx(
        tmp_path,
        [
            {
                "email": "to@example.com, cc1@example.com",
                "name": "",
                "mall": "Афимолл",
                "city": "Москва",
                "rim": "111",
            },
            {
                "email": "to@example.com; cc2@example.com и cc3@example.com",
                "name": "",
                "mall": "Афимолл",
                "city": "Москва",
                "rim": "222",
            },
        ],
    )

    # Act
    contacts = get_contacts_from_excel(path)

    # Assert
    assert isinstance(contacts, list) and len(contacts) == 1
    c = contacts[0]
    assert c["email"] == "to@example.com"
    assert c["name"] == "Коллеги"  # Defaulted
    assert c["city"] == "Москва" and c["mall"] == "Афимолл"
    assert c["rim"] == "111\n222"  # Grouped with newline
    assert set(c["_cc_emails"]) == {"cc1@example.com", "cc2@example.com", "cc3@example.com"}


def test_get_contacts_from_excel_template_validation_missing_doc(tmp_path):
    # Arrange: есть rim, но требуется DOC (его нет)
    path = _write_xlsx(
        tmp_path,
        [{"email": "a@b.com", "mall": "ТЦ Мега", "city": "СПб", "rim": "R1"}],
    )
    template_text = "Привет ${NAME}. RIM: ${RIM}. Ссылка: ${DOC}"

    # Act / Assert
    with pytest.raises(ValueError) as ei:
        get_contacts_from_excel(path, template_text=template_text, doc=None)
    assert "Нет необходимого поля doc" in str(ei.value)


def test_get_contacts_from_excel_template_validation_missing_link_column(tmp_path):
    # Arrange: отсутствует колонка link, но плейсхолдер ${LINK} присутствует
    path = _write_xlsx(
        tmp_path,
        [{"email": "a@b.com", "mall": "ТЦ Мега", "city": "СПб", "rim": "R1"}],
    )
    template_text = "Инфо: ${RIM} / ${LINK}"

    # Act / Assert
    with pytest.raises(ValueError) as ei:
        get_contacts_from_excel(path, template_text=template_text, doc="https://doc")
    msg = str(ei.value)
    assert "Нет необходимого столбца(ов)" in msg and "link" in msg


def test_send_emails_builds_message_and_combines_cc(monkeypatch):
    # Arrange
    sent = {}

    class DummySMTP:
        instances = []

        def __init__(self, host, port, context=None):
            self.host = host
            self.port = port
            self.context = context
            self.logged_in = None
            self.sent_messages = []
            DummySMTP.instances.append(self)

        def set_debuglevel(self, level):  # noqa: D401
            self.debuglevel = level

        def ehlo(self):  # noqa: D401
            self.did_ehlo = True

        def login(self, user, pwd):
            self.logged_in = (user, pwd)

        def send_message(self, msg: Message, from_addr: str, to_addrs):
            self.sent_messages.append((msg, from_addr, list(to_addrs)))

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

    # Patch SMTP_SSL in our module
    monkeypatch.setattr("app.email_sender.smtplib.SMTP_SSL", DummySMTP)

    my_address = "me@example.com"
    password = "secret"
    contacts = [
        {
            "email": "to@example.com",
            "name": "Иван",
            "mall": '"Афимолл"',  # quotes should be removed in subject/body
            "city": "Москва",
            "rim": "R",
            "_cc_emails": ["cc2@example.com", "cc3@example.com", "cc2@example.com"],
        }
    ]
    cc_addresses = ["cc1@example.com", "cc2@example.com"]  # cc2 duplicated
    brand = "BrandX"
    period = "01"
    doc = "https://example.com/doc"
    template_text = "Hello ${NAME} at ${MALL}. ${DOC}"
    display_name = "Sender"

    # Act
    send_emails(my_address, password, contacts, cc_addresses, brand, period, doc, template_text, display_name)

    # Assert
    assert DummySMTP.instances, "SMTP_SSL was not instantiated"
    smtp = DummySMTP.instances[-1]
    assert smtp.host == "smtp.yandex.ru" and smtp.port == 465
    assert smtp.logged_in == (my_address, password)
    assert len(smtp.sent_messages) == 1

    msg, from_addr, to_addrs = smtp.sent_messages[0]
    assert from_addr == my_address
    # Primary + unique CCs
    assert set(to_addrs) == {"to@example.com", "cc1@example.com", "cc2@example.com", "cc3@example.com"}

    # Headers
    assert msg["From"] is not None
    assert msg["To"] == "to@example.com"
    assert set(map(str.strip, msg.get("Cc", "").split(","))) == {"cc1@example.com", "cc2@example.com", "cc3@example.com"}
    assert msg["Subject"] == "Афимолл (г. Москва) // BrandX // 01"

    # Body (first part is text/plain)
    payload = msg.get_payload()
    assert payload, "Multipart payload is empty"
    body = payload[0].get_payload(decode=True).decode("utf-8")
    assert "Hello Иван" in body
    assert "Афимолл" in body and '"Афимолл"' not in body  # quotes removed
    assert doc in body