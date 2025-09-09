import io
import pytest
import pandas as pd

from run import app as flask_app


@pytest.fixture()
def client():
    # Arrange (shared): configure Flask test client
    flask_app.config["TESTING"] = True
    with flask_app.test_client() as c:
        yield c


def test_send_emails_requires_session(client):
    # Arrange: no session credentials
    form = {
        "brand": "X",
        "period": "01.01.2025 - 31.01.2025",
        "doc": "",
        "message_template": "Hi",
    }

    # Act
    resp = client.post("/send-emails", data=form, content_type="multipart/form-data")

    # Assert
    assert resp.status_code == 401
    assert "Сессия истекла" in resp.get_data(as_text=True)


def test_send_emails_with_session_but_no_file_returns_error(client):
    # Arrange: put credentials into session
    with client.session_transaction() as sess:
        sess["DISPLAY_NAME"] = "User"
        sess["MY_ADDRESS"] = "user@example.com"
        sess["PASSWORD"] = "secret"

    form = {
        "brand": "X",
        "period": "01.01.2025 - 31.01.2025",
        "doc": "",
        "message_template": "Hi",
    }

    # Act
    resp = client.post("/send-emails", data=form, content_type="multipart/form-data")

    # Assert
    text = resp.get_data(as_text=True)
    assert resp.status_code == 200
    assert "Файл не загружен" in text