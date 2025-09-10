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


def _excel_file_from_rows(rows):
    # Helper: build in-memory Excel file
    df = pd.DataFrame(rows)
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    bio.seek(0)
    return bio


def test_preview_excel_no_file_returns_400(client):
    # Arrange
    # Act
    resp = client.post("/preview-excel")
    # Assert
    assert resp.status_code == 400
    assert "Файл не загружен" in resp.get_data(as_text=True)


def test_preview_excel_missing_required_columns_returns_400(client):
    # Arrange: only email column, missing mall and city
    bio = _excel_file_from_rows([{"email": "a@b.com"}])

    # Act
    resp = client.post(
        "/preview-excel",
        data={"contacts_file": (bio, "contacts.xlsx")},
        content_type="multipart/form-data",
    )

    # Assert
    text = resp.get_data(as_text=True)
    assert resp.status_code == 400
    assert "обязательные столбцы" in text


def test_preview_excel_empty_email_rows_returns_400(client):
    # Arrange: empty email value after normalization
    bio = _excel_file_from_rows(
        [{"email": "", "mall": "Афимолл", "city": "Москва"}]
    )

    # Act
    resp = client.post(
        "/preview-excel",
        data={"contacts_file": (bio, "contacts.xlsx")},
        content_type="multipart/form-data",
    )

    # Assert
    text = resp.get_data(as_text=True)
    assert resp.status_code == 400
    assert "строки без email" in text


def test_preview_excel_adds_tc_prefix(client):
    # Arrange
    rows = [
        {"email": "x@y.com", "mall": "Афимолл", "city": "Москва", "name": "", "rim": "111"}
    ]
    bio = _excel_file_from_rows(rows)

    # Act
    resp = client.post(
        "/preview-excel",
        data={"contacts_file": (bio, "contacts.xlsx"), "add_tc_prefix": "true"},
        content_type="multipart/form-data",
    )

    # Assert
    text = resp.get_data(as_text=True)
    assert resp.status_code == 200
    assert "ТЦ Афимолл" in text  # prefix added


def test_preview_excel_defaults_name_to_kollegi(client):
    # Arrange
    rows = [
        {"email": "x@y.com", "mall": "Афимолл", "city": "Москва", "name": "", "rim": "111"}
    ]
    bio = _excel_file_from_rows(rows)

    # Act
    resp = client.post(
        "/preview-excel",
        data={"contacts_file": (bio, "contacts.xlsx"), "add_tc_prefix": "true"},
        content_type="multipart/form-data",
    )

    # Assert
    text = resp.get_data(as_text=True)
    assert resp.status_code == 200
    assert "Коллеги" in text
    

def test_preview_excel_renders_hidden_first_row_attrs(client):
    # Arrange
    rows = [
        {"email": "x@y.com", "mall": "Афимолл", "city": "Москва", "name": "", "rim": "111"}
    ]
    bio = _excel_file_from_rows(rows)

    # Act
    resp = client.post(
        "/preview-excel",
        data={"contacts_file": (bio, "contacts.xlsx"), "add_tc_prefix": "true"},
        content_type="multipart/form-data",
    )

    # Assert
    text = resp.get_data(as_text=True)
    assert resp.status_code == 200
    assert 'id="first-row-data"' in text
    assert 'data-mall="ТЦ Афимолл"' in text
    assert 'data-city="Москва"' in text


def test_preview_excel_no_prefix_when_disabled(client):
    # Arrange
    rows = [{"email": "x@y.com", "mall": "Афимолл", "city": "Москва"}]
    bio = _excel_file_from_rows(rows)

    # Act
    resp = client.post(
        "/preview-excel",
        data={
            "contacts_file": (bio, "contacts.xlsx"),
            "add_tc_prefix": "false",
        },
        content_type="multipart/form-data",
    )

    # Assert
    text = resp.get_data(as_text=True)
    assert resp.status_code == 200
    assert "ТЦ Афимолл" not in text  # no auto-prefix
    assert ">Афимолл<" in text or 'data-mall="Афимолл"' in text


def test_preview_excel_groups_rim_with_br(client):
    # Arrange: two rows same (city, mall, email, name) with different rim
    rows = [
        {"email": "a@b.com", "name": "Иван", "city": "Москва", "mall": "ТРЦ Мега", "rim": "123"},
        {"email": "a@b.com", "name": "Иван", "city": "Москва", "mall": "ТРЦ Мега", "rim": "456"},
    ]
    bio = _excel_file_from_rows(rows)

    # Act
    resp = client.post(
        "/preview-excel",
        data={"contacts_file": (bio, "contacts.xlsx"), "add_tc_prefix": "true"},
        content_type="multipart/form-data",
    )

    # Assert
    text = resp.get_data(as_text=True)
    assert resp.status_code == 200
    assert "123<br>456" in text  # newline converted to <br>
    # Existing prefixes should not be duplicated
    assert "ТРЦ Мега" in text