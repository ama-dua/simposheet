import io
from pathlib import Path

from openpyxl import load_workbook

from app import app


BASE_DIR = Path(__file__).resolve().parents[1]
SAMPLE = BASE_DIR / "samples" / "sample_statement.pdf"


def test_home_and_health():
    client = app.test_client()
    assert client.get("/").status_code == 200
    health = client.get("/health")
    assert health.status_code == 200
    assert health.get_json() == {"service": "SimpoSheet", "status": "ok"}


def test_sample_statement_conversion():
    client = app.test_client()
    with SAMPLE.open("rb") as pdf:
        response = client.post(
            "/convert",
            data={"document": (pdf, "sample_statement.pdf")},
            content_type="multipart/form-data",
        )

    assert response.status_code == 200
    workbook = load_workbook(io.BytesIO(response.data))
    assert workbook.sheetnames == ["Transactions", "Summary"]
    assert workbook["Transactions"].max_row == 13

    summary = {
        row[0]: row[1]
        for row in workbook["Summary"].iter_rows(min_row=3, values_only=True)
        if row[0]
    }
    assert summary["Reconciliation"] == "MATCH ✓"
