import io
import json
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
    privacy = client.get("/privacy")
    assert privacy.status_code == 200
    assert b"Privacy Policy" in privacy.data


def test_pwa_manifest_and_service_worker():
    client = app.test_client()
    manifest = client.get("/manifest.webmanifest")
    assert manifest.status_code == 200
    data = json.loads(manifest.data)
    assert data["name"] == "SimpoSheet"
    assert data["display"] == "standalone"
    assert {icon["sizes"] for icon in data["icons"]} >= {"192x192", "512x512"}
    sw = client.get("/sw.js")
    assert sw.status_code == 200
    assert b"simposheet" in sw.data.lower()


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
    # Branded layout: rows 1-5 header block, row 6 column headers,
    # data rows 7-18 (12 sample transactions).
    assert workbook["Transactions"].max_row == 18
    assert workbook["Transactions"]["B1"].value == "SimpoSheet"

    summary = {
        row[0]: row[1]
        for row in workbook["Summary"].iter_rows(min_row=7, values_only=True)
        if row[0]
    }
    assert summary["Reconciliation"] == "MATCH ✓"
