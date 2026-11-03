# SimpoSheet

SimpoSheet converts text-based PDF bank statements, receipts, and bills into clean, formatted Excel workbooks.

Instead of producing a raw text dump, SimpoSheet creates workbooks with real Excel dates and currency values, structured worksheets, and reconciliation checks that show whether the extracted totals agree with the source document.

**Created by Gifty Arthur Danquah — Founder & Builder**

## What it does

### Bank statements
- Extracts dated transactions into a formatted table
- Separates debits and credits
- Preserves running balances when present
- Adds a Summary sheet with a reconciliation check

### Receipts
- Creates one clean expense record
- Extracts the date, vendor, reference, amount, and payment method
- Keeps only the card's last four digits when printed on the receipt

### Bills
- Extracts the vendor, bill date, due date, previous balance, current charges, and total due
- Lists supported per-line charges
- Checks line items against the bill's current-charge total
- Checks previous balance + current charges against total due

## Privacy approach

- Uploads are processed in a temporary folder.
- The uploaded PDF and generated workbook are deleted when the download response closes.
- There is no user account or document database in this version.
- Document details are kept out of application logs.
- No real customer documents are included in this repository.

## Tech stack

- Python
- Flask
- Gunicorn
- pdfplumber
- OpenPyXL
- HTML, CSS, and vanilla JavaScript

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open http://localhost:5000.

A fictional sample statement is included at `samples/sample_statement.pdf` so the converter can be tested without using a real financial document.

## Run the smoke test

```bash
pip install -r requirements.txt
python -m pytest tests/smoke_test.py
```

The test checks the home page, health endpoint, and a real conversion of the fictional sample statement.

## Deploy

### Render

This repository includes `render.yaml` and a `Procfile`.

1. Push the repository to GitHub.
2. In Render, choose **New > Blueprint** and select the repository.
3. Confirm the `simposheet` web service. The included configuration starts on Render's Free plan for a public beta launch.
4. Deploy and test the generated `onrender.com` address.
5. Upgrade to a paid always-on plan before marketing it as a production service, then add the custom domain.

### Other Python hosts

Any host that can run Gunicorn can use:

```bash
gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
```

## Current limitations

- Text-based PDFs only; scanned-image PDFs require OCR.
- Statement and bill layouts vary by provider, so new formats may need parser tuning.
- No accounts, subscriptions, saved conversion history, or payment processing in this version.
- Currency output is currently formatted for U.S. dollars.

## Roadmap

- OCR for scanned statements and photographed receipts
- More bank and bill-provider layouts
- User accounts and subscription plans
- International date and currency preferences
- Batch uploads
- QuickBooks-friendly CSV exports

## Portfolio note

SimpoSheet began as a product idea from founder Gifty Arthur Danquah: financial documents should become workbooks people can use immediately — sorted, filtered, totaled, and verified — rather than files that still need hours of cleanup.
