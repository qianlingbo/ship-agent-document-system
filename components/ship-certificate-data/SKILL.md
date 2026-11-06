---
name: ship-certificate-data
description: Extract vessel certificate categories, numbers, issue dates, expiry dates, and latest annual verification dates into structured JSON. Use for certificate data entry preparation, not for PDF merging.
---

# Ship Certificate Data

Run:

```bash
python3 "{baseDir}/scripts/run.py" "/path/documents" --output "/path/certificates.json"
```

The extractor prefers a vessel-particulars certificate table when present, then falls back to PDF text/OCR. WRC must not appear in this Excel-entry data because it is PDF-only. Treat blank or OCR-noted fields as unresolved and report them for human confirmation; never infer a date or certificate number.
