---
name: ship-certificate-pdf
description: Classify and merge vessel certificate PDFs into ordered vessel, flag competency, and dangerous-goods PDF groups. Use when PDF bundles are needed; do not use for extracting certificate dates or creating Excel rows.
---

# Ship Certificate PDF

Run on a directory containing extracted vessel documents:

```bash
python3 "{baseDir}/scripts/run.py" "/path/documents" --output "/path/vessel/海事"
```

Required outputs, when source files exist, are `船舶证书.pdf`, `适任证书.pdf`, and four dangerous-goods groups under `危险品证书/`. Vessel certificates follow the prescribed 01–13 order, with WRC last. IOPP accepts Form A or Form B. WRC is retained in the merged vessel PDF but remains excluded from single-window Excel by the Excel component. Verify returned paths and `证书登记.json`; do not claim a missing category was generated.
