---
name: ship-crew-photo-workbook
description: Generate the crew information Excel workbook from single-window crew rows, numbered face photos, and numbered passport PDFs, including biometric-page cropping and 20–200 KB image validation. Use only for the photo workbook.
---

# Ship Crew Photo Workbook

Run:

```bash
python3 "{baseDir}/scripts/run.py" --crew "/path/单一窗口录入.xlsx" --photos "/path/CREW PHOTO" --passports "/path/TRAVEL PASSPORT" --output "/path/船员信息.xlsx"
```

Crew, face photos, and passports must use the same continuous sequence. Crop the passport to the data region containing portrait, identity fields, expiry, and MRZ. Require two embedded images per person, each 20–200 KB. Leave 联系方式、船名、许可照片、许可开始/截止、许可类型、换班类型 blank. Stop on missing, duplicate, or extra sequence files.
