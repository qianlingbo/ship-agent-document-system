---
name: ship-passport-mrz
description: OCR numbered passport PDFs, match them to crew rows, and normalize non-Chinese crew names to passport MRZ order. Use when Crew List names must be verified, especially for Egyptian, Philippine, and Indonesian passports.
---

# Ship Passport MRZ

Run with crew JSON from `$ship-crew-list` or an original Crew List:

```bash
python3 "{baseDir}/scripts/run.py" --crew "/path/crew.json" --passports "/path/TRAVEL PASSPORT" --output "/path/crew-mrz.json"
```

Passport filenames must start with crew sequence numbers. MRZ order overrides printed Full Name order. For example, `P<EGYKHEDR<<IHAB<ELSAYED<KHEDR` becomes `KHEDR IHAB ELSAYED KHEDR`.

Exit code `2` or a non-empty `manual_review_sequences` means human review is required. Record confirmed corrections in a JSON object keyed by passport number and rerun with `--overrides`. Never silently accept an unverified name.
