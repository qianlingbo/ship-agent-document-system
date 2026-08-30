---
name: ship-document-organizer
description: Organize local vessel documents or an EML into a ship-named folder, preserve originals, and recursively extract ZIP, RAR, TAR, GZ, TGZ, BZ2, and XZ archives. Use for file intake and archive normalization only.
---

# Ship Document Organizer

Run:

```bash
python3 "{baseDir}/scripts/run.py" "/path/input-or-message.eml" --output "/path/work" --ship-name "VESSEL"
```

Preserve originals and directory structure. Do not execute attachments. The extractor rejects path traversal and unsafe TAR links. Verify `整理清单.json`; report every entry in `errors`. This Skill does not classify certificates or generate office files.
