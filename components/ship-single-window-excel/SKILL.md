---
name: ship-single-window-excel
description: Assemble the China MSA single-window Excel workbook from crew rows, port rows, and certificate JSON while preserving the official template. Use after the relevant component conversions; not for website submission.
---

# Ship Single Window Excel

Run with any available component outputs:

```bash
python3 "{baseDir}/scripts/run.py" --crew "/path/crew-mrz.json" --ports "/path/ports.json" --certificates "/path/certificates.json" --output "/path/单一窗口录入.xlsx"
```

Each input is optional, but report which sections were populated. WRC must not be written. Preserve bundled sheet names and template structure. Highlight or report unresolved nationality/port fields and blank certificate dates. This Skill prepares a local import file only; government-site upload requires a separate explicit request and final-data approval.
