---
name: ship-crew-list
description: Convert an IMO Crew List XLSX or CSV into structured China MSA crew rows with duty codes, document types, compact dates, mapped embarkation ports, and Chinese nationality as birthplace. Use before MRZ verification or single-window assembly.
---

# Ship Crew List

Run:

```bash
python3 "{baseDir}/scripts/run.py" "/path/IMO CREW LIST.xlsx" --output "/path/crew.json"
```

Optional defaults: `--embark-date YYYY-MM-DD --embark-port PORT`. Preserve the crew sequence. Birthplace is the Chinese nationality name, not the passport city. This step does not prove passport name order; use `$ship-passport-mrz` for that. Report unmapped nationality, duty, or embarkation-port fields.
