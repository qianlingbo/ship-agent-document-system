---
name: ship-port-of-call
description: Convert a vessel Port of Call or last-ten-ports XLSX, CSV, or table PDF into structured China MSA port rows with country and UN/LOCODE mappings. Use for port-history conversion only.
---

# Ship Port of Call

Run:

```bash
python3 "{baseDir}/scripts/run.py" "/path/last 10 ports.xlsx" --output "/path/ports.json"
```

Keep source order. Review blank country or port-code mappings in the JSON before import. Do not invent missing arrival/departure dates. This Skill does not create the final Excel workbook; pass the JSON to `$ship-single-window-excel`.
