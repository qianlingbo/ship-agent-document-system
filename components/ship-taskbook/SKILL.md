---
name: ship-taskbook
description: Fill the vessel plan taskbook DOCX from ship particulars, an archived email body, and optional confirmed fields. Use for generating or refreshing the taskbook only.
---

# Ship Taskbook

Run:

```bash
python3 "{baseDir}/scripts/run.py" --ship-name "VESSEL" --particulars "/path/SHIP PARTICULARS.pdf" --mail-body "/path/mail.txt" --output "/path/船舶计划任务书.docx"
```

Use `--fields confirmed-fields.json` for verified overrides and `--completed 海事 危险品 单证` only for work actually completed. Fill only facts present in the supplied sources; leave unknown values blank and never reuse values from a different vessel.
