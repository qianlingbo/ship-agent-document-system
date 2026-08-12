---
name: ship-crew-info-export
description: Generate a crew-information Excel workbook from a crew XLSX/CSV, numbered face photos, numbered passport PDFs, and an existing Excel template. Use for ship-agent crew entry when passport pages must be auto-rotated, cropped to the biometric data area, embedded with face photos, constrained to 20–200 KB per image, and validated while contact, vessel, permit-photo, permit-period, permit-type, and crew-change fields remain blank.
---

# Ship Crew Info Export

Use the deterministic script. Do not manually transcribe or crop images unless the script reports an error.

If the runtime is unknown, run `python3 "{baseDir}/scripts/check_runtime.py"` first. Read the installation reference only when it reports missing dependencies.

## Inputs

Require four paths:

- Existing target `.xlsx` template.
- Crew `.xlsx` or `.csv` containing sequence, name, gender, nationality, and passport number.
- Face-photo directory with filenames starting `01`, `02`, `03`... in crew order.
- Passport-PDF directory with filenames starting with the same sequence numbers.

Names in the crew source must already use passport MRZ order. Do not replace them with the passport's printed Full Name. Use Chinese nationality names where the target system expects Chinese.

## Run

Run from any directory. Replace `{baseDir}` with this skill directory if the host does not expand it.

```bash
python3 "{baseDir}/scripts/export_crew_info.py" \
  --template "/path/船员信息模板.xlsx" \
  --crew "/path/单一窗口录入.xlsx" \
  --photos "/path/CREW PHOTO" \
  --passports "/path/TRAVEL PASSPORT" \
  --output "/path/船员信息（已录入）.xlsx"
```

Add `--crew-sheet NAME` or `--target-sheet NAME` only when automatic sheet detection fails.

## Interpret result

Treat exit code `0` and JSON `"status": "success"` as complete. Report:

- Output path.
- Crew count and embedded image count.
- Minimum and maximum embedded image KB.
- Confirmation that required fields remain blank.

On failure, report the exact JSON error. Do not claim completion and do not silently skip a crew member.

## Fixed rules

- Keep the original template unchanged; write a new output file.
- Map document type to numeric code `2` unless the user specifies another value.
- Leave 联系方式、船名、许可照片、许可开始、许可截止、许可类型、换班类型 blank.
- Embed one face photo and one passport data-region image per crew member.
- Preserve portrait, identity fields, expiry information, and MRZ; remove passport cover and large blank scan margins.
- Require every embedded image to be 20–200 KB.
- Require continuous sequence numbers starting at `1` across crew, photos, and passports.
- Stop on duplicate, missing, or extra numbered files.

Read [references/INSTALL.md](references/INSTALL.md) only when a dependency is missing or the skill must be installed on another machine.
