# Component Skill routing

Use the total workflow only when the user asks to process a vessel end to end. For a narrow request, invoke exactly one component Skill:

| Skill | Responsibility | Main output |
|---|---|---|
| `$ship-mail-intake` | List/read Agent Mail and download one message | Vessel raw folder |
| `$ship-document-organizer` | Preserve, group, and recursively extract local files | `整理清单.json` |
| `$ship-certificate-pdf` | Merge vessel, competency, and dangerous-goods PDFs | `海事/*.pdf` |
| `$ship-certificate-data` | Extract certificate number and dates | `certificates.json` |
| `$ship-crew-list` | Convert IMO Crew List | `crew.json` |
| `$ship-port-of-call` | Convert Port of Call / last ten ports | `ports.json` |
| `$ship-passport-mrz` | Verify name order against passport MRZ | `crew-mrz.json` |
| `$ship-crew-photo-workbook` | Crop/compress photos and build crew workbook | `船员信息.xlsx` |
| `$ship-taskbook` | Fill vessel plan taskbook | `船舶计划任务书.docx` |
| `$ship-single-window-excel` | Assemble crew, port, and certificate sections | `单一窗口录入.xlsx` |

The component Skills share the deterministic runtime in `$ship-agent-document-system`; installing the bundle places them as sibling Skill directories so each can be discovered and called independently.
