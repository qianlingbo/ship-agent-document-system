# Runtime and installation

## Requirements

- Python 3.10+
- Python packages: `openpyxl`, `Pillow`, `numpy`
- Commands on `PATH`: `pdftoppm`, `magick`, `tesseract`
- Tesseract English language data

Install Python packages in an isolated environment:

```bash
python3 -m venv .venv
. .venv/bin/activate
python3 -m pip install openpyxl Pillow numpy
```

macOS system packages:

```bash
brew install poppler imagemagick tesseract
```

Debian/Ubuntu system packages:

```bash
sudo apt-get update
sudo apt-get install -y poppler-utils imagemagick tesseract-ocr tesseract-ocr-eng
```

## Skill locations

This directory follows the AgentSkills `SKILL.md` convention.

- OpenClaw workspace: `<workspace>/skills/ship-crew-info-export/`
- OpenClaw shared: `~/.openclaw/skills/ship-crew-info-export/`
- OpenClaw local install: `openclaw skills install /path/ship-crew-info-export`
- Hermes personal: `~/.hermes/skills/ship-crew-info-export/`
- Codex personal: `~/.codex/skills/ship-crew-info-export/`

After installation, start a new agent session if the skill list is cached.

## Low-cost model guidance

Give the model only the four input paths and desired output path. The script performs image rotation, OCR-guided MRZ location, biometric-region cropping, compression, workbook insertion, and validation. The model should not inspect all passport pages unless the command fails.
