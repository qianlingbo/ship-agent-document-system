# Installation and authorization

Run `python3 "{baseDir}/scripts/check_runtime.py"` before changing the environment.

## Python

Python 3.10 or newer is required. Install the deterministic runtime packages with:

```bash
python3 -m pip install -r "{baseDir}/scripts/requirements.txt"
```

## System tools

Required for PDF OCR and passport-image processing:

- Poppler (`pdftoppm`)
- Tesseract OCR with English; Chinese improves certificate OCR
- ImageMagick (`magick`)

Recommended for RAR and other archives: `7zz`, `7z`, or `unar`.

On macOS with Homebrew:

```bash
brew install poppler tesseract tesseract-lang imagemagick sevenzip
```

## Agent Mail CLI

Install the official CLI:

```bash
npm install -g @tencent-qqmail/agently-cli
```

Authorize interactively:

```bash
agently-cli auth login
```

Confirm the connected account:

```bash
agently-cli +me
```

Authorization determines the active mailbox, so the Skill is not tied to one address. Re-run `agently-cli auth login` to switch accounts. Never place OAuth tokens or mailbox data inside this Skill.

## Skill locations

Copy the entire `ship-agent-document-system` directory to one of these locations, then start a new agent session if its Skill list is cached:

- Hermes: `~/.hermes/skills/ship-agent-document-system/`
- OpenClaw workspace: `<workspace>/skills/ship-agent-document-system/`
- Codex: `~/.codex/skills/ship-agent-document-system/`
