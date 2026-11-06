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

`xlrd` is included in `scripts/requirements.txt` so legacy `.xls` Crew List and Port of Call files can be read.

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

## Install the total and component Skills

Install all 11 sibling Skills so each component can be discovered independently:

```bash
python3 "{baseDir}/scripts/install_skills.py" --target ~/.codex/skills
```

Alternative targets:

- Hermes: `~/.hermes/skills`
- OpenClaw workspace: `<workspace>/skills`
- Codex: `~/.codex/skills`

The installer keeps one shared runtime in `ship-agent-document-system` and installs ten small component Skill folders beside it. Start a new agent session if the Skill list is cached.
