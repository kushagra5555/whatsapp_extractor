# WhatsApp Contact Extractor

Fast, resumable WhatsApp Web automation for collecting unsaved phone numbers from a user’s own chat list and exporting them to CSV and Excel.

Built with Python, Playwright, and `openpyxl`, this project demonstrates browser automation, persistent sessions, DOM-based extraction, resumable processing, and reliable streaming output.

> Use this tool only with an account and data you are authorized to access. Respect WhatsApp’s Terms of Service, privacy laws, and the consent of the people whose data you process.

## Highlights

- **DOM-first fast path:** Reads phone numbers directly from the loaded chat list before doing expensive interaction.
- **Persistent login:** Stores the browser session locally so the QR code is normally scanned only once.
- **Safe resume behavior:** Writes results as they are found, avoids duplicates, and can be stopped with `Ctrl+C` without losing saved output.
- **Dual export:** Produces both `contacts.csv` and a formatted `contacts.xlsx` file.
- **Resilient scanning:** Continues past individual parsing or rendering problems instead of failing the complete run.
- **Windows-friendly launcher:** Includes a ready-to-run batch file for a straightforward local workflow.

## How it works

1. Launches installed Google Chrome with a dedicated persistent profile.
2. Opens WhatsApp Web and waits for the chat list to become available.
3. Scrolls the virtualized chat pane in small, reliable steps.
4. Detects unsaved phone-number titles using the rendered DOM.
5. Normalizes and deduplicates numbers before writing them immediately.
6. Rebuilds the Excel workbook from the CSV so both outputs stay consistent.

## Quick start

### Requirements

- Windows 10 or newer
- Python 3.12+
- Google Chrome
- Permission to access the WhatsApp account being automated

### Install

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

### Run

```powershell
python main.py
```

Or double-click [`open_whatsapp.bat`](open_whatsapp.bat).

On the first run, scan the QR code in the opened browser window. Later runs reuse the local session automatically. Results are written to `contacts.csv` and `contacts.xlsx` in the project directory.

## Configuration

The default session directory is `wa_session`. To use a fresh login or a different browser executable:

```powershell
$env:WHATSAPP_SESSION_DIR = "$PWD\wa_session_fresh"
$env:WHATSAPP_BROWSER_PATH = "C:\Program Files\Google\Chrome\Application\chrome.exe"
python main.py
```

The script also supports graceful interruption with `Ctrl+C`; already collected numbers remain saved.

## Project structure

```text
whatsapp_extractor/
├── main.py              # Playwright automation and export pipeline
├── requirements.txt     # Runtime dependencies
├── open_whatsapp.bat    # Windows launcher
├── README.md            # Project documentation
└── .gitignore           # Protects sessions and personal exports
```

## Privacy and security

Browser sessions contain authentication state and must stay local. Contact exports may contain personal information. The repository intentionally excludes session folders, browser profiles, exports, and debug files through `.gitignore`. Never commit those files to a public repository.

## Technical notes

- The extractor targets unsaved phone-number titles visible in WhatsApp Web’s chat list.
- Saved contact names are skipped by design.
- WhatsApp Web’s DOM can change over time; selectors may need maintenance after major UI updates.
- This project is intended for controlled, authorized, personal workflows—not bulk messaging, scraping of unauthorized accounts, or bypassing platform protections.
