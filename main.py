"""
WhatsApp Phone Number Extractor
================================
- Uses the installed Google Chrome browser with a persistent session folder.
- Scan QR code ONCE. All future runs reuse the session automatically.
- Extracts ONLY unsaved phone numbers directly from the chat list.
- Does NOT click on or open any chats (hyper-fast).
- Saves instantly to contacts.xlsx (text format, no scientific notation).
- Press Ctrl+C anytime to stop safely — all saved data is preserved.
"""

import asyncio
import csv
import re
import os
import sys
import time
import signal
import ctypes
from playwright.async_api import async_playwright, Page, BrowserContext
import openpyxl
from openpyxl.styles import Font

# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────
CSV_FILE     = "contacts.csv"
XLSX_FILE    = "contacts.xlsx"
# Set WHATSAPP_SESSION_DIR in PowerShell to use a separate, fresh login session.
SESSION_DIR  = os.path.abspath(os.environ.get("WHATSAPP_SESSION_DIR", "wa_session"))
SCROLL_PAUSE = 500                   # ms between scrolls (allow WhatsApp to render)
SCROLL_STEP  = 300                   # px per scroll — small steps = nothing missed


def find_browser_executable():
    """Use Chrome instead of Playwright's version-specific bundled Chromium."""
    candidates = [
        os.environ.get("WHATSAPP_BROWSER_PATH"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "Application", "chrome.exe"),
    ]
    return next((path for path in candidates if path and os.path.isfile(path)), None)

stop_requested = False

def _handle_sigint(signum, frame):
    global stop_requested
    stop_requested = True

signal.signal(signal.SIGINT, _handle_sigint)


# ─────────────────────────────────────────────────────────────
# OUTPUT FILE HELPERS
# ─────────────────────────────────────────────────────────────

def check_files_writable():
    locked = []
    for path in [CSV_FILE, XLSX_FILE]:
        try:
            with open(path, 'a'):
                pass
        except PermissionError:
            locked.append(path)
    if locked:
        print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        print("❌  OUTPUT FILE IS LOCKED")
        print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        for f in locked:
            print(f"  • {f}")
        print("\nClose Microsoft Excel or any program")
        print("that has these files open, then retry.")
        print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        sys.exit(1)


def init_output_files():
    """Ensure output files exist. Rebuild XLSX from CSV so it's always complete."""
    # Ensure CSV exists with header
    if not os.path.exists(CSV_FILE):
        with open(CSV_FILE, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f).writerow(['Phone Number'])

    # Always rebuild XLSX from CSV — this way XLSX is always complete
    # even if it was open in Excel during a previous run.
    try:
        if os.path.exists(XLSX_FILE):
            os.remove(XLSX_FILE)
    except Exception:
        pass

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Phone Numbers"
    ws.column_dimensions['A'].width = 22
    hdr = ws.cell(row=1, column=1)
    hdr.value = 'Phone Number'
    hdr.font = Font(bold=True)

    # Copy all numbers already in CSV into XLSX
    if os.path.exists(CSV_FILE):
        with open(CSV_FILE, 'r', encoding='utf-8') as f:
            for row in csv.reader(f):
                if row and row[0] and row[0] != 'Phone Number':
                    r = ws.max_row + 1
                    cell = ws.cell(row=r, column=1)
                    cell.value = str(row[0])
                    cell.number_format = '@'

    wb.save(XLSX_FILE)
    print(f'✅ XLSX rebuilt with {ws.max_row - 1} existing numbers.')


def write_number(number: str, csv_writer, csv_fh, xlsx_ws, xlsx_wb):
    """Write one number to CSV (instant flush) and XLSX (text-locked cell)."""
    csv_writer.writerow([number])
    csv_fh.flush()

    row = xlsx_ws.max_row + 1
    cell = xlsx_ws.cell(row=row, column=1)
    cell.value = str(number)
    cell.number_format = '@'  # force text format — prevents scientific notation
    xlsx_wb.save(XLSX_FILE)


# ─────────────────────────────────────────────────────────────
# PHONE NUMBER HELPERS
# ─────────────────────────────────────────────────────────────

_CLEAN  = re.compile(r'[\s\-\(\)\u200e\u200f\u202a\u202b\u202c\u202d\u202e\u2066\u2069\u00a0\xa0\u200b]')
_PHONE  = re.compile(r'^\+?\d{7,16}$')

def parse_phone(text: str) -> str:
    """Return a normalised +XXXX phone string, or '' if not a phone number."""
    if not text:
        return ''
    c = _CLEAN.sub('', text)
    if _PHONE.match(c):
        return c if c.startswith('+') else '+' + c
    return ''


# ─────────────────────────────────────────────────────────────
# MAIN EXTRACTION LOOP
# ─────────────────────────────────────────────────────────────

async def process_chats(page: Page):
    global stop_requested

    extracted   = 0
    skipped     = 0
    processed_titles:  set[str] = set()
    processed_numbers: set[str] = set()
    start_time  = time.time()

    # Create a debug log to see exactly what strings WhatsApp provides
    debug_log = open("debug_titles.txt", "w", encoding="utf-8")

    # Load already-saved numbers so re-runs don't duplicate
    if os.path.exists(CSV_FILE):
        with open(CSV_FILE, 'r', encoding='utf-8') as f:
            for row in csv.reader(f):
                if row and row[0] != 'Phone Number':
                    processed_numbers.add(row[0].strip())

    csv_fh  = open(CSV_FILE, 'a', newline='', encoding='utf-8')
    csv_wr  = csv.writer(csv_fh)
    xlsx_wb = openpyxl.load_workbook(XLSX_FILE)
    xlsx_ws = xlsx_wb.active

    def save(number: str) -> bool:
        nonlocal extracted
        if number and number not in processed_numbers:
            processed_numbers.add(number)
            write_number(number, csv_wr, csv_fh, xlsx_ws, xlsx_wb)
            extracted += 1
            return True
        return False

    # ── Wait for WhatsApp to finish loading ───────────────────────────────
    print("\n⏳ Waiting for WhatsApp Web to load...")
    try:
        await page.wait_for_selector('#pane-side', timeout=120_000)
        await page.wait_for_selector(
            '#pane-side [role="listitem"], #pane-side [role="row"]',
            timeout=60_000
        )
    except Exception:
        print("❌ Chat list not found. Make sure you are logged in to WhatsApp Web.")
        csv_fh.close()
        debug_log.close()
        return

    pane = page.locator('#pane-side')

    # Try to read total chat count from the virtualised list
    total = 300
    try:
        el = page.locator('[aria-rowcount]').first
        rc = await el.get_attribute('aria-rowcount', timeout=1000)
        if rc and rc.isdigit():
            total = int(rc)
    except Exception:
        pass

    # Enable ANSI escape codes on Windows so the status block renders correctly
    if sys.platform == 'win32':
        try:
            k32 = ctypes.windll.kernel32
            k32.SetConsoleMode(k32.GetStdHandle(-11), 7)
        except Exception:
            pass

    STATUS_LINES = 10

    # ── Confirmation prompt ───────────────────────────────────────────────
    os.system('cls' if os.name == 'nt' else 'clear')
    print('\n')
    print('  ✅ WhatsApp Web is open and ready.')
    print('  ───────────────────────────────────')
    print('  Press ENTER to start extracting...')
    print('  (Close this window to cancel)      ')
    print('  ───────────────────────────────────')
    input()

    os.system('cls' if os.name == 'nt' else 'clear')
    print('\n' * STATUS_LINES)

    def render_status(chat: str):
        elapsed   = time.time() - start_time
        done      = extracted + skipped
        remaining = max(0, total - done)
        eta       = (elapsed / done * remaining) if done > 0 else 0
        e_str     = time.strftime('%H:%M:%S', time.gmtime(elapsed))
        eta_str   = time.strftime('%H:%M:%S', time.gmtime(eta)) if eta > 0 else 'Calculating...'

        sys.stdout.write(f'\033[{STATUS_LINES}A\033[J')
        print('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━')
        print('        WHATSAPP EXTRACTOR          ')
        print('  Press Ctrl+C anytime to stop safely')
        print('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━')
        print(f'  Chat    : {chat[:36]:<36}')
        print(f'  Saved   : {extracted} phone numbers')
        print(f'  Skipped : {skipped} (saved contacts)')
        print(f'  Left    : ~{remaining}')
        print(f'  Time    : {e_str}   ETA: {eta_str}')
        print('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━')
        sys.stdout.flush()

    # ── Scroll + extract loop ─────────────────────────────────────────────
    # First scroll back to TOP so we never miss chats at the beginning
    await pane.evaluate('el => el.scrollTo(0, 0)')
    await page.wait_for_timeout(600)

    no_change_streak = 0

    JS_GET_TITLES = """() => {
        const results = [];
        const pane = document.querySelector('#pane-side');
        if (!pane) return results;
        const rows = pane.querySelectorAll(
            '[role="listitem"], [role="row"], [data-testid="cell-frame-container"]'
        );
        for (const row of rows) {
            let found = false;
            for (const s of row.querySelectorAll('span[title]')) {
                const t = (s.getAttribute('title') || '').trim();
                if (t) { results.push(t); found = true; break; }
            }
            if (found) continue;
            
            const al = (row.getAttribute('aria-label') || '').trim();
            if (al) { results.push(al); continue; }
            
            const txt = ((row.innerText || '').split('\\n')[0] || '').trim();
            if (txt) results.push(txt);
        }
        return results;
    }"""

    try:
        while not stop_requested:
            titles = await page.evaluate(JS_GET_TITLES)
            found_new = False

            for raw_title in titles:
                if stop_requested:
                    break

                clean_title = _CLEAN.sub('', raw_title).strip()

                if not clean_title or clean_title in processed_titles:
                    continue

                processed_titles.add(clean_title)
                found_new = True
                debug_log.write(f"TITLE: {repr(clean_title)}\n")
                debug_log.flush()
                render_status(clean_title)

                # FAST PATH ONLY
                phone = parse_phone(clean_title)
                if phone:
                    debug_log.write(f"  -> PHONE: {phone}\n")
                    debug_log.flush()
                    if not save(phone):
                        skipped += 1
                else:
                    skipped += 1

            prev_h = await pane.evaluate('el => el.scrollTop')
            await pane.evaluate(f'el => el.scrollBy(0, {SCROLL_STEP})')
            await page.wait_for_timeout(SCROLL_PAUSE)
            new_h  = await pane.evaluate('el => el.scrollTop')

            if not found_new and prev_h == new_h:
                no_change_streak += 1
                if no_change_streak >= 6:   # increased — be very sure we hit the end
                    break
            else:
                no_change_streak = 0

    except (asyncio.CancelledError, KeyboardInterrupt):
        pass
    finally:
        try:
            xlsx_wb.save(XLSX_FILE)
        except Exception:
            pass
        csv_fh.close()
        debug_log.close()

    os.system('cls' if os.name == 'nt' else 'clear')
    status = '🛑  STOPPED SAFELY' if stop_requested else '🎉  DONE'
    print('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━')
    print(f'  {status}')
    print('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━')
    print(f'  Phone numbers saved : {extracted}')
    print(f'  Saved names skipped : {skipped}')
    print(f'  CSV  → {os.path.abspath(CSV_FILE)}')
    print(f'  XLSX → {os.path.abspath(XLSX_FILE)}')
    print('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━')
    print('\n  ✅ Open contacts.xlsx in Excel')
    print('     Numbers display in full — no scientific notation.')


# ─────────────────────────────────────────────────────────────
# PLAYWRIGHT LAUNCHER
# ─────────────────────────────────────────────────────────────

async def run():
    os.makedirs(SESSION_DIR, exist_ok=True)
    browser_executable = find_browser_executable()
    if not browser_executable:
        raise RuntimeError(
            "Google Chrome was not found. Install Google Chrome or set "
            "WHATSAPP_BROWSER_PATH to your browser's chrome.exe path."
        )

    print('\n🚀 Launching browser...')
    if os.listdir(SESSION_DIR):
        print('✅ Existing session found — no QR code needed.')
    else:
        print('📱 First run: please scan the WhatsApp QR code.')
        print('   After that, this step is skipped forever.')

    async with async_playwright() as p:
        browser: BrowserContext = await p.chromium.launch_persistent_context(
            user_data_dir=SESSION_DIR,
            executable_path=browser_executable,
            headless=False,
            args=['--start-maximized'],
        )

        page = browser.pages[0] if browser.pages else await browser.new_page()

        print('🌐 Opening WhatsApp Web...')
        await page.goto('https://web.whatsapp.com/')

        try:
            await process_chats(page)
        finally:
            try:
                await browser.close()
            except Exception:
                pass
            print('\n👋 Browser closed. Session saved — next run needs no QR.')


# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────

def main():
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

    check_files_writable()
    init_output_files()

    try:
        asyncio.run(run())
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    main()
