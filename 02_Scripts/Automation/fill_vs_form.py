import json, os, sys, time
from pathlib import Path
from typing import List, Optional
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ---------- helpers ----------
def get_browser_and_context(pw, headless: bool, cdp_url: str | None = None,
                            storage_state: str | None = None):
    # If CDP not provided, launch Chromium inside WSL
    if cdp_url:
        browser = pw.chromium.connect_over_cdp(cdp_url)
        ctx = browser.contexts[0] if browser.contexts else browser.new_context(
            storage_state=storage_state if storage_state else None
        )
        return browser, ctx
    # Local launch
    browser = pw.chromium.launch(headless=headless)
    ctx = browser.new_context(storage_state=storage_state if storage_state else None)
    return browser, ctx

def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))

def guess_people_from_json(data) -> List[str]:
    # Flexible: accept ["Fiore, Steve", ...] or {"people":[...]} or [{"name": "..."}]
    if isinstance(data, list):
        if data and isinstance(data[0], dict) and "name" in data[0]:
            return [d["name"] for d in data]
        return [str(x) for x in data]
    if isinstance(data, dict):
        if "people" in data:
            return guess_people_from_json(data["people"])
        if "names" in data:
            return guess_people_from_json(data["names"])
    return []

def wait_for_file(folder: Path, pattern: str = "*.json", timeout_s: int = 900) -> Optional[Path]:
    """Wait up to timeout_s for a new JSON file to appear in folder."""
    folder.mkdir(parents=True, exist_ok=True)
    start = time.time()
    seen = {p: p.stat().st_mtime for p in folder.glob(pattern)}
    while time.time() - start < timeout_s:
        time.sleep(1)
        candidates = list(folder.glob(pattern))
        for p in candidates:
            if p not in seen or p.stat().st_mtime > seen[p]:
                return p
    return None

def load_roster(interactive: bool, roster_json_path: Optional[str]) -> List[str]:
    if roster_json_path:
        return guess_people_from_json(read_json(Path(roster_json_path)))

    if interactive:
        # Launch your GUI and wait for its export to land in /opt/folder_ops/out
        try:
            sys.path.append("/opt/folder_ops")
            import assignments_gui  # must pop its GUI when imported or via a function call
            # Try to call an entrypoint if it exists; otherwise just import and let it run.
            if hasattr(assignments_gui, "open_and_export"):
                assignments_gui.open_and_export()
        except Exception:
            pass
        out = wait_for_file(Path("/opt/folder_ops/out"), pattern="*.json", timeout_s=900)
        if not out:
            return []
        return guess_people_from_json(read_json(out))

    # Non-interactive fallback: read a plain text roster file if present
    roster_txt = Path("05_Dev_Env/Dependencies/roster.txt")
    if roster_txt.exists():
        return [line.strip() for line in roster_txt.read_text(encoding="utf-8").splitlines() if line.strip()]
    return []

def add_personnel_from_assignments(page, people: List[str]):
    # You need to tweak these selectors once against your UI. This is the pattern:
    for person in people:
        try:
            # Example flow; replace as needed:
            # 1) Find the "Add User" search box in the modal
            page.get_by_placeholder("Search").fill(person)
            page.keyboard.press("Enter")
            page.get_by_role("button", name="Add", exact=False).click()
            time.sleep(0.2)
        except Exception:
            # As a last resort, use a broader locator and nth()
            pass

def fill_form_and_save(page, payload: dict):
    # Labels must match your instance text. Adjust if they differ.
    page.get_by_label("Where did the training take place?").fill(payload["location"])
    for label in payload.get("training_types", []):
        try:
            page.get_by_label(label).check()
        except Exception:
            page.get_by_role("checkbox", name=label, exact=False).check()
    page.get_by_label("Give a description of the training that was completed.").fill(payload["description"])
    page.get_by_label("How long was the training?").fill(str(payload["duration_hours"]))
    page.get_by_label("Date Complete").fill(payload["date_complete"])
    if payload.get("instructor"):
        page.get_by_label("Who led the training?").fill(payload["instructor"])
    page.get_by_role("button", name="Save and Add Users", exact=False).click()

def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("url")
    ap.add_argument("payload_json")
    ap.add_argument("--interactive", action="store_true")
    ap.add_argument("--roster-json", default=None)
    ap.add_argument("--roster-dir", default=None)
    ap.add_argument("--roster-pattern", default="*.json")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--cdp", default=None, help="Optional: CDP URL. Omit to launch local Chromium.")
    ap.add_argument("--storage-state", default="05_Dev_Env/Dependencies/storage_state.json",
                    help="Path to Playwright storage state (cookies/session).")
    ap.add_argument("--capture-login", action="store_true",
                    help="Open login once and save storage state, then exit.")
    args = ap.parse_args()

    payload = read_json(Path(args.payload_json))

    with sync_playwright() as pw:
        # If we’re capturing login, open a clean context, let you log in, then save and exit
        if args.capture_login:
            browser = pw.chromium.launch(headless=False)
            ctx = browser.new_context()
            page = ctx.new_page()
            page.goto(args.url, wait_until="domcontentloaded")
            print("[login] Log in manually, then press ENTER here...")
            input()
            Path(args.storage_state).parent.mkdir(parents=True, exist_ok=True)
            ctx.storage_state(path=args.storage_state)
            print(f"[login] Saved storage state -> {args.storage_state}")
            return 0

        # Normal mode
        browser, ctx = get_browser_and_context(
            pw, headless=args.headless, cdp_url=args.cdp, storage_state=args.storage_state
        )
        page = ctx.new_page()
        page.goto(args.url, wait_until="domcontentloaded")

        fill_form_and_save(page, payload)

        people = load_roster(
            interactive=args.interactive,
            roster_json_path=args.roster_json,
            roster_dir=args.roster_dir,
            roster_pattern=args.roster_pattern
        )
        if people:
            add_personnel_from_assignments(page, people)

        print("Filled form and clicked 'Save and Add Users'.")
        return 0
