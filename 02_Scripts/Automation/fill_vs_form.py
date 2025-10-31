import json, os, sys, time
from pathlib import Path
from typing import List, Optional
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ---------- helpers ----------
def get_browser_and_context(pw, headless: bool):
    cdp = os.environ.get("BROWSER_CDP_URL", "").strip()
    if cdp:
        browser = pw.chromium.connect_over_cdp(cdp)
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()
    else:
        browser = pw.chromium.launch(headless=headless)
        ctx = browser.new_context()
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
    # Args: url, payload_json, [--interactive] [--roster-json PATH] [--headless]
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("url")
    ap.add_argument("payload_json")
    ap.add_argument("--interactive", action="store_true", help="Open assignments GUI and wait for its JSON export")
    ap.add_argument("--roster-json", default=None, help="Use this roster JSON (skips GUI)")
    ap.add_argument("--headless", action="store_true")
    args = ap.parse_args()

    payload = read_json(Path(args.payload_json))
    people = load_roster(interactive=args.interactive, roster_json_path=args.roster_json)

    with sync
