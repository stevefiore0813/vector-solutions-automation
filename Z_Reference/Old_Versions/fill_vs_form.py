#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import time
from pathlib import Path
from typing import Optional, List

from playwright.sync_api import sync_playwright

# ---------- helpers ----------

def get_browser_and_context(pw, headless: bool, cdp_url: Optional[str] = None,
                            storage_state: Optional[str] = None):
    if cdp_url:
        browser = pw.chromium.connect_over_cdp(cdp_url)
        ctx = browser.contexts[0] if browser.contexts else browser.new_context(
            storage_state=storage_state if storage_state else None,
            ignore_https_errors=True
        )
        return browser, ctx

    browser = pw.chromium.launch(headless=headless)
    ctx = browser.new_context(
        storage_state=storage_state if storage_state else None,
        ignore_https_errors=True
    )
    return browser, ctx


def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def guess_people_from_json(data) -> List[str]:
    # Accepts {"people":[{"name":"..."}]}, [{"name":"..."}], or {"name":"..."}
    if isinstance(data, dict):
        if "people" in data and isinstance(data["people"], list):
            return [p.get("name", "").strip() for p in data["people"] if p.get("name")]
        if "name" in data:
            return [data["name"]]
    if isinstance(data, list) and data and isinstance(data[0], dict):
        if "name" in data[0]:
            return [p.get("name", "").strip() for p in data if p.get("name")]
    return []


def wait_for_file(folder: Path, pattern: str = "*.json", timeout_s: int = 900) -> Optional[Path]:
    folder.mkdir(parents=True, exist_ok=True)
    start = time.time()
    while time.time() - start < timeout_s:
        matches = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
        if matches:
            return matches[0]
        time.sleep(1)
    return None


def load_roster(interactive: bool, roster_json_path: Optional[str] = None) -> List[str]:
    if roster_json_path:
        return guess_people_from_json(read_json(Path(roster_json_path)))

    out_dir = Path("/opt/folder_ops/out")
    if interactive:
        try:
            import assignments_gui  # provided by your env
            assignments_gui.main_export()
        except Exception:
            # if GUI picker isn't present, fall back to waiting on /opt/folder_ops/out
            pass

        f = wait_for_file(out_dir, pattern="*.json", timeout_s=900)
        if f:
            return guess_people_from_json(read_json(f))

    return []


def locate_form_frame(page):
    """Return the frame that actually contains the ISO form fields."""
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(500)

    # Fast path: any frame that already exposes the location label
    for f in page.frames:
        try:
            f.wait_for_selector("label:has-text('Where did the training')", timeout=800)
            return f
        except Exception:
            pass

    # Match on URL scent of the activity page, then verify the label exists
    for f in page.frames:
        url = ""
        try:
            url = f.url
        except Exception:
            pass
        if ("c_pro_custom_activities" in url) or ("showUserCustomActivity" in url):
            try:
                f.wait_for_selector("label:has-text('Where did the training')", timeout=1200)
                return f
            except Exception:
                pass

    # Last resort: any non-main frame that has typical form controls and key labels
    non_main = [f for f in page.frames if f != page.main_frame]
    for f in non_main:
        try:
            f.wait_for_selector("input, textarea, button, select", timeout=800)
            f.wait_for_selector("label:has-text('Date'), label:has-text('Instructor')", timeout=800)
            return f
        except Exception:
            pass

    # Dump debug info to help tighten matcher if needed
    urls = []
    for f in page.frames:
        try:
            urls.append(f.url)
        except Exception:
            urls.append("<no-url>")
    raise RuntimeError("Could not find TargetSolutions form iframe. Frames:\n" + "\n".join(urls))


def fill_form_and_save(frame, payload: dict):
    # Location
    try:
        frame.get_by_label("Where did the training take place?").fill(payload["location"])
    except Exception:
        frame.locator("label:has-text('Where did the training')").locator("xpath=following::input[1]").fill(payload["location"])

    # Topics: skip any that aren't present
    for t in payload.get("topics", []):
        try:
            frame.get_by_label(t).check()
        except Exception:
            try:
                frame.locator(f"label:has-text('{t}')").locator("xpath=preceding::input[1]").check()
            except Exception:
                pass

    # Description
    try:
        frame.get_by_label("Give a description of the training that was completed.").fill(payload["description"])
    except Exception:
        frame.locator("label:has-text('description')").locator("xpath=following::textarea[1]").fill(payload["description"])

    # Duration
    try:
        frame.get_by_label("How long was the training?").fill(str(payload["duration_hours"]))
    except Exception:
        frame.locator("label:has-text('How long')").locator("xpath=following::input[1]").fill(str(payload["duration_hours"]))

    # Date
    try:
        frame.get_by_label("Date Completed").fill(payload["date"])
    except Exception:
        frame.locator("label:has-text('Date Complete'), label:has-text('Date Completed')").locator("xpath=following::input[1]").fill(payload["date"])

    # Instructor
    try:
        frame.get_by_label("Who led the training?").fill(payload["instructor"])
    except Exception:
        frame.locator("label:has-text('Who led the training')").locator("xpath=following::input[1]").fill(payload["instructor"])

    # Click Save and Add Users
    try:
        frame.get_by_role("button", name="Save and Add Users").click()
    except Exception:
        frame.locator("button:has-text('Save and Add Users')").click()


def add_people(frame, roster: List[str]):
    for person in roster:
        # Basic search + enter; adjust if the UI needs an explicit add button
        try:
            frame.get_by_label("User search").fill(person)
        except Exception:
            # fallback: common search box patterns
            frame.locator("input[placeholder*='search' i], input[type='search']").first.fill(person)
        frame.keyboard.press("Enter")
        time.sleep(0.3)


# ---------- main ----------

def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("url")
    ap.add_argument("payload_json")
    ap.add_argument("--interactive", action="store_true")
    ap.add_argument("--roster-json", default=None)
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--cdp", default=None)
    ap.add_argument("--storage-state", default="05_Dev_Env/Dependencies/storage_state.json")
    ap.add_argument("--capture-login", action="store_true")
    args = ap.parse_args()

    # Be explicit about the URL we are using and kill stray whitespace
    args.url = (args.url or "").strip()
    print(f"[nav] using URL -> {repr(args.url)}")

    with sync_playwright() as pw:
        # ---- LOGIN CAPTURE ----
        if args.capture_login:
            browser = pw.chromium.launch(headless=False)
            ctx = browser.new_context(ignore_https_errors=True)
            page = ctx.new_page()
            page.goto(args.url, wait_until="domcontentloaded")
            input("Login manually, then press ENTER here to save session...")
            Path(args.storage_state).parent.mkdir(parents=True, exist_ok=True)
            ctx.storage_state(path=args.storage_state)
            print(f"[saved login] {args.storage_state}")
            return

        # ---- NORMAL RUN ----
        payload = read_json(Path(args.payload_json))

        browser, ctx = get_browser_and_context(
            pw,
            headless=args.headless,
            cdp_url=args.cdp,
            storage_state=args.storage_state
        )

        page = ctx.new_page()
        page.set_default_navigation_timeout(120000)

        # Hardened nav to avoid about:blank stalls
        page.goto(args.url, wait_until="commit")
        page.wait_for_load_state("domcontentloaded")
        if page.url == "about:blank":
            print("[nav] still at about:blank; forcing window.location")
            page.evaluate(f"window.location.href = {json.dumps(args.url)}")
            page.wait_for_load_state("domcontentloaded")
        print(f"[nav] landed on -> {page.url}")

        # Find the form frame and run the workflow
        frame = locate_form_frame(page)
        frame.wait_for_selector("label:has-text('Where did the training')", timeout=5000)

        fill_form_and_save(frame, payload)

        roster = load_roster(args.interactive, roster_json_path=args.roster_json)
        if roster:
            add_people(frame, roster)

        print("✅ Training submitted & roster handled")

if __name__ == "__main__":
    main()
