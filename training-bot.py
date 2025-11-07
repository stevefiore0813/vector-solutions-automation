#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
training-bot.py
Single-file, sync Playwright flow:
1) (Optional) fetch MiniCAD (left intact, do not break what's working )
2) Login and persist session
3) Navigate to Training form (either direct URL or by clicking a dashboard element CSS selector)
4) Fill form (location label: "Where did this training take place?")
5) Click "Save and Add Users" -> search & select roster names (Last, First) in the middle list -> Continue
6) Return to form -> Submit/Save
"""

import argparse
import csv
import os
import re
import sys
import time
from typing import Dict, List, Optional

from playwright.sync_api import Playwright, sync_playwright

AUTH_STATE = os.environ.get("AUTH_STATE_PATH", "auth_state.json")
DEFAULT_CSV = "/home/training-bot/projects/vector-solutions/trainings.csv"

# ------------------------------
# Logging helpers
# ------------------------------
def log(msg: str) -> None:
    print(msg, flush=True)

def debug(msg: str) -> None:
    print(f"[debug] {msg}", flush=True)

def warn(msg: str) -> None:
    print(f"[warn] {msg}", flush=True)

def fail(msg: str) -> None:
    print(f"[error] {msg}", flush=True)
    sys.exit(1)

# ------------------------------
# Robust page helpers
# ------------------------------
def retry(times: int, sleep_sec: float, fn, *args, **kwargs):
    last_exc = None
    for i in range(times):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            debug(f"retry {i+1}/{times} -> {e}")
            time.sleep(sleep_sec)
    raise last_exc

def scroll_into_view_safe(page, locator):
    try:
        locator.scroll_into_view_if_needed(timeout=1500)
    except Exception:
        pass

def click_hard(page, locator, name_for_logs: str = "element"):
    scroll_into_view_safe(page, locator)
    retry(3, 0.5, locator.click, timeout=3000)
    debug(f"clicked {name_for_logs}")

def goto_robust(page, url: str, tag: str):
    debug(f"navigating to {tag}: {url}")
    for i in range(3):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=30000)
            if "about:blank" in page.url:
                raise RuntimeError("Stuck at about:blank")
            page.wait_for_load_state("networkidle", timeout=10000)
            debug(f"arrived at {tag}")
            return
        except Exception as e:
            warn(f"navigate attempt {i+1} failed for {tag}: {e}")
            time.sleep(1)
    page.reload(wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle", timeout=10000)

def find_by_label_like(page, label_text: str):
    strategies = [
        lambda: page.get_by_label(re.compile(label_text, re.I)),
        lambda: page.get_by_placeholder(re.compile(label_text, re.I)),
        lambda: page.get_by_role("textbox", name=re.compile(label_text, re.I)),
        lambda: page.locator(f"label:has-text('{label_text}')").locator("xpath=following::*[self::input or self::textarea or self::select][1]"),
        lambda: page.locator(f"[aria-label*='{label_text}'],[name*='{label_text}'],[id*='{label_text}']"),
    ]
    for s in strategies:
        try:
            el = s()
            _ = el.first
            return el.first
        except Exception:
            continue
    raise RuntimeError(f"Could not find control by label '{label_text}'")

def safe_fill_by_label(page, label_text: str, value: str):
    el = find_by_label_like(page, label_text)
    scroll_into_view_safe(page, el)
    try:
        el.select_option(value)
        debug(f"select_option ok -> {label_text} = {value}")
        return
    except Exception:
        pass
    try:
        el.fill("")
    except Exception:
        pass
    retry(2, 0.25, el.fill, value)
    debug(f"filled -> {label_text} = {value}")

def safe_check_by_label(page, label_text: str, should_check: bool = True):
    try:
        cb = page.get_by_role("checkbox", name=re.compile(label_text, re.I)).first
        scroll_into_view_safe(page, cb)
        if should_check:
            if not cb.is_checked():
                cb.check(timeout=3000)
        else:
            if cb.is_checked():
                cb.uncheck(timeout=3000)
        debug(f"checkbox -> {label_text} = {should_check}")
        return
    except Exception:
        pass
    try:
        cb = page.locator(f"label:has-text('{label_text}')").locator("xpath=following::input[@type='checkbox'][1]").first
        scroll_into_view_safe(page, cb)
        if should_check:
            if cb.get_attribute("checked") is None:
                cb.check(timeout=3000)
        else:
            if cb.get_attribute("checked") is not None:
                cb.uncheck(timeout=3000)
        debug(f"checkbox(fallback) -> {label_text} = {should_check}")
        return
    except Exception:
        pass
    warn(f"checkbox not found -> {label_text}")

def click_button_like(page, text: str):
    try:
        btn = page.get_by_role("button", name=re.compile(text, re.I)).first
        click_hard(page, btn, f"button:{text}")
        return
    except Exception:
        pass
    try:
        btn = page.locator(f"input[type='submit'][value*='{text}'], input[type='button'][value*='{text}']").first
        click_hard(page, btn, f"submit:{text}")
        return
    except Exception:
        pass
    btn = page.get_by_text(re.compile(text, re.I)).first
    click_hard(page, btn, f"text:{text}")

def wait_text(page, text: str, timeout_ms: int = 8000):
    page.get_by_text(re.compile(text, re.I)).first.wait_for(timeout=timeout_ms)
    debug(f"found text: {text}")

def detect_form_ready(page):
    for key in [
        r"Company Training",
        r"Where did this training take place?",
        r"Training Details",
        r"Description",
    ]:
        try:
            wait_text(page, key, timeout_ms=5000)
            return
        except Exception:
            continue
    warn("Couldn't confirm a specific form hallmark; proceeding anyway.")

# ------------------------------
# Browser/context
# ------------------------------
def browser_ctx(pw: Playwright, headed: bool):
    debug("launching bundled Chromium")
    browser = pw.chromium.launch(headless=not headed, args=["--disable-renderer-backgrounding"])
    context = browser.new_context(ignore_https_errors=True, viewport={"width": 1440, "height": 900})
    context.set_default_timeout(15000)
    return browser, context
# ------------------------------
# Login
# ------------------------------
def scripted_login(page, user: Optional[str], password: Optional[str]):
    if not user or not password:
        debug("no creds provided; skipping scripted login")
        return False

    candidates = [
        ("username", r"(email|user(?:name)?|login)", user),
        ("password", r"(pass|password)", password),
    ]
    ok = False
    for label, patt, value in candidates:
        try:
            safe_fill_by_label(page, patt, value)
            ok = True
        except Exception:
            try:
                sel = "input[type='text'],input[type='email']" if "user" in label else "input[type='password']"
                el = page.locator(sel).first
                el.fill(value)
                ok = True
            except Exception:
                pass

    try:
        click_button_like(page, "Login")
    except Exception:
        warn("couldn't find a login button; if SSO, session might auto-continue")

    try:
        page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    debug("scripted login attempted")
    return ok

# ------------------------------
# MiniCAD / Roster
# ------------------------------
def fetch_json_basic_auth(url: str, user: str, password: str) -> dict:
    with sync_playwright() as pw:
        req = pw.request.new_context(
            http_credentials={"username": user, "password": password},
            timeout=30000,
        )
        resp = req.get(url)
        if not resp.ok:
            raise RuntimeError(f"GET {url} -> {resp.status} {resp.status_text()}")
        return resp.json()

def get_roster_from_minicad(url: str, basic_user: str, basic_pass: str) -> List[str]:
    data = fetch_json_basic_auth(url, basic_user, basic_pass)
    roster = []

    def norm(s): return (s or "").strip()
    items = data if isinstance(data, list) else data.get("items") or data.get("Units") or []

    for it in items:
        last = norm(it.get("LastName") or it.get("LName") or it.get("Last") or "")
        first = norm(it.get("FirstName") or it.get("FName") or it.get("First") or "")
        disp = None
        if last and first:
            disp = f"{last}, {first}"
        elif it.get("FullName"):
            parts = norm(it["FullName"]).split()
            if len(parts) >= 2:
                disp = f"{parts[-1]}, {' '.join(parts[:-1])}"
        if disp:
            roster.append(disp)

    seen = set()
    uniq = []
    for n in roster:
        if n not in seen:
            uniq.append(n)
            seen.add(n)
    return uniq

# ------------------------------
# CSV Training Scenario
# ------------------------------
def read_training_row(csv_path: str) -> Dict[str, str]:
    if not os.path.exists(csv_path):
        fail(f"CSV not found: {csv_path}")
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if any(v.strip() for v in row.values() if v is not None):
                def g(k): return (row.get(k) or "").strip()
                return {
                    "Location": g("Location"),
                    "Checkbox Label": g("Checkbox Label"),
                    "Description": g("Description"),
                    "Duration": g("Duration"),
                    "Instructor": g("Instructor"),
                }
    fail("No non-empty rows found in CSV")

# ------------------------------
# Fill Training Form
# ------------------------------
def fill_training_form(page, data: Dict[str, str]):
    if data.get("Location"):
        safe_fill_by_label(page, r"Where did this training take place\??", data["Location"])
    if data.get("Checkbox Label"):
        safe_check_by_label(page, data["Checkbox Label"], True)
    if data.get("Description"):
        safe_fill_by_label(page, r"Description", data["Description"])
    if data.get("Duration"):
        try:
            safe_fill_by_label(page, r"Duration", data["Duration"])
        except Exception:
            safe_fill_by_label(page, r"(Total\s*Time|Hours)", data["Duration"])
    if data.get("Instructor"):
        try:
            safe_fill_by_label(page, r"Instructor", data["Instructor"])
        except Exception:
            inp = find_by_label_like(page, r"Instructor")
            inp.type(data["Instructor"])
            time.sleep(0.4)
            try:
                page.keyboard.press("Enter")
            except Exception:
                pass
    debug("form fields filled")

# ------------------------------
# Add Users flow
# ------------------------------
def add_users_flow(page, names: List[str]):
    search_box = None
    for patt in [r"Search", r"Find", r"Filter", r"Look up"]:
        try:
            search_box = page.get_by_role("textbox", name=re.compile(patt, re.I)).first
            break
        except Exception:
            try:
                cand = page.get_by_placeholder(re.compile(patt, re.I)).first
                _ = cand
                search_box = cand
                break
            except Exception:
                pass
    if not search_box:
        search_box = page.locator("input[type='text']").first

    def click_middle_list_item(full_name: str):
        try:
            opt = page.get_by_role("option", name=re.compile(rf"^{re.escape(full_name)}$", re.I)).first
            click_hard(page, opt, f"option:{full_name}")
            return True
        except Exception:
            pass
        list_candidates = page.locator("ul,ol,div[role='listbox'],div[aria-label*='Available'],div:has(> div[role='option'])")
        count = list_candidates.count()
        for i in range(min(count, 6)):
            cont = list_candidates.nth(i)
            try:
                item = cont.get_by_text(re.compile(rf"^{re.escape(full_name)}$", re.I)).first
                click_hard(page, item, f"middle-list:{full_name}")
                return True
            except Exception:
                continue
        try:
            item = page.get_by_text(re.compile(rf"^{re.escape(full_name)}$", re.I)).first
            click_hard(page, item, f"text:{full_name}")
            return True
        except Exception:
            return False

    for name in names:
        debug(f"adding user -> {name}")
        try:
            scroll_into_view_safe(page, search_box)
            search_box.click(timeout=3000)
            try:
                search_box.fill("")
            except Exception:
                pass
            search_box.type(name, delay=40)
            time.sleep(0.4)

            if not click_middle_list_item(name):
                warn(f"couldn't find middle-list entry for: {name}")
            else:
                try:
                    click_button_like(page, r"Add\b")
                except Exception:
                    pass
        except Exception as e:
            warn(f"error adding {name}: {e}")

    try:
        click_button_like(page, "Continue")
    except Exception:
        click_button_like(page, "Next")
    page.wait_for_load_state("networkidle", timeout=15000)
    debug("add users complete, returned to form")

# ------------------------------
# Submit Form
# ------------------------------
def submit_form(page):
    for label in ["Submit", "Save", "Finish"]:
        try:
            click_button_like(page, label)
            page.wait_for_load_state("networkidle", timeout=15000)
            debug(f"submitted via '{label}'")
            return
        except Exception:
            continue
    warn("couldn't find a Submit/Save/Finish button")

# ------------------------------
# Dashboard navigation helper
# ------------------------------
def click_dashboard_to_form(page, dash_selector: str):
    debug(f"attempting dashboard click -> selector: {dash_selector}")
    try:
        el = page.locator(f"css={dash_selector}").first
        scroll_into_view_safe(page, el)
        el.click(timeout=8000)
    except Exception as e:
        warn(f"dashboard click failed: {e}")
        raise
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    debug("dashboard click done; waiting for form hallmark...")
    try:
        detect_form_ready(page)
    except Exception:
        warn("form hallmark not detected after dashboard click")

# ------------------------------
# Main Flow
# ------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--headed", action="store_true", help="Run headed")
    ap.add_argument("--login-url", required=True)
    ap.add_argument("--form-url", required=False, help="Optional: direct URL to the training form")
    ap.add_argument("--dash-click-selector", default=os.environ.get("DASH_CLICK_SELECTOR"), help="CSS selector to click on the dashboard to reach the form")
    ap.add_argument("--user", default=os.environ.get("VS_USER"))
    ap.add_argument("--password", default=os.environ.get("VS_PASS"))
    ap.add_argument("--csv-path", default=os.environ.get("TRAININGS_CSV", DEFAULT_CSV))
    ap.add_argument("--roster-url", default=os.environ.get("ROSTER_URL"))
    ap.add_argument("--roster-basic-user", default=os.environ.get("ROSTER_BASIC_USER"))
    ap.add_argument("--roster-basic-pass", default=os.environ.get("ROSTER_BASIC_PASS"))
    ap.add_argument("--skip-roster", action="store_true", help="Skip Add Users portion")
    args = ap.parse_args()

    training = read_training_row(args.csv_path)
    debug(f"training row: {training}")

    roster_names: List[str] = []
    if not args.skip_roster and args.roster_url and args.roster_basic_user and args.roster_basic_pass:
        try:
            roster_names = get_roster_from_minicad(args.roster_url, args.roster_basic_user, args.roster_basic_pass)
            debug(f"roster names count: {len(roster_names)}")
        except Exception as e:
            warn(f"Roster fetch failed (continuing without): {e}")
    else:
        debug("roster not requested or missing creds; skipping Add Users")

    with sync_playwright() as pw:
        browser, context = browser_ctx(pw, args.headed)
        page = context.new_page()

        # Try reuse auth state if it exists
        if os.path.exists(AUTH_STATE):
            try:
                context.close()
                browser.close()
                browser = pw.chromium.launch(headless=not args.headed)
                context = browser.new_context(storage_state=AUTH_STATE, ignore_https_errors=True, viewport={"width": 1440, "height": 900})
                context.set_default_timeout(15000)
                page = context.new_page()
                debug("reused existing auth_state")
            except Exception:
                pass

        # Login flow
        goto_robust(page, args.login_url, "login-url")
        try:
            scripted_login(page, args.user, args.password)
        except Exception as e:
            warn(f"scripted_login failed: {e}")

        # Persist cookies/session
        try:
            context.storage_state(path=AUTH_STATE)
            debug(f"auth state saved -> {AUTH_STATE}")
        except Exception:
            pass

        # Reach the form either via dashboard click or direct URL
        if args.dash_click_selector:
            try:
                click_dashboard_to_form(page, args.dash_click_selector)
            except Exception:
                warn("dashboard selector path failed; trying form-url")
                if args.form_url:
                    goto_robust(page, args.form_url, "form-url")
        elif args.form_url:
            goto_robust(page, args.form_url, "form-url")

        detect_form_ready(page)

        # Fill form from CSV data
        fill_training_form(page, training)

        # Save and Add Users (optional)
        if roster_names:
            try:
                click_button_like(page, r"Save and Add Users")
            except Exception:
                click_button_like(page, r"Add Users")
            page.wait_for_load_state("networkidle", timeout=15000)
            add_users_flow(page, roster_names)

        # Submit
        submit_form(page)

        log("DONE")
        try:
            context.storage_state(path=AUTH_STATE)
        except Exception:
            pass

        try:
            context.close()
            browser.close()
        except Exception:
            pass

if __name__ == "__main__":
    main()
