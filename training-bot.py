import os
import time
import traceback
import sys
import re
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

PROFILE = os.environ.get(
    "PLAYWRIGHT_PROFILE", "/home/training-bot/playwright-profiles/training-profile"
).strip()
FORM_URL = os.environ.get("FORM_URL", "https://app.targetsolutions.com/").strip()

HEADLESS = os.environ.get("HEADLESS", "1") not in ("0", "false", "False")
SLOW_MO_MS = int(os.environ.get("SLOW_MO_MS", "0"))
DEFAULT_TIMEOUT = int(os.environ.get("PW_TIMEOUT_MS", "20000"))

CHECKBOX_LABEL = os.environ.get("TRAINING_CHECKBOX_LABEL", "Hands-on training").strip()
DESCRIPTION_LABEL = os.environ.get(
    "DESCRIPTION_LABEL", "Description of Training"
).strip()
DESCRIPTION_TEXT = os.environ.get(
    "DESCRIPTION_TEXT", "Pump ops drill: 1500 GPM, relay ops, safety briefing."
).strip()


def log(msg: str):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def is_login_page(page):
    url = page.url.lower()
    if "login" in url or "signin" in url or "auth" in url:
        return True
    # heuristic: username/password fields visible
    if page.locator("input[type='password']").first.is_visible():
        return True
    if page.get_by_label(re.compile("username|email", re.I)).count() > 0:
        return True
    return False


def assert_authenticated(page):
    if is_login_page(page):
        raise RuntimeError(
            f"Not authenticated. Landed on login: {page.url}\n"
            f"Make sure cookies were seeded with the SAME Playwright profile at {PROFILE}"
        )


def check_checkbox_by_label(scope, label_text):
    try:
        scope.get_by_role("checkbox", name=label_text, exact=False).check()
        return
    except BaseException:
        pass
    try:
        scope.get_by_label(label_text, exact=False).check()
        return
    except BaseException:
        pass
    scope.get_by_text(label_text, exact=False).first.click()


def fill_description(scope, label_or_placeholder, text):
    for fn in (
        lambda: scope.get_by_role(
            "textbox", name=label_or_placeholder, exact=False
        ).fill(text),
        lambda: scope.get_by_label(label_or_placeholder, exact=False).fill(text),
        lambda: scope.get_by_placeholder(label_or_placeholder, exact=False).fill(text),
        lambda: scope.locator("[contenteditable='true']").first.fill(text),
    ):
        try:
            fn()
            return
        except BaseException:
            pass
    raise RuntimeError(
        f"Could not find description field via hint: {
            label_or_placeholder!r}"
    )

def ts_login_if_needed(page):
    # If on login, actually log in
    if "login" in page.url.lower() or page.locator("input[type='password']").count():
        user = os.environ.get("TS_USER", "").strip()
        pwd  = os.environ.get("TS_PASS", "").strip()
        if not user or not pwd:
            raise RuntimeError("TS_USER/TS_PASS not set and we are on the login page.")

        log("On login page. Filling credentials…")
        # Try common selectors in order
        filled = False
        for fill_user in (
            lambda: page.get_by_label("Username", exact=False).fill(user),
            lambda: page.get_by_placeholder("Username", exact=False).fill(user),
            lambda: page.locator("input[name='username']").fill(user),
            lambda: page.locator("input[type='email']").fill(user),
        ):
            try: fill_user(); filled = True; break
            except: pass
        if not filled: raise RuntimeError("Could not locate username field.")

        filled = False
        for fill_pass in (
            lambda: page.get_by_label("Password", exact=False).fill(pwd),
            lambda: page.get_by_placeholder("Password", exact=False).fill(pwd),
            lambda: page.locator("input[name='password']").fill(pwd),
            lambda: page.locator("input[type='password']").fill(pwd),
        ):
            try: fill_pass(); filled = True; break
            except: pass
        if not filled: raise RuntimeError("Could not locate password field.")

        # Click Login / Sign In
        clicked = False
        for click_login in (
            lambda: page.get_by_role("button", name=re.compile(r"sign in|log ?in|login", re.I)).click(),
            lambda: page.locator("button[type='submit']").click(),
            lambda: page.locator("input[type='submit']").click(),
        ):
            try: click_login(); clicked = True; break
            except: pass
        if not clicked: raise RuntimeError("Could not find login submit button.")

        page.wait_for_load_state("networkidle")
        time.sleep(0.5)
        log(f"Post-login URL: {page.url}")
        if "login" in page.url.lower() or page.locator("input[type='password']").count():
            raise RuntimeError("Login appeared to fail or 2FA required.")


def main():
    log("Booting Playwright…")
    with sync_playwright() as pw:
        STATE = os.path.join(PROFILE, "state.json")
        log(f"Using state.json at: {STATE}")
        log(f"HEADLESS={HEADLESS}")

        browser = pw.chromium.launch(headless=HEADLESS)
        log("Launched browser")

        context = browser.new_context(storage_state=STATE, ignore_https_errors=True)
        context.set_default_timeout(DEFAULT_TIMEOUT)
        log("Created context from storage_state")

        page = context.new_page()
        log("New page created")

        # 1) Go to dashboard first
        log("Navigating to dashboard…")
        page.goto("https://app.targetsolutions.com/", wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        log(f"Dashboard URL after nav: {page.url}")

        # 2) Try scripted login if we're on the login page
        log(f"TS_USER set? [{'yes' if os.environ.get('TS_USER') else 'no'}]")
        log("Calling ts_login_if_needed()…")
        ts_login_if_needed(page)   # <-- THIS MUST RUN BEFORE ANY RAISE

        # 3) If we're still on login AFTER the helper, then bail
        if "login" in page.url.lower() or page.locator("input[type='password']").count():
            raise RuntimeError(f"Auth failed after scripted login attempt: {page.url}")

        # 4) Now hit the form
        log("Navigating to form URL…")
        page.goto(FORM_URL, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle")
        log(f"Form URL after nav: {page.url}")

        # 3) Actions
        log(f"Checking checkbox: {CHECKBOX_LABEL!r}")
        check_checkbox_by_label(scope, CHECKBOX_LABEL)

        log(f"Filling description using hint {DESCRIPTION_LABEL!r}")
        fill_description(scope, DESCRIPTION_LABEL, DESCRIPTION_TEXT)

        try:
            log("Clicking submit/save…")
            scope.get_by_role("button", name=re.compile(r"submit|save", re.I)).click()
            log("Submitted form.")
        except PWTimeout:
            log("No submit button found. Stopping before submit.")

        if not HEADLESS:
            log("Headful run: pausing 2s for inspection")
            time.sleep(2)

        log("Closing context/browser")
        context.close()
        browser.close()
        log("Done.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"FATAL: {e}")
        traceback.print_exc()
        sys.exit(1)