# bootstrap_login.py
from playwright.sync_api import sync_playwright
import os

PROFILE = os.environ.get("PLAYWRIGHT_PROFILE", "/home/training-bot/playwright-profiles/training-profile")
LOGIN_URL = os.environ.get("LOGIN_URL", "https://app.targetsolutions.com")

with sync_playwright() as pw:
    ctx = pw.chromium.launch_persistent_context(
        user_data_dir=PROFILE,
        headless=False,          # MUST be False so you can log in
        ignore_https_errors=True
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    print("Log in fully, get to the dashboard, then press Enter here...")
    input()
    # Persist to disk explicitly so we know cookies are written
    ctx.storage_state(path=os.path.join(PROFILE, "state.json"))
    ctx.close()
