#!/usr/bin/env python3
from playwright.sync_api import sync_playwright

def main():
    print("[diag] starting Playwright smoke test...", flush=True)
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        ctx = browser.new_context(viewport={"width": 1200, "height": 800})
        page = ctx.new_page()
        page.goto("https://example.org", wait_until="domcontentloaded", timeout=30000)
        print("[diag] page title:", page.title(), flush=True)
        ctx.close()
        browser.close()
    print("[diag] success.", flush=True)

if __name__ == "__main__":
    main()
