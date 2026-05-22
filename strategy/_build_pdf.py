"""Generic: render a self-contained HTML report to a sibling PDF via Playwright.
Usage: python _build_pdf.py <file.html>   ->  <file.pdf>
"""
import os, sys
from playwright.sync_api import sync_playwright

html = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "VWC_AI_LeadGen_Development_Overview.html")
pdf = os.path.splitext(html)[0] + ".pdf"

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto(f"file:///{html.replace(os.sep, '/')}")
    page.wait_for_load_state("networkidle")
    # report any clipped fixed-height pages
    clips = page.eval_on_selector_all(".page", "els => els.map((e,i)=>({i:i+1, clip:e.scrollHeight-e.clientHeight}))")
    bad = [c["i"] for c in clips if c["clip"] > 2 and c["i"] != 1]  # page 1 cover footer is a known benign artifact
    page.pdf(path=pdf, format="Letter", print_background=True, margin={"top":"0","right":"0","bottom":"0","left":"0"})
    browser.close()

print(f"Generated: {pdf} ({os.path.getsize(pdf)/1024:.0f} KB)")
print("Clip check (excl. cover):", "ALL CLEAN" if not bad else f"** CLIPPED PAGES {bad}")
