"""Render the VWC Persona GTM Strategy HTML to PDF via Playwright.

Mirrors the toolchain used for VWC_Proposal.pdf in the SEO repo.
Run:  python _generate_pdf.py
"""
import os
from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, 'VWC_Persona_GTM_Strategy.html')
PDF = os.path.join(HERE, 'VWC_Persona_GTM_Strategy.pdf')


def build():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(f"file:///{HTML.replace(os.sep, '/')}")
        page.wait_for_load_state('networkidle')
        page.pdf(path=PDF, format='Letter', print_background=True,
                 margin={'top': '0', 'right': '0', 'bottom': '0', 'left': '0'})
        browser.close()
    kb = os.path.getsize(PDF) / 1024
    print(f"Generated: {PDF} ({kb:.0f} KB)")


if __name__ == '__main__':
    build()
