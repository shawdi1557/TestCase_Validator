import openpyxl
from playwright.sync_api import sync_playwright

print("openpyxl ok")

with sync_playwright() as p:
    b = p.chromium.launch()
    b.close()
    print("chromium ok")