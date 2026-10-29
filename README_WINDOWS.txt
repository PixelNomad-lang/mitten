MITTEN EXTRACTS MAP AUDIT — WINDOWS STARTER

1) Install Python 3.12 from https://www.python.org/downloads/windows/ and tick “Add python.exe to PATH”.
2) Extract this ZIP to a folder, e.g. Desktop\mitten_map_python_starter.
3) Copy your retailer Excel file into this folder and rename it exactly: Retailers_with_Websites.xlsx
4) Open the folder in File Explorer. Click the address bar, type cmd, press Enter.
5) Run these commands one at a time:
   py -m venv .venv
   .venv\Scripts\activate
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   python -m playwright install chromium
   python scanner.py
6) Results: MAP_Scan_Results.xlsx and screenshots in evidence\

WHAT IT DOES: scans each unique website homepage plus up to 5 relevant same-domain links, saves screenshots and candidate page/product matches.
IMPORTANT: It does not guarantee every product/banner is found. Candidate dollar amounts may belong to other products. It never confirms compliance or a violation automatically. Human verification is required.

If “py is not recognized”, reinstall Python and tick Add python.exe to PATH, then reopen cmd.
If a site blocks scanning, requires login, or shows promotions only in images/cart, review it manually.
If scanning takes too long, stop with Ctrl+C and change MAX_LINKS_PER_SITE in scanner.py from 5 to 2.
