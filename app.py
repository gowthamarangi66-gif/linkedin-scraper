import streamlit as st
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import random
import time
import re
import io
from datetime import datetime

st.set_page_config(page_title="LinkedIn Scraper", page_icon="linkedin", layout="centered")

st.title("LinkedIn Profile Scraper")
st.write("Search Google for LinkedIn profiles and save them to your Outreach Tracker.")

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
]

def get_headers():
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "DNT": "1",
    }

def extract_name(title):
    for sep in [" - ", " | ", " - ", " - "]:
        if sep in title:
            return title.split(sep)[0].strip()
    cleaned = re.sub(r"\s*[-|]\s*LinkedIn.*", "", title, flags=re.IGNORECASE)
    return cleaned.strip() or "Unknown"

def scrape_google(query, max_results):
    profiles = []
    seen = set()
    start = 0
    error_msg = None

    while len(profiles) < max_results:
        url = "https://www.google.com/search"
        params = {"q": query, "start": start, "num": 10, "hl": "en"}
        try:
            resp = requests.get(url, headers=get_headers(), params=params, timeout=15)
        except Exception as e:
            error_msg = str(e)
            break

        if "captcha" in resp.text.lower() or "unusual traffic" in resp.text.lower():
            error_msg = "Google CAPTCHA triggered. Please wait 10 minutes and try again."
            break

        soup = BeautifulSoup(resp.text, "html.parser")
        blocks = soup.select("div.g")
        if not blocks:
            blocks = soup.select("div.tF2Cxc")

        found = 0
        for block in blocks:
            a = block.select_one("a")
            if not a:
                continue
            href = a.get("href", "")
            if "linkedin.com/in/" not in href:
                continue
            href = href.split("?")[0].rstrip("/")
            if href in seen:
                continue
            seen.add(href)

            h3 = block.select_one("h3")
            name = extract_name(h3.get_text(strip=True)) if h3 else "Unknown"

            profiles.append({"name": name, "url": href})
            found += 1

            if len(profiles) >= max_results:
                break

        if found == 0:
            break

        start += 10
        time.sleep(random.uniform(2, 5))

    return profiles, error_msg

def append_to_excel(excel_bytes, profiles):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb["1 - Outreach Tracker"]

    existing = set()
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[2]:
            existing.add(str(row[2]).strip().rstrip("/"))

    added = 0
    skipped = 0

    for p in profiles:
        clean_url = p["url"].rstrip("/")
        if clean_url in existing:
            skipped += 1
            continue

        next_row = 3
        while ws.cell(row=next_row, column=1).value:
            next_row += 1

        ws.cell(row=next_row, column=1).value = p["name"]
        ws.cell(row=next_row, column=2).value = "LinkedIn"

        url_cell = ws.cell(row=next_row, column=3)
        url_cell.value = p["url"]
        url_cell.hyperlink = p["url"]
        url_cell.font = Font(color="0563C1", underline="single")

        ws.cell(row=next_row, column=4).value = ""
        ws.cell(row=next_row, column=8).value = "Not Contacted"

        existing.add(clean_url)
        added += 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), added, skipped


# --- UI ---

uploaded = st.file_uploader("Step 1: Upload your Excel Tracker", type=["xlsx"])

query = st.text_input(
    "Step 2: Enter your search query",
    value='site:linkedin.com/in "marketing manager" India'
)

st.caption('Tip: Try  site:linkedin.com/in "salon owner" India  or  site:linkedin.com/in "startup founder" Hyderabad')

max_results = st.slider("Max profiles to scrape", min_value=5, max_value=50, value=10)

if st.button("Scrape LinkedIn Profiles"):
    if not uploaded:
        st.error("Please upload your Excel tracker first.")
    elif not query.strip():
        st.error("Please enter a search query.")
    else:
        excel_bytes = uploaded.read()
        with st.spinner("Scraping Google for LinkedIn profiles..."):
            profiles, error = scrape_google(query, max_results)

        if error:
            st.error(error)
        elif not profiles:
            st.warning("No profiles found. Try a different search query.")
        else:
            updated_bytes, added, skipped = append_to_excel(excel_bytes, profiles)

            col1, col2, col3 = st.columns(3)
            col1.metric("Profiles Found", len(profiles))
            col2.metric("Added to Sheet", added)
            col3.metric("Duplicates Skipped", skipped)

            st.subheader("Results Preview")
            for p in profiles:
                st.write(f"**{p['name']}** — {p['url']}")

            fname = uploaded.name.replace(".xlsx", f"_updated_{datetime.now().strftime('%d%b')}.xlsx")
            st.download_button(
                label="Download Updated Outreach Tracker",
                data=updated_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
