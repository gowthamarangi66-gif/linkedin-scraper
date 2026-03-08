import streamlit as st
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font
import random
import time
import re
import io
from datetime import datetime
from urllib.parse import quote_plus, urlparse, parse_qs, unquote

st.set_page_config(page_title="LinkedIn Scraper", page_icon="🔍", layout="centered")

st.title("🔍 LinkedIn Profile Scraper")
st.write("Finds LinkedIn profiles from Google and saves them to your Outreach Tracker.")

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36",
]

SEARCH_ENGINES = ["google", "bing"]

def get_headers():
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "DNT": "1",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Cache-Control": "max-age=0",
    }

def extract_name(title):
    title = re.sub(r"\s*[-|–—]\s*LinkedIn.*", "", title, flags=re.IGNORECASE)
    for sep in [" - ", " | ", " – ", " — "]:
        if sep in title:
            return title.split(sep)[0].strip()
    return title.strip() or "Unknown"

def clean_url(href):
    if not href:
        return None
    # Handle Google redirect URLs
    if "/url?" in href:
        parsed = urlparse(href)
        qs = parse_qs(parsed.query)
        if "q" in qs:
            href = qs["q"][0]
        elif "url" in qs:
            href = qs["url"][0]
    href = unquote(href)
    if "linkedin.com/in/" not in href:
        return None
    # Clean tracking params
    href = href.split("?")[0].rstrip("/")
    return href

def scrape_bing(query, max_results):
    profiles = []
    seen = set()
    offset = 0

    session = requests.Session()
    session.headers.update(get_headers())

    while len(profiles) < max_results:
        try:
            url = f"https://www.bing.com/search?q={quote_plus(query)}&first={offset}&count=10"
            resp = session.get(url, timeout=15)
            soup = BeautifulSoup(resp.text, "html.parser")

            results = soup.select("li.b_algo")
            found = 0

            for r in results:
                a = r.select_one("h2 a")
                if not a:
                    continue
                href = clean_url(a.get("href", ""))
                if not href or href in seen:
                    continue
                seen.add(href)

                title = a.get_text(strip=True)
                name = extract_name(title)
                profiles.append({"name": name, "url": href})
                found += 1
                if len(profiles) >= max_results:
                    break

            if found == 0:
                break
            offset += 10
            time.sleep(random.uniform(1.5, 3.5))

        except Exception as e:
            break

    return profiles

def scrape_google(query, max_results):
    profiles = []
    seen = set()
    start = 0

    session = requests.Session()
    session.headers.update(get_headers())

    while len(profiles) < max_results:
        try:
            url = f"https://www.google.com/search?q={quote_plus(query)}&start={start}&num=10&hl=en&gl=in"
            resp = session.get(url, timeout=15)

            if "captcha" in resp.text.lower() or "unusual traffic" in resp.text.lower():
                return profiles, "google_blocked"

            soup = BeautifulSoup(resp.text, "html.parser")

            # Try multiple selectors
            links = soup.select("a[href]")
            found = 0

            for a in links:
                href = clean_url(a.get("href", ""))
                if not href or href in seen:
                    continue
                seen.add(href)

                # Get name from parent or title
                parent = a.find_parent("div")
                h3 = a.find("h3") or (parent.find("h3") if parent else None)
                title = h3.get_text(strip=True) if h3 else a.get_text(strip=True)
                name = extract_name(title)

                profiles.append({"name": name, "url": href})
                found += 1
                if len(profiles) >= max_results:
                    break

            if found == 0:
                break
            start += 10
            time.sleep(random.uniform(2, 5))

        except Exception as e:
            break

    return profiles, None

def append_to_excel(excel_bytes, profiles):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    ws = wb["1 - Outreach Tracker"]

    existing = set()
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[2]:
            existing.add(str(row[2]).strip().rstrip("/"))

    added, skipped = 0, 0

    for p in profiles:
        clean = p["url"].rstrip("/")
        if clean in existing:
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

        existing.add(clean)
        added += 1

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), added, skipped


# --- UI ---
uploaded = st.file_uploader("📂 Step 1: Upload your Excel Tracker", type=["xlsx"])

keyword = st.text_input(
    "👤 Step 2: Who are you targeting?",
    placeholder='e.g. chartered accountant, HR consultant, fractional CFO'
)

location = st.text_input("📍 Location (optional)", value="India")

engine = st.radio("🔎 Search Engine", ["Bing (Recommended)", "Google"], horizontal=True)

max_results = st.slider("Max profiles", 5, 50, 15)

if keyword:
    full_query = f'site:linkedin.com/in "{keyword}" {location}'.strip()
    st.code(full_query, language=None)

if st.button("🚀 Scrape LinkedIn Profiles", use_container_width=True):
    if not uploaded:
        st.error("Please upload your Excel tracker first.")
    elif not keyword.strip():
        st.error("Please enter who you are targeting.")
    else:
        excel_bytes = uploaded.read()
        full_query = f'site:linkedin.com/in "{keyword}" {location}'.strip()

        with st.spinner(f"Searching for '{keyword}' profiles..."):
            if "Bing" in engine:
                profiles = scrape_bing(full_query, max_results)
                error = None
            else:
                profiles, error = scrape_google(full_query, max_results)
                if error == "google_blocked":
                    st.warning("Google blocked the search. Switching to Bing...")
                    profiles = scrape_bing(full_query, max_results)
                    error = None

        if not profiles:
            st.warning("No profiles found. Try a broader keyword or different location.")
            st.info("💡 Tips:\n- Remove quotes around keyword\n- Try just: chartered accountant India\n- Switch to the other search engine")
        else:
            updated_bytes, added, skipped = append_to_excel(excel_bytes, profiles)

            col1, col2, col3 = st.columns(3)
            col1.metric("✅ Found", len(profiles))
            col2.metric("➕ Added", added)
            col3.metric("⏭️ Skipped", skipped)

            st.subheader("Preview")
            for p in profiles:
                st.write(f"**{p['name']}** — [{p['url']}]({p['url']})")

            fname = uploaded.name.replace(".xlsx", f"_updated_{datetime.now().strftime('%d%b')}.xlsx")
            st.download_button(
                label="⬇️ Download Updated Outreach Tracker",
                data=updated_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
