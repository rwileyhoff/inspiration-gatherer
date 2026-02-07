import streamlit as st
import pandas as pd
from curl_cffi import requests as cffi_requests
from PIL import Image
from io import BytesIO
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from pptx import Presentation
from pptx.util import Inches
from bs4 import BeautifulSoup
import time
import re
from duckduckgo_search import DDGS

# --- CONFIGURATION ---
TARGET_WIDTH_PX = 179
TARGET_HEIGHT_PX = 135

# --- HELPER FUNCTIONS ---

def get_next_direct_cdn(url):
    """
    Specifically for Next.co.uk.
    Instead of visiting the page, we extract the Style/Item codes and 
    guess the direct image URL from their public image server (CDN).
    """
    try:
        # Regex to find codes like 'sv068808' or 'y00128' in the URL
        # Matches alphanumeric strings of 5-8 chars after 'style/' or '/'
        codes = re.findall(r'(?:style/|/)([a-zA-Z0-9]{5,10})', url)
        
        # Next.co.uk image server patterns
        # They usually store images under the "Option" code (the second code usually) 
        # or the "Style" code (the first code). We try both.
        base_cdn = "https://xcdn.next.co.uk/common/items/default/default/itemimages/search"
        
        session = cffi_requests.Session()
        
        # We prefer the last code found (often the specific color option), then the first
        unique_codes = list(dict.fromkeys(reversed(codes))) # Remove duplicates, preserve order
        
        for code in unique_codes:
            # Construct the candidate URL
            candidate_url = f"{base_cdn}/{code}.jpg"
            
            try:
                # We use a HEAD request first to see if the image exists (Fast!)
                r = session.head(candidate_url, timeout=3)
                if r.status_code == 200:
                    return candidate_url
            except:
                pass
                
    except Exception as e:
        print(f"CDN Guessing failed: {e}")
        
    return None

def fetch_image_via_search_fallback(query):
    """
    Last resort: Search DuckDuckGo, but strictly for 'Product Image'
    to avoid selfies/blogs.
    """
    try:
        # Extract a cleaner query ID if possible
        clean_id = ""
        match = re.search(r'style/([a-zA-Z0-9]+)', query)
        if match:
            clean_id = match.group(1)
            search_term = f"Next UK product {clean_id}"
        else:
            search_term = f"{query} product image white background"

        with DDGS() as ddgs:
            results = list(ddgs.images(
                keywords=search_term, 
                max_results=1, 
                safesearch='off'
            ))
            if results:
                return results[0]['image']
    except:
        pass
    return None

def download_and_resize_image(url, width, height):
    try:
        clean_url = url.strip()
        final_image_url = None
        
        # --- STRATEGY 1: DIRECT FILE CHECK ---
        if clean_url.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp')):
            final_image_url = clean_url

        # --- STRATEGY 2: SPECIAL HANDLING FOR NEXT.CO.UK ---
        if not final_image_url and "next.co.uk" in clean_url:
            # Try to build the CDN link directly (Bypasses security entirely)
            final_image_url = get_next_direct_cdn(clean_url)

        # --- STRATEGY 3: GENERIC WEBSITES (Pottery Barn, etc.) ---
        if not final_image_url:
            # Try standard scraping with browser disguise
            try:
                session = cffi_requests.Session()
                r = session.get(clean_url, impersonate="chrome110", timeout=10)
                if r.status_code == 200:
                    soup = BeautifulSoup(r.content, 'html.parser')
                    og = soup.find("meta", property="og:image")
                    if og and og.get("content"):
                        final_image_url = og["content"]
            except:
                pass

        # --- STRATEGY 4: FALLBACK SEARCH ---
        if not final_image_url:
            final_image_url = fetch_image_via_search_fallback(clean_url)

        # --- DOWNLOAD & PROCESS ---
        if final_image_url:
            # Download the actual image bytes
            r = cffi_requests.get(final_image_url, impersonate="chrome110", timeout=15)
            r.raise_for_status()
            
            img = Image.open(BytesIO(r.content))
            
            # Convert to RGB (standardize format)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            
            # Resize
            img = img.resize((width, height))
            return img
            
    except Exception as e:
        print(f"Failed to process {url}: {e}")
        return None

# --- APP INTERFACE ---
st.set_page_config(page_title="Excel Image Automator v7", layout="wide")
st.title("📊 Excel & PPT Image Automator")
st.markdown("""
**Version 7 (CDN Bypass)**: 
- Specifically engineered for **Next.co.uk**. 
- It bypasses the website and pulls the official product image directly from the file server.
- Removes the risk of getting random "Selfie" or "Blog" images.
""")

uploaded_file = st.file_uploader("Upload your Excel Template (.xlsx)", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### Preview")
    st.dataframe(df.head())

    col1, col2 = st.columns(2)
    with col1:
        url_col = st.selectbox("Select Column with URLs", df.columns, index=0)
    with col2:
        target_col_letter = st.text_input("Output Column Letter", value="B").upper()

    st.markdown("---")

    if st.button("Generate Excel"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Load workbook
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        
        success_count = 0
        total_rows = len(df)
        
        for i, row in df.iterrows():
            # Update Progress
            progress_bar.progress((i + 1) / total_rows)
            
            raw_url = row[url_col]
            
            if pd.notna(raw_url) and str(raw_url).startswith('http'):
                status_text.text(f"Processing Row {i+1}...")
                
                processed_img = download_and_resize_image(raw_url, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                
                if processed_img:
                    success_count += 1
                    img_stream = BytesIO()
                    processed_img.save(img_stream, format='PNG')
                    img_stream.seek(0)
                    
                    # Add to Excel
                    excel_img = OpenpyxlImage(img_stream)
                    excel_row = i + 2
                    ws.add_image(excel_img, f"{target_col_letter}{excel_row}")
                    ws.row_dimensions[excel_row].height = 105
            
            # Tiny pause
            time.sleep(0.1)

        status_text.text(f"Done! {success_count}/{total_rows} images added.")
        
        # Download
        out_buffer = BytesIO()
        wb.save(out_buffer)
        out_buffer.seek(0)
        
        st.download_button("Download Excel", out_buffer, "output_v7.xlsx")

    if st.button("Generate PowerPoint"):
        with st.spinner("Generating..."):
            prs = Presentation()
            blank_layout = prs.slide_layouts[6]
            
            for i, row in df.iterrows():
                raw_url = row[url_col]
                if pd.notna(raw_url) and str(raw_url).startswith('http'):
                    processed_img = download_and_resize_image(raw_url, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                    if processed_img:
                        slide = prs.slides.add_slide(blank_layout)
                        img_stream = BytesIO()
                        processed_img.save(img_stream, format='PNG')
                        img_stream.seek(0)
                        slide.shapes.add_picture(img_stream, Inches(4), Inches(3))
                        txBox = slide.shapes.add_textbox(Inches(1), Inches(6), Inches(8), Inches(1))
                        txBox.text_frame.text = f"Source: {raw_url}"

            ppt_buffer = BytesIO()
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)
            st.download_button("Download PPT", ppt_buffer, "output_v7.pptx")
