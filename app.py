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

# --- CONFIGURATION ---
TARGET_WIDTH_PX = 179
TARGET_HEIGHT_PX = 135

def get_real_image_url(url):
    """
    Fetches the URL using a browser impersonator to bypass 403 errors.
    Parses the HTML to find the high-quality Open Graph image.
    """
    try:
        # 1. Clean the URL
        clean_url = url.strip()
        
        # 2. Check if it's already an image file
        if clean_url.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp')):
            return clean_url

        # 3. Request the page masquerading as a real Chrome browser
        # 'impersonate="chrome"' generates the correct TLS fingerprint to fool security systems
        response = cffi_requests.get(clean_url, impersonate="chrome", timeout=15)
        
        # If we still get a 403, we return None (it will be logged later)
        if response.status_code != 200:
            print(f"Failed to access {clean_url} - Status: {response.status_code}")
            return None

        # 4. Parse HTML to find the "og:image" (The main product image used for social sharing)
        soup = BeautifulSoup(response.content, 'html.parser')
        og_image = soup.find("meta", property="og:image")
        
        if og_image and og_image.get("content"):
            return og_image["content"]
            
    except Exception as e:
        print(f"Error scraping page {url}: {e}")
        pass
    
    # If scraping failed, return original URL (it might work directly if we are lucky)
    return url

def download_and_resize_image(url, width, height):
    try:
        # 1. Get the direct image link
        image_url = get_real_image_url(url)
        
        if not image_url:
            return None
        
        # 2. Download the actual image using the same browser impersonation
        response = cffi_requests.get(image_url, impersonate="chrome", timeout=15)
        response.raise_for_status()
        
        # 3. Process with Pillow
        img = Image.open(BytesIO(response.content))
        
        # Convert to RGB (fixes issues with transparent PNGs or CMYK images)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
            
        img = img.resize((width, height))
        return img
    except Exception as e:
        print(f"Error downloading image {url}: {e}")
        return None

# --- APP INTERFACE ---
st.set_page_config(page_title="Excel Image Automator v3", layout="wide")
st.title("📊 Excel & PPT Image Automator")
st.markdown("""
**Version 3 Update**: Now using `curl_cffi` to bypass "403 Forbidden" errors on sites like Pottery Barn.
""")

uploaded_file = st.file_uploader("Upload your Excel Template (.xlsx)", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### 1. Preview Data")
    st.dataframe(df.head())

    col1, col2 = st.columns(2)
    with col1:
        # User selects which column has the links
        url_col = st.selectbox("Select Column with URLs", df.columns, index=0)
    with col2:
        # User types which column letter gets the image
        target_col_letter = st.text_input("Output Column Letter for Image", value="B").upper()

    st.markdown("---")
    st.write("### 2. Generate Files")

    if st.button("Generate Excel with Images"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Prepare Excel File
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        
        total_rows = len(df)
        success_count = 0
        
        for i, row in df.iterrows():
            # Update GUI
            progress = (i + 1) / total_rows
            progress_bar.progress(progress)
            
            raw_url = row[url_col]
            
            if pd.notna(raw_url) and str(raw_url).startswith('http'):
                status_text.text(f"Processing ({i+1}/{total_rows}): {str(raw_url)[:40]}...")
                
                # The Magic: Download & Resize
                processed_img = download_and_resize_image(raw_url, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                
                if processed_img:
                    success_count += 1
                    
                    # Convert processed image to bytes for Excel
                    img_stream = BytesIO()
                    processed_img.save(img_stream, format='PNG')
                    img_stream.seek(0)
                    
                    # Place Image in Excel
                    excel_img = OpenpyxlImage(img_stream)
                    excel_row = i + 2 # Header is 1, Data starts at 2
                    cell_address = f"{target_col_letter}{excel_row}"
                    
                    excel_img.anchor = cell_address
                    ws.add_image(excel_img)
                    
                    # Set Row Height
                    ws.row_dimensions[excel_row].height = 105
            
            # Tiny sleep to prevent rate limiting
            time.sleep(0.5)

        status_text.text(f"Done! Processed {total_rows} rows. Images found: {success_count}.")
        
        # Save output
        excel_buffer = BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        if success_count == 0:
            st.error("No images were added. If you saw '403' errors in the logs, the site might still be blocking the request.")
        else:
            st.success(f"Success! {success_count} images added.")
            st.download_button(
                label="Download Final Excel",
                data=excel_buffer,
                file_name="output_with_images.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    if st.button("Generate PowerPoint"):
        with st.spinner("Generating Slides..."):
            prs = Presentation()
            blank_slide_layout = prs.slide_layouts[6]
            
            for i, row in df.iterrows():
                raw_url = row[url_col]
                if pd.notna(raw_url) and str(raw_url).startswith('http'):
                    
                    processed_img = download_and_resize_image(raw_url, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                    
                    if processed_img:
                        slide = prs.slides.add_slide(blank_slide_layout)
                        
                        img_stream = BytesIO()
                        processed_img.save(img_stream, format='PNG')
                        img_stream.seek(0)
                        
                        left = Inches(4)
                        top = Inches(3)
                        slide.shapes.add_picture(img_stream, left, top)
                        
                        txBox = slide.shapes.add_textbox(Inches(1), Inches(6), Inches(8), Inches(1))
                        txBox.text_frame.text = f"Source: {raw_url}"

            ppt_buffer = BytesIO()
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)
            
            st.success("PowerPoint Generated!")
            st.download_button(
                label="Download PowerPoint",
                data=ppt_buffer,
                file_name="presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )