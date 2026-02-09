"""
VERSION 15 - MANUAL WORKAROUND
For sites with extreme bot protection (Etsy, Next.co.uk).

HOW IT WORKS:
1. You visit the URLs manually in your browser
2. Right-click the product image → "Copy image address"
3. Paste the direct image URLs into a new Excel column
4. This script uses those direct image URLs

This bypasses ALL bot detection because you're using the actual CDN URLs.
"""

import streamlit as st
import pandas as pd
from PIL import Image
from io import BytesIO
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from pptx import Presentation
from pptx.util import Inches
from bs4 import BeautifulSoup
import time
import re
import requests

# --- CONFIGURATION ---
TARGET_WIDTH_PX = 179
TARGET_HEIGHT_PX = 135

def download_direct_image(image_url, width, height):
    """
    Download image from direct URL (no scraping needed)
    """
    try:
        print(f"Downloading: {image_url[:80]}...")
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8',
        }
        
        r = requests.get(image_url, headers=headers, timeout=20)
        r.raise_for_status()
        
        img = Image.open(BytesIO(r.content))
        
        if img.mode in ("RGBA", "P", "LA"):
            img = img.convert("RGB")
        
        img = img.resize((width, height), Image.Resampling.LANCZOS)
        
        print(f"✅ Success")
        return img
        
    except Exception as e:
        print(f"❌ Failed: {e}")
        return None

def try_scrape_fallback(url):
    """
    Try basic scraping for other sites (not Etsy/Next)
    """
    try:
        if "etsy.com" in url or "next.co.uk" in url:
            return None  # Don't even try - they're too protected
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        r = requests.get(url, headers=headers, timeout=10)
        
        if r.status_code == 200:
            soup = BeautifulSoup(r.content, 'html.parser')
            
            # Try meta tags
            for prop in ['og:image', 'twitter:image']:
                meta = soup.find("meta", property=prop)
                if meta and meta.get("content"):
                    return meta["content"]
    except:
        pass
    
    return None

# --- STREAMLIT UI ---
st.set_page_config(page_title="Image Automator v15", layout="wide")
st.title("📊 Excel & PPT Image Automator v15")

st.markdown("""
## 🎯 Solution for Heavily Protected Sites (Etsy, Next.co.uk)

These sites have **extreme bot protection** that blocks all automated scraping. 

### ✅ **Working Solution - Manual Image URLs:**

**Step 1:** In your Excel, add a new column called "Image URL"

**Step 2:** For each product:
- Visit the URL in your browser
- Right-click the main product image
- Select **"Copy image address"** (or "Copy image URL")
- Paste into the "Image URL" column

**Step 3:** Upload your Excel and select "Image URL" as the source column

---

### 📋 **Example URLs to Copy:**

**Etsy format:**
```
https://i.etsystatic.com/12345678/r/il_fullxfull.1234567890_abcd.jpg
```

**Next.co.uk format:**
```
https://xcdn.next.co.uk/common/items/default/default/itemimages/AltItemZoom/y00128.jpg
```

---

### 🔧 **For Other Sites:**
This tool will still try to auto-scrape sites with weaker protection (Pottery Barn, Amazon, etc.)
""")

st.info("💡 **Pro Tip**: Use Excel's autofill or a browser extension like 'Image Downloader' to speed up collecting image URLs.")

uploaded_file = st.file_uploader("Upload Excel Template", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### Data Preview")
    st.dataframe(df.head())
    
    # Show column selection
    col1, col2 = st.columns(2)
    with col1:
        source_col = st.selectbox(
            "Source Column", 
            df.columns,
            help="Select 'Image URL' if you manually copied direct image links, or a URL column to try auto-scraping"
        )
    with col2:
        target_col_letter = st.text_input("Output Column (Excel)", value="B").upper()
    
    # Detect if user is using direct image URLs
    sample_value = str(df[source_col].iloc[0]) if len(df) > 0 else ""
    is_direct_images = any(x in sample_value.lower() for x in ['.jpg', '.png', '.jpeg', 'etsystatic', 'xcdn'])
    
    if is_direct_images:
        st.success("✅ Detected direct image URLs! This will work great.")
    else:
        st.warning("⚠️ Detected page URLs. Auto-scraping will be attempted but may fail for Etsy/Next.co.uk. Consider adding direct image URLs for 100% success.")
    
    if st.button("🔍 Test First 3 URLs"):
        st.write("### Testing...")
        
        success_count = 0
        for i, row in df.head(3).iterrows():
            value = row[source_col]
            
            if pd.notna(value):
                value_str = str(value).strip()
                
                with st.spinner(f"Testing: {value_str[:50]}..."):
                    img = None
                    
                    # Check if it's a direct image URL
                    if value_str.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp')) or 'etsystatic' in value_str or 'xcdn' in value_str:
                        # Direct image URL
                        img = download_direct_image(value_str, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                    else:
                        # Try scraping (will fail for Etsy/Next)
                        image_url = try_scrape_fallback(value_str)
                        if image_url:
                            img = download_direct_image(image_url, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                    
                    if img:
                        success_count += 1
                        st.success(f"✅ Success")
                        st.image(img, caption=value_str[:60], width=200)
                    else:
                        st.error(f"❌ Failed: {value_str[:60]}")
                        st.info("💡 For Etsy/Next.co.uk: Right-click product image → Copy image address → Paste in Excel")
        
        st.info(f"**Results: {success_count}/3 successful**")
    
    if st.button("Generate Excel"):
        progress = st.progress(0)
        status = st.empty()
        
        uploaded_file.seek(0)
        wb = openpyxl.load_workbook(uploaded_file)
        ws = wb.active
        
        count = 0
        failed = []
        total = len(df)
        
        for i, row in df.iterrows():
            progress.progress((i + 1) / total)
            value = row[source_col]
            
            if pd.notna(value):
                value_str = str(value).strip()
                status.text(f"Processing {i+1}/{total}: {value_str[:40]}...")
                
                img = None
                
                # Direct image URL
                if value_str.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp')) or 'etsystatic' in value_str or 'xcdn' in value_str:
                    img = download_direct_image(value_str, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                else:
                    # Try scraping
                    image_url = try_scrape_fallback(value_str)
                    if image_url:
                        img = download_direct_image(image_url, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                
                if img:
                    count += 1
                    buf = BytesIO()
                    img.save(buf, format='PNG')
                    buf.seek(0)
                    excel_img = OpenpyxlImage(buf)
                    ws.add_image(excel_img, f"{target_col_letter}{i+2}")
                    ws.row_dimensions[i+2].height = 105
                else:
                    failed.append(i+2)  # Row number in Excel
                    
            time.sleep(0.5)
        
        if failed:
            status.warning(f"⚠️ Complete! {count}/{total} images added. Failed rows: {', '.join(map(str, failed))}")
        else:
            status.success(f"✅ Complete! {count}/{total} images added")
        
        out = BytesIO()
        wb.save(out)
        out.seek(0)
        
        st.download_button("📥 Download Excel", out, "output_v15.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        if failed:
            st.info(f"""
            **Failed rows:** {', '.join(map(str, failed))}
            
            **To fix:**
            1. Visit those URLs manually in your browser
            2. Right-click the product image → "Copy image address"
            3. Paste those direct image URLs into Excel
            4. Re-run this tool
            """)
    
    if st.button("Generate PowerPoint"):
        with st.spinner("Generating slides..."):
            prs = Presentation()
            blank = prs.slide_layouts[6]
            count = 0
            
            for i, row in df.iterrows():
                value = row[source_col]
                
                if pd.notna(value):
                    value_str = str(value).strip()
                    img = None
                    
                    if value_str.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp')) or 'etsystatic' in value_str or 'xcdn' in value_str:
                        img = download_direct_image(value_str, 800, 600)
                    else:
                        image_url = try_scrape_fallback(value_str)
                        if image_url:
                            img = download_direct_image(image_url, 800, 600)
                    
                    if img:
                        count += 1
                        slide = prs.slides.add_slide(blank)
                        buf = BytesIO()
                        img.save(buf, format='PNG')
                        buf.seek(0)
                        slide.shapes.add_picture(buf, Inches(1), Inches(1), width=Inches(8))
                        
                        tx = slide.shapes.add_textbox(Inches(1), Inches(6.5), Inches(8), Inches(0.5))
                        tx.text_frame.text = f"Source: {value_str}"
            
            out = BytesIO()
            prs.save(out)
            out.seek(0)
            st.download_button("📥 Download PowerPoint", out, "output_v15.pptx",
                              mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            st.success(f"✅ Created {count} slides")

st.markdown("---")
st.markdown("""
### 📖 **Quick Guide: How to Copy Image URLs**

**Chrome/Edge:**
1. Go to product page
2. Right-click main image
3. Select "Copy image address"

**Firefox:**
1. Go to product page
2. Right-click main image
3. Select "Copy Image Link"

**Safari:**
1. Go to product page
2. Right-click main image
3. Select "Copy Image Address"

Paste these URLs directly into your Excel sheet!
""")
