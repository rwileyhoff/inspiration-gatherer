"""
VERSION 15.1 - DUAL COLUMN SUPPORT
Handles both 'URL' and 'IMAGE URL' columns intelligently:
- If IMAGE URL exists → use it (direct download)
- If IMAGE URL is empty → try auto-scraping the URL column
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
    """Download image from direct URL"""
    try:
        print(f"  Downloading: {image_url[:80]}...")
        
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
        
        print(f"  ✅ Success")
        return img
        
    except Exception as e:
        print(f"  ❌ Failed: {e}")
        return None

def try_scrape_page(url):
    """Try to scrape image from product page (for non-Etsy/Next sites)"""
    try:
        # Don't even try for heavily protected sites
        if "etsy.com" in url or "next.co.uk" in url:
            print(f"  ⚠️ {url.split('/')[2]} requires manual image URL")
            return None
        
        print(f"  Trying to scrape: {url[:60]}...")
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        r = requests.get(url, headers=headers, timeout=10)
        
        if r.status_code == 200:
            soup = BeautifulSoup(r.content, 'html.parser')
            
            # Try meta tags
            for prop in ['og:image', 'twitter:image', 'og:image:url']:
                meta = soup.find("meta", property=prop)
                if not meta:
                    meta = soup.find("meta", attrs={"name": prop})
                
                if meta and meta.get("content"):
                    img_url = meta["content"]
                    print(f"  ✓ Found via {prop}")
                    return img_url
    except Exception as e:
        print(f"  ❌ Scraping failed: {e}")
    
    return None

def process_row(row, url_col, image_url_col):
    """
    Process a single row - intelligently handles both columns
    
    Priority:
    1. If IMAGE URL exists and is valid → use it
    2. If IMAGE URL is empty → try scraping URL column
    """
    
    # Get values from both columns
    url_value = row.get(url_col) if url_col else None
    image_url_value = row.get(image_url_col) if image_url_col else None
    
    # Clean values
    url_value = str(url_value).strip() if pd.notna(url_value) else None
    image_url_value = str(image_url_value).strip() if pd.notna(image_url_value) else None
    
    final_image_url = None
    
    print(f"\n{'='*70}")
    
    # STRATEGY 1: Check if IMAGE URL column has a direct image link
    if image_url_value and image_url_value != 'nan':
        print(f"IMAGE URL provided: {image_url_value[:60]}...")
        
        # Verify it's actually an image URL
        if any(x in image_url_value.lower() for x in ['.jpg', '.jpeg', '.png', '.gif', '.webp', 'etsystatic', 'xcdn']):
            final_image_url = image_url_value
        else:
            print(f"  ⚠️ IMAGE URL doesn't look like an image, will try URL column")
    
    # STRATEGY 2: If no IMAGE URL, try scraping the URL column
    if not final_image_url and url_value and url_value != 'nan':
        print(f"No IMAGE URL, trying to scrape URL: {url_value[:60]}...")
        
        # Check if URL itself is a direct image
        if url_value.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.webp')):
            final_image_url = url_value
        else:
            # Try scraping the product page
            final_image_url = try_scrape_page(url_value)
    
    return final_image_url

# --- STREAMLIT UI ---
st.set_page_config(page_title="Image Automator v15.1", layout="wide")
st.title("📊 Excel & PPT Image Automator v15.1")

st.markdown("""
## 🎯 Dual Column Support

**This version intelligently handles both URL columns:**

### Column Setup:
- **URL** column: Product page URLs (e.g., `https://www.potterybarn.com/products/...`)
- **IMAGE URL** column: Direct image URLs (e.g., `https://i.etsystatic.com/.../fullxfull.jpg`)

### How It Works:
1. ✅ **If IMAGE URL exists** → Downloads directly (100% success)
2. ✅ **If IMAGE URL is empty** → Tries to auto-scrape URL column
3. ✅ **Mix both in same file** → Some rows manual, some auto

### Which Sites Need Manual IMAGE URLs:
- ❌ **Etsy** (copy image URL manually)
- ❌ **Next.co.uk** (copy image URL manually)
- ✅ **Pottery Barn, Target, West Elm, etc.** (auto-scrapes from URL)

---
""")

uploaded_file = st.file_uploader("Upload Excel Template", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("### Data Preview")
    st.dataframe(df.head())
    
    # Column detection
    columns = df.columns.tolist()
    
    # Try to auto-detect URL and IMAGE URL columns
    url_col_default = None
    image_url_col_default = None
    
    for col in columns:
        col_lower = col.lower()
        if 'image' in col_lower and 'url' in col_lower:
            image_url_col_default = col
        elif col_lower == 'url':
            url_col_default = col
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        url_col = st.selectbox(
            "URL Column (product pages)", 
            ['(None)'] + columns,
            index=columns.index(url_col_default) + 1 if url_col_default in columns else 0,
            help="Column with product page URLs (will try auto-scraping)"
        )
        if url_col == '(None)':
            url_col = None
    
    with col2:
        image_url_col = st.selectbox(
            "IMAGE URL Column (direct images)",
            ['(None)'] + columns,
            index=columns.index(image_url_col_default) + 1 if image_url_col_default in columns else 0,
            help="Column with direct image URLs (copied manually)"
        )
        if image_url_col == '(None)':
            image_url_col = None
    
    with col3:
        target_col_letter = st.text_input("Output Column (Excel)", value="B").upper()
    
    if not url_col and not image_url_col:
        st.error("⚠️ Please select at least one source column (URL or IMAGE URL)")
    else:
        st.success(f"✅ Using: {url_col or 'None'} (URL) + {image_url_col or 'None'} (IMAGE URL)")
    
    if st.button("🔍 Test First 3 URLs"):
        if not url_col and not image_url_col:
            st.error("Please select at least one source column")
        else:
            st.write("### Testing (check console for details)...")
            
            success_count = 0
            for i, row in df.head(3).iterrows():
                final_image_url = process_row(row, url_col, image_url_col)
                
                if final_image_url:
                    img = download_direct_image(final_image_url, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                    
                    if img:
                        success_count += 1
                        st.success(f"✅ Row {i+1}: Success")
                        st.image(img, width=200)
                    else:
                        st.error(f"❌ Row {i+1}: Download failed")
                else:
                    st.error(f"❌ Row {i+1}: No image URL found")
                    url_val = row.get(url_col) if url_col else None
                    if url_val and pd.notna(url_val):
                        if 'etsy.com' in str(url_val) or 'next.co.uk' in str(url_val):
                            st.info("💡 This is Etsy/Next.co.uk - please add direct image URL in IMAGE URL column")
            
            st.info(f"**Results: {success_count}/3 successful**")
    
    if st.button("Generate Excel"):
        if not url_col and not image_url_col:
            st.error("Please select at least one source column")
        else:
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
                
                status.text(f"Processing {i+1}/{total}...")
                
                final_image_url = process_row(row, url_col, image_url_col)
                
                if final_image_url:
                    img = download_direct_image(final_image_url, TARGET_WIDTH_PX, TARGET_HEIGHT_PX)
                    
                    if img:
                        count += 1
                        buf = BytesIO()
                        img.save(buf, format='PNG')
                        buf.seek(0)
                        excel_img = OpenpyxlImage(buf)
                        ws.add_image(excel_img, f"{target_col_letter}{i+2}")
                        ws.row_dimensions[i+2].height = 105
                    else:
                        failed.append(i+2)
                else:
                    failed.append(i+2)
                        
                time.sleep(0.5)
            
            if failed:
                status.warning(f"⚠️ Complete! {count}/{total} images added. Failed rows: {', '.join(map(str, failed[:10]))}{'...' if len(failed) > 10 else ''}")
            else:
                status.success(f"✅ Complete! {count}/{total} images added")
            
            out = BytesIO()
            wb.save(out)
            out.seek(0)
            
            st.download_button("📥 Download Excel", out, "output_v15.1.xlsx",
                              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            if failed:
                st.info(f"""
                **{len(failed)} failed rows:** {', '.join(map(str, failed[:20]))}{'...' if len(failed) > 20 else ''}
                
                **Common causes:**
                - Etsy/Next.co.uk URLs without manual IMAGE URL
                - Broken links
                - Unsupported image formats
                """)
    
    if st.button("Generate PowerPoint"):
        if not url_col and not image_url_col:
            st.error("Please select at least one source column")
        else:
            with st.spinner("Generating slides..."):
                prs = Presentation()
                blank = prs.slide_layouts[6]
                count = 0
                
                for i, row in df.iterrows():
                    final_image_url = process_row(row, url_col, image_url_col)
                    
                    if final_image_url:
                        img = download_direct_image(final_image_url, 800, 600)
                        
                        if img:
                            count += 1
                            slide = prs.slides.add_slide(blank)
                            buf = BytesIO()
                            img.save(buf, format='PNG')
                            buf.seek(0)
                            slide.shapes.add_picture(buf, Inches(1), Inches(1), width=Inches(8))
                            
                            tx = slide.shapes.add_textbox(Inches(1), Inches(6.5), Inches(8), Inches(0.5))
                            source = final_image_url if final_image_url else "Unknown"
                            tx.text_frame.text = f"Source: {source}"
                
                out = BytesIO()
                prs.save(out)
                out.seek(0)
                st.download_button("📥 Download PowerPoint", out, "output_v15.1.pptx",
                                  mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                st.success(f"✅ Created {count} slides")

st.markdown("---")
st.markdown("""
### 📖 **Pro Tips:**

**For auto-scraping sites (Pottery Barn, Target, West Elm):**
- Just fill in the URL column
- Leave IMAGE URL blank
- Script auto-scrapes ✅

**For protected sites (Etsy, Next.co.uk):**
1. Right-click product image → "Copy image address"
2. Paste in IMAGE URL column
3. Leave URL column as-is (for reference)

**You can mix both approaches in the same file!**
""")
