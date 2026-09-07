"""
Streamlit UI for Extracting Eye Scan Images from InstaKC PDF reports.

Run with:  streamlit run image_app.py
"""

import io
import tempfile
import os
import platform
import zipfile
from pathlib import Path

import streamlit as st
import numpy as np
import cv2
from PIL import Image
from pdf2image import convert_from_path

# Poppler path for Windows
_SCRIPT_DIR = Path(__file__).resolve().parent
if platform.system() == "Windows":
    POPPLER_PATH = str(_SCRIPT_DIR / "poppler" / "poppler-24.08.0" / "Library" / "bin")
else:
    POPPLER_PATH = None

st.set_page_config(
    page_title="Eye Scan Extractor",
    page_icon="👁️",
    layout="wide",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .hero {
        background: linear-gradient(135deg, #2b5876 0%, #4e4376 100%);
        border-radius: 16px;
        padding: 2.5rem 2rem;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 8px 32px rgba(0,0,0,.35);
        color: white;
    }
    .hero h1 { margin: 0 0 .4rem 0; font-weight: 800; }
    .hero p { margin: 0; color: #dcdcdc; }
    </style>
    <div class="hero">
        <h1>👁️ Eye Scan Extractor</h1>
        <p>Upload InstaKC PDF reports to crop and download the 6 eye scan images in an organized ZIP file.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

def extract_eye_scans(pdf_bytes: bytes, dpi: int = 600):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name

    try:
        pages = convert_from_path(tmp_path, dpi=dpi, poppler_path=POPPLER_PATH)
    finally:
        os.remove(tmp_path)

    if not pages:
        raise ValueError("Could not extract any pages from the PDF.")

    page_img = np.array(pages[0])
    gray = cv2.cvtColor(page_img, cv2.COLOR_RGB2GRAY)
    
    # Slight blur to reduce noise and connect adjacent pixels
    gray = cv2.GaussianBlur(gray, (7, 7), 0)
    
    _, thresh = cv2.threshold(gray, 230, 255, cv2.THRESH_BINARY_INV)
    
    # Use RETR_LIST to find all contours (in case images are inside another bounding box)
    contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    
    bounding_boxes = []
    
    # Scale expected area by DPI. (base is 200 DPI)
    scale_factor = (dpi / 200.0) ** 2
    min_area = 50000 * scale_factor
    max_area = 800000 * scale_factor
    
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        if min_area < area < max_area and 0.5 < w/h < 1.5:
            bounding_boxes.append((x, y, w, h))
            
    # Sort by area descending and take the top 6
    bounding_boxes.sort(key=lambda b: b[2]*b[3], reverse=True)
    
    if len(bounding_boxes) < 6:
        raise ValueError(f"Only detected {len(bounding_boxes)} scan regions, expected 6. The PDF format might be different or scans are missing.")
    
    # Take exactly 6
    bounding_boxes = bounding_boxes[:6]
    
    # Separate rows by Y coordinate
    bounding_boxes.sort(key=lambda b: b[1])
    top_row = bounding_boxes[:3]
    bottom_row = bounding_boxes[3:]
    
    # Sort each row by X coordinate
    top_row.sort(key=lambda b: b[0])
    bottom_row.sort(key=lambda b: b[0])
    
    def crop_img(bbox):
        x, y, w, h = bbox
        panel = page_img[y:y+h, x:x+w]
        
        # Find the colored heatmap to crop tightly
        hsv = cv2.cvtColor(panel, cv2.COLOR_RGB2HSV)
        s = hsv[:, :, 1]
        
        # Grey background has near-zero saturation. Colored heatmap has high saturation.
        _, color_mask = cv2.threshold(s, 20, 255, cv2.THRESH_BINARY)
        
        contours, _ = cv2.findContours(color_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        if contours:
            # Get the bounding box that encompasses all colored regions
            all_points = np.vstack(contours)
            cx, cy, cw, ch = cv2.boundingRect(all_points)
            
            # Add a tiny padding margin (2% of the panel width)
            pad = int(w * 0.02)
            cx1 = max(0, cx - pad)
            cy1 = max(0, cy - pad)
            cx2 = min(w, cx + cw + pad)
            cy2 = min(h, cy + ch + pad)
            
            final_cropped = panel[cy1:cy2, cx1:cx2]
            
            # --- Morphological Artifact Removal ---
            gray = cv2.cvtColor(final_cropped, cv2.COLOR_RGB2GRAY)
            
            # Kernel size scaled by DPI for line thickness
            # At 400 DPI, lines/text are roughly 3-6 pixels thick, so kernel of ~11 is good.
            k_size = max(5, int(11 * (dpi / 400)))
            if k_size % 2 == 0: k_size += 1
            kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k_size, k_size))
            
            # 1. Black-Hat for thin black lines
            # Closing removes small dark features, subtracting original reveals them
            closed = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, kernel)
            black_hat = cv2.subtract(closed, gray)
            _, mask_black = cv2.threshold(black_hat, 20, 255, cv2.THRESH_BINARY)
            
            # 2. Top-Hat for thin white text and crosshairs
            # Opening removes small bright features, subtracting from original reveals them
            opened = cv2.morphologyEx(gray, cv2.MORPH_OPEN, kernel)
            top_hat = cv2.subtract(gray, opened)
            _, mask_white = cv2.threshold(top_hat, 20, 255, cv2.THRESH_BINARY)
            
            defect_mask = cv2.bitwise_or(mask_black, mask_white)
            
            # Slight dilation to cover the anti-aliased borders of the artifacts
            dilate_kernel = np.ones((3,3), np.uint8)
            defect_mask = cv2.dilate(defect_mask, dilate_kernel, iterations=1)
            
            # Telea Inpainting to fill the masked artifacts with surrounding colors
            final_cleaned = cv2.inpaint(final_cropped, defect_mask, 3, cv2.INPAINT_TELEA)
            
            return Image.fromarray(final_cleaned)
        
        # Fallback if no color detected (shouldn't happen)
        return Image.fromarray(panel)
        
    left_eye_imgs = [crop_img(b) for b in top_row]
    right_eye_imgs = [crop_img(b) for b in bottom_row]
    
    return left_eye_imgs, right_eye_imgs


def process_and_zip(files, zip_name="extracted_eye_scans.zip"):
    # We will create a zip file in memory containing all images
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        progress = st.progress(0, text="Starting extraction...")
        
        for i, pdf_file in enumerate(files):
            progress.progress((i) / len(files), text=f"Processing {pdf_file.name}...")
            
            try:
                left_imgs, right_imgs = extract_eye_scans(pdf_file.getvalue())
                
                # Base folder name = PDF name without extension
                base_name = os.path.splitext(pdf_file.name)[0]
                
                # Save Left Eye images
                for idx, img in enumerate(left_imgs, 1):
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='JPEG')
                    zip_file.writestr(f"{base_name}/LEFT/left_{idx}.jpg", img_byte_arr.getvalue())
                    
                # Save Right Eye images
                for idx, img in enumerate(right_imgs, 1):
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='JPEG')
                    zip_file.writestr(f"{base_name}/RIGHT/right_{idx}.jpg", img_byte_arr.getvalue())
                    
            except Exception as e:
                st.error(f"Error processing {pdf_file.name}: {e}")
        
        progress.progress(1.0, text="Extraction complete!")
        
    zip_buffer.seek(0)
    
    st.success("✅ Extraction complete! Download your ZIP file below.")
    st.download_button(
        label=f"⬇️ Download Extracted Images ({len(files)} file{'s' if len(files)>1 else ''})",
        data=zip_buffer,
        file_name=zip_name,
        mime="application/zip",
        type="primary",
        use_container_width=True
    )


tab_single, tab_bulk = st.tabs(["📄  Single PDF Upload", "📁  Bulk PDF Upload"])

with tab_single:
    st.markdown("#### Upload a single InstaKC PDF report")
    single_file = st.file_uploader(
        "Choose a PDF file",
        type=["pdf"],
        accept_multiple_files=False,
        key="single_pdf",
        help="Select a single Remidio InstaKC corneal topography report."
    )
    
    if single_file:
        if st.button("🚀 Extract Images to ZIP", key="btn_single", type="primary"):
            process_and_zip([single_file], zip_name=f"scans_{os.path.splitext(single_file.name)[0]}.zip")

with tab_bulk:
    st.markdown("#### Upload multiple InstaKC PDF reports")
    bulk_files = st.file_uploader(
        "Choose PDF files",
        type=["pdf"],
        accept_multiple_files=True,
        key="bulk_pdfs",
        help="Select multiple Remidio InstaKC corneal topography reports."
    )
    
    if bulk_files:
        if st.button("🚀 Extract Images to ZIP", key="btn_bulk", type="primary"):
            process_and_zip(bulk_files, zip_name="bulk_eye_scans.zip")

st.markdown("---")
st.markdown('<p style="text-align:center; color:#555; font-size:.82rem;">InstaKC Extractor &middot; Built with Streamlit &amp; OpenCV</p>', unsafe_allow_html=True)
