import os
import re
import io
import zipfile
import base64
import streamlit as st
import pandas as pd

# --- CONFIGURATION ---
VAULT_DIR = r"C:\Users\willi\OneDrive\Desktop\Cleaned_Vault_Folder"
LOG_EXCEL_PATH = os.path.join(VAULT_DIR, "COA_Processing_Log.xlsx")

st.set_page_config(page_title="COA Vault Finder", layout="wide", page_icon="🧪")

@st.cache_data(ttl=60)
def load_vault_index():
    """Reads the Excel log or scans the vault directory to index available COAs."""
    if os.path.exists(LOG_EXCEL_PATH):
        df = pd.read_excel(LOG_EXCEL_PATH)
        df.fillna("N/A", inplace=True)
        return df
    else:
        records = []
        for root, _, files in os.walk(VAULT_DIR):
            for f in files:
                if f.endswith(".pdf"):
                    records.append({
                        "Original File": f,
                        "New File Name": f,
                        "Flavor": f.split("_")[0] if "_" in f else "Unknown",
                        "Lot / Lab": f.split("_")[1] if len(f.split("_")) > 1 else "Unknown",
                        "Status": "Indexed"
                    })
        return pd.DataFrame(records)

def find_pdf_file(target_lot):
    """Recursively searches the Cleaned Vault for a PDF matching the given lot string."""
    clean_target = re.sub(r'[^A-Za-z0-9]', '', str(target_lot)).lower()
    if not clean_target:
        return None

    for root, _, files in os.walk(VAULT_DIR):
        for file in files:
            if file.lower().endswith(".pdf"):
                clean_filename = re.sub(r'[^A-Za-z0-9]', '', file).lower()
                if clean_target in clean_filename:
                    return os.path.join(root, file)
    return None

# --- UI LAYOUT ---
st.title("🧪 Flavor COA Vault System")
st.markdown("Search, view, and print Certificates of Analysis across all flavor brands.")

tab1, tab2, tab3 = st.tabs(["🔍 Single Lot Lookup", "📦 Bulk Lot Search", "🛒 Order Lookup (Finale)"])

# ==========================================
# TAB 1: SINGLE LOT LOOKUP
# ==========================================
with tab1:
    st.subheader("Look Up Single COA")
    search_lot = st.text_input("Enter Lot / Batch Number:", placeholder="e.g., E015261, L114006, 0001053231").strip()

    if search_lot:
        match_path = find_pdf_file(search_lot)
        
        if match_path:
            st.success(f"Found COA: **{os.path.basename(match_path)}**")
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.markdown("### Document Options")
                with open(match_path, "rb") as pdf_file:
                    pdf_bytes = pdf_file.read()
                    
                st.download_button(
                    label="📥 Download PDF for Printing",
                    data=pdf_bytes,
                    file_name=os.path.basename(match_path),
                    mime="application/pdf"
                )
                
            with col2:
                st.markdown("### Document Preview")
                base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="600" type="application/pdf"></iframe>'
                st.markdown(pdf_display, unsafe_allow_html=True)
        else:
            st.error(f"No COA found in the vault matching lot: **{search_lot}**")

# ==========================================
# TAB 2: BULK LOT LOOKUP
# ==========================================
with tab2:
    st.subheader("Bulk Search & Package")
    bulk_input = st.text_area("Paste Lot Numbers (One per line or comma-separated):", height=150)
    
    if st.button("Search All Lots"):
        raw_lots = re.split(r'[\n,]+', bulk_input)
        lots_to_search = [l.strip() for l in raw_lots if l.strip()]
        
        if lots_to_search:
            found_files = []
            missing_lots = []
            
            for lot in lots_to_search:
                filepath = find_pdf_file(lot)
                if filepath:
                    found_files.append((lot, filepath))
                else:
                    missing_lots.append(lot)
            
            st.write(f"Matched **{len(found_files)}** of **{len(lots_to_search)}** lot numbers.")
            
            if missing_lots:
                st.warning(f"Missing COAs for lots: {', '.join(missing_lots)}")
                
            if found_files:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for lot, path in found_files:
                        zf.write(path, os.path.basename(path))
                
                st.download_button(
                    label="📦 Download Matched COAs as ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="Bulk_COAs.zip",
                    mime="application/zip"
                )
        else:
            st.warning("Please enter at least one lot number.")

# ==========================================
# TAB 3: FINALE INVENTORY ORDER LOOKUP
# ==========================================
with tab3:
    st.subheader("Finale Inventory Integration")
    st.info("This tab will connect to Finale Inventory to pull lot numbers associated with an Order ID.")
    order_id = st.text_input("Enter Finale Order Number:", placeholder="e.g., 10452")
    if order_id:
        st.write(f"Connecting to Finale to locate lots for Order #{order_id}...")