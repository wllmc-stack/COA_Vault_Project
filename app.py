import os
import io
import streamlit as st
import zipfile
from pypdf import PdfWriter
from coa_processor import COAProcessor

# Initialize processor
processor = COAProcessor()

st.set_page_config(page_title="COA Vault System", page_icon="🧪", layout="wide")

st.title("🧪 Flavor COA Vault System")

tab1, tab2, tab3 = st.tabs(["🔍 Single Lot Lookup", "📦 Bulk Lot Viewer & Downloader", "🛒 Order Lookup (Finale)"])

# --- TAB 1: Single Lot Lookup ---
with tab1:
    st.header("Search Single COA")
    lot_input = st.text_input("Enter Lot Number:", placeholder="e.g. E015261")
    
    if lot_input:
        matched_file = processor.find_coa_by_lot(lot_input.strip())
        if matched_file:
            st.success(f"Found COA: {os.path.basename(matched_file)}")
            
            # Display PDF
            with open(matched_file, "rb") as f:
                pdf_bytes = f.read()
            st.download_button("📥 Download COA PDF", data=pdf_bytes, file_name=os.path.basename(matched_file), mime="application/pdf")
            st.pdf(pdf_bytes)
        else:
            st.error("No matching COA found for that lot number.")

# --- TAB 2: Bulk Lot Viewer & Downloader ---
with tab2:
    st.header("Bulk COA Review & Packaging")
    st.write("Paste a list of lot numbers below (one per line or comma-separated):")
    
    bulk_input = st.text_area("Lot Numbers:", placeholder="E015261\nL114006\nB98201", height=150)
    
    if st.button("🔍 Search & Batch Load"):
        # Parse inputs
        raw_lots = [lot.strip() for lot in bulk_input.replace(",", "\n").split("\n") if lot.strip()]
        
        if not raw_lots:
            st.warning("Please enter at least one lot number.")
        else:
            found_files = {}
            missing_lots = []
            
            for lot in raw_lots:
                file_path = processor.find_coa_by_lot(lot)
                if file_path:
                    found_files[lot] = file_path
                else:
                    missing_lots.append(lot)
            
            # Save results in session state so viewers persist across clicks
            st.session_state['found_files'] = found_files
            st.session_state['missing_lots'] = missing_lots

    # Render results if available
    if 'found_files' in st.session_state and st.session_state['found_files']:
        found_files = st.session_state['found_files']
        missing_lots = st.session_state['missing_lots']
        
        st.divider()
        st.subheader(f"Results: Found {len(found_files)} of {len(found_files) + len(missing_lots)} COAs")
        
        if missing_lots:
            st.error(f"Missing Lot Numbers: {', '.join(missing_lots)}")
            
        # Action Bar: Download ZIP & Print/Merged PDF
        col1, col2 = st.columns(2)
        
        # 1. Download All as ZIP
        with col1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for lot, filepath in found_files.items():
                    zf.write(filepath, arcname=os.path.basename(filepath))
            
            st.download_button(
                label="📦 Download All as .ZIP",
                data=zip_buffer.getvalue(),
                file_name="COA_Batch_Package.zip",
                mime="application/zip",
                use_container_width=True
            )
            
        # 2. Combined PDF for Printing All
        with col2:
            merger = PdfWriter()
            for filepath in found_files.values():
                merger.append(filepath)
            
            merged_pdf_buffer = io.BytesIO()
            merger.write(merged_pdf_buffer)
            merger.close()
            
            st.download_button(
                label="🖨️ Download Combined PDF (Print All)",
                data=merged_pdf_buffer.getvalue(),
                file_name="Combined_COAs_Print_All.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            
        st.divider()
        st.subheader("👁️ Document Viewer Window")
        
        # View All COAs in dynamic expandable cards
        for lot, filepath in found_files.items():
            with st.expander(f"📄 Lot #{lot} — {os.path.basename(filepath)}", expanded=True):
                with open(filepath, "rb") as f:
                    pdf_data = f.read()
                st.pdf(pdf_data)

# --- TAB 3: Finale Inventory (Placeholder) ---
with tab3:
    st.header("Finale Inventory Lookup")
    st.info("Order lookup integration pending.")
