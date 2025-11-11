import streamlit as st
import pandas as pd
import io
from datetime import datetime
from typing import Optional, Tuple

# Page configuration
st.set_page_config(
    page_title="Verifikasi Tagihan Fulfillment",
    page_icon="�",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for clean purple theme (light & dark mode compatible)
st.markdown("""
<style>
    .main-header {
        font-size: 2.8rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 2rem;
        color: #ffffff !important;
        padding: 2rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
    }
    
    .section-header {
        font-size: 1.4rem;
        font-weight: 600;
        color: var(--text-color) !important;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding: 1rem;
        border-left: 4px solid #7c3aed;
        border-radius: 8px;
        background-color: var(--background-color);
    }
    
    /* Default variables (light theme) */
    :root {
        --text-color: #2d3748;
        --background-color: rgba(248, 250, 252, 0.9);
        --tab-bg: #f8fafc;
        --tab-panel-bg: #ffffff;
        --instructions-bg: #ffffff;
        --instructions-border: #cbd5e0;
        --instructions-header: #2b6cb0;
    }
    
    /* Light mode explicit */
    [data-theme="light"] {
        --text-color: #2d3748;
        --background-color: rgba(248, 250, 252, 0.9);
        --tab-bg: #f8fafc;
        --tab-panel-bg: #ffffff;
        --instructions-bg: #ffffff;
        --instructions-border: #cbd5e0;
        --instructions-header: #2b6cb0;
    }
    
    /* Dark mode */
    [data-theme="dark"] {
        --text-color: #f7fafc;
        --background-color: rgba(45, 55, 72, 0.8);
        --tab-bg: #2d3748;
        --tab-panel-bg: #1a202c;
        --instructions-bg: #2d3748;
        --instructions-border: #4a5568;
        --instructions-header: #e2e8f0;
    }
    
    /* Auto detect theme - Dark */
    @media (prefers-color-scheme: dark) {
        :root {
            --text-color: #f7fafc;
            --background-color: rgba(45, 55, 72, 0.8);
            --tab-bg: #2d3748;
            --tab-panel-bg: #1a202c;
            --instructions-bg: #2d3748;
            --instructions-border: #4a5568;
            --instructions-header: #e2e8f0;
        }
    }
    
    /* Auto detect theme - Light */
    @media (prefers-color-scheme: light) {
        :root {
            --text-color: #2d3748;
            --background-color: rgba(248, 250, 252, 0.9);
            --tab-bg: #f8fafc;
            --tab-panel-bg: #ffffff;
            --instructions-bg: #ffffff;
            --instructions-border: #cbd5e0;
            --instructions-header: #2b6cb0;
        }
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #7c3aed 0%, #5b21b6 100%) !important;
        color: white !important;
        border: none !important;
        padding: 0.75rem 1.5rem;
        border-radius: 10px;
        font-weight: 600;
        width: 100%;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #5b21b6 0%, #4c1d95 100%) !important;
        transform: translateY(-1px);
        color: white !important;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        padding-left: 20px;
        padding-right: 20px;
        background-color: var(--tab-bg) !important;
        border-radius: 10px 10px 0 0;
        color: var(--text-color) !important;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #7c3aed 0%, #5b21b6 100%) !important;
        color: white !important;
    }
    
    .stTabs [data-baseweb="tab-panel"] {
        padding: 2rem 1rem;
        background-color: var(--tab-panel-bg) !important;
        border-radius: 0 0 10px 10px;
        margin-bottom: 2rem;
    }
    
    /* Fix text visibility in dark mode */
    .stMarkdown h3 {
        color: var(--text-color) !important;
    }
    
    .stCaption {
        color: var(--text-color) !important;
        opacity: 0.8;
    }
    
    /* Ensure all text is visible */
    .stMarkdown p, .stMarkdown li, .stMarkdown span, .stMarkdown div {
        color: var(--text-color) !important;
    }
    
    /* Force visibility for all markdown content */
    .stMarkdown, .stMarkdown * {
        color: var(--text-color) !important;
    }
    
    /* Instructions section styling */
    .instructions-section {
        color: var(--text-color) !important;
        background-color: var(--instructions-bg) !important;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid var(--instructions-border);
        margin: 1rem 0;
    }
    
    .instructions-section h3 {
        color: var(--instructions-header) !important;
        font-weight: 600;
        margin-bottom: 1rem;
    }
    
    .instructions-section p, 
    .instructions-section li, 
    .instructions-section strong {
        color: var(--text-color) !important;
    }
    
    .instructions-section ol {
        padding-left: 1.2rem;
    }
    
    .instructions-section li {
        margin-bottom: 0.5rem;
        line-height: 1.6;
    }
</style>

<script>
// Dynamic theme detection
function updateTheme() {
    const isDark = window.matchMedia('(prefers-color-scheme: dark)').matches || 
                   document.querySelector('[data-theme="dark"]') ||
                   document.body.classList.contains('dark') ||
                   getComputedStyle(document.body).backgroundColor === 'rgb(14, 17, 23)';
    
    const root = document.documentElement;
    if (isDark) {
        root.style.setProperty('--text-color', '#f7fafc');
        root.style.setProperty('--background-color', 'rgba(45, 55, 72, 0.8)');
        root.style.setProperty('--tab-bg', '#2d3748');
        root.style.setProperty('--tab-panel-bg', '#1a202c');
        root.style.setProperty('--instructions-bg', '#2d3748');
        root.style.setProperty('--instructions-border', '#4a5568');
        root.style.setProperty('--instructions-header', '#e2e8f0');
    } else {
        root.style.setProperty('--text-color', '#2d3748');
        root.style.setProperty('--background-color', 'rgba(248, 250, 252, 0.9)');
        root.style.setProperty('--tab-bg', '#f8fafc');
        root.style.setProperty('--tab-panel-bg', '#ffffff');
        root.style.setProperty('--instructions-bg', '#ffffff');
        root.style.setProperty('--instructions-border', '#cbd5e0');
        root.style.setProperty('--instructions-header', '#2b6cb0');
    }
}

// Update theme on load and when it changes
updateTheme();
window.matchMedia('(prefers-color-scheme: dark)').addListener(updateTheme);

// Observe for Streamlit theme changes
const observer = new MutationObserver(updateTheme);
observer.observe(document.body, { attributes: true, attributeFilter: ['class', 'data-theme'] });
</script>
""", unsafe_allow_html=True)

def load_excel_file(uploaded_file) -> Optional[pd.DataFrame]:
    """Load Excel file and return DataFrame"""
    try:
        if uploaded_file is not None:
            return pd.read_excel(uploaded_file)
        return None
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

def verify_data(outgoing_df: pd.DataFrame, everpro_df: pd.DataFrame, shopee_df: pd.DataFrame) -> pd.DataFrame:
    """Verify data by checking column D against reference databases"""
    verified_df = outgoing_df.copy()
    
    # Create verification column M
    verification_results = []
    
    for index, row in outgoing_df.iterrows():
        value_to_check = row.iloc[3] if len(row) > 3 else None  # Column D (index 3)
        
        if pd.isna(value_to_check):
            verification_results.append("Tidak Terverifikasi")
            continue
        
        # Check in Everpro (Column C - index 2)
        found_in_everpro = False
        if everpro_df is not None and len(everpro_df.columns) > 2:
            found_in_everpro = value_to_check in everpro_df.iloc[:, 2].values
        
        # Check in Shopee JNE Surabaya (Column E - index 4)
        found_in_shopee = False
        if shopee_df is not None and len(shopee_df.columns) > 4:
            found_in_shopee = value_to_check in shopee_df.iloc[:, 4].values
        
        # Verification: found in either Everpro OR Shopee (not both required)
        if found_in_everpro or found_in_shopee:
            verification_results.append("Terverifikasi")
        else:
            verification_results.append("Tidak Terverifikasi")
    
    # Add verification column safely
    # Always add as a new column to avoid iloc issues
    verified_df['Status_Verifikasi'] = verification_results
    
    return verified_df

def create_combined_excel(jne_df: Optional[pd.DataFrame] = None, non_jne_df: Optional[pd.DataFrame] = None) -> bytes:
    """Create combined Excel file with separate sheets for JNE and Non-JNE"""
    output = io.BytesIO()
    
    # Generate filename with current date and time
    now = datetime.now()
    filename = f"{now.strftime('%Y-%m-%d-%H')}-Tagihan Fulfillment.xlsx"
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if jne_df is not None:
            jne_df.to_excel(writer, sheet_name='Outgoing JNE', index=False)
        if non_jne_df is not None:
            non_jne_df.to_excel(writer, sheet_name='Outgoing Non JNE', index=False)
    
    return output.getvalue(), filename

def main():
    # Header
    st.markdown('<div class="main-header">Verifikasi Tagihan Fulfillment</div>', unsafe_allow_html=True)
    
    # Initialize session state
    if 'files_uploaded' not in st.session_state:
        st.session_state.files_uploaded = {
            'everpro': False,
            'shopee': False,
            'outgoing_jne': False,
            'outgoing_non_jne': False
        }
    
    # File upload section with tabs
    st.markdown('<div class="section-header">Upload File Excel</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4 = st.tabs(["Database Everpro", "Shopee JNE Surabaya", "Outgoing JNE", "Outgoing Non JNE"])
    
    with tab1:
        st.markdown("### Database Everpro")
        st.caption("File Excel dengan kolom sampai BO - digunakan sebagai referensi kolom C")
        everpro_file = st.file_uploader("Pilih file Excel Everpro", type=['xlsx', 'xls'], key="everpro")
        if everpro_file:
            st.session_state.everpro_df = load_excel_file(everpro_file)
            st.session_state.files_uploaded['everpro'] = True
            st.success("✅ File Everpro berhasil diupload")
            st.info(f"📊 Total baris data: {len(st.session_state.everpro_df)}")
    
    with tab2:
        st.markdown("### Shopee JNE Surabaya")
        st.caption("File Excel dengan kolom sampai AW - digunakan sebagai referensi kolom E")
        shopee_file = st.file_uploader("Pilih file Excel Shopee JNE Surabaya", type=['xlsx', 'xls'], key="shopee")
        if shopee_file:
            st.session_state.shopee_df = load_excel_file(shopee_file)
            st.session_state.files_uploaded['shopee'] = True
            st.success("✅ File Shopee JNE Surabaya berhasil diupload")
            st.info(f"📊 Total baris data: {len(st.session_state.shopee_df)}")
    
    with tab3:
        st.markdown("### Outgoing JNE")
        st.caption("File Excel dengan kolom sampai L - akan diverifikasi berdasarkan kolom D")
        outgoing_jne_file = st.file_uploader("Pilih file Excel Outgoing JNE", type=['xlsx', 'xls'], key="outgoing_jne")
        if outgoing_jne_file:
            st.session_state.outgoing_jne_df = load_excel_file(outgoing_jne_file)
            st.session_state.files_uploaded['outgoing_jne'] = True
            st.success("✅ File Outgoing JNE berhasil diupload")
            st.info(f"📊 Total baris data: {len(st.session_state.outgoing_jne_df)}")
    
    with tab4:
        st.markdown("### Outgoing Non JNE")
        st.caption("File Excel dengan kolom sampai L - akan diverifikasi berdasarkan kolom D")
        outgoing_non_jne_file = st.file_uploader("Pilih file Excel Outgoing Non JNE", type=['xlsx', 'xls'], key="outgoing_non_jne")
        if outgoing_non_jne_file:
            st.session_state.outgoing_non_jne_df = load_excel_file(outgoing_non_jne_file)
            st.session_state.files_uploaded['outgoing_non_jne'] = True
            st.success("✅ File Outgoing Non JNE berhasil diupload")
            st.info(f"📊 Total baris data: {len(st.session_state.outgoing_non_jne_df)}")
    
    # Add some spacing
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Verification section
    st.markdown('<div class="section-header">Proses Verifikasi Data</div>', unsafe_allow_html=True)
    
    # Check if minimum required files are uploaded
    minimum_files_uploaded = (
        st.session_state.files_uploaded['everpro'] or 
        st.session_state.files_uploaded['shopee']
    ) and (
        st.session_state.files_uploaded['outgoing_jne'] or 
        st.session_state.files_uploaded['outgoing_non_jne']
    )
    
    if minimum_files_uploaded:
        st.markdown("<br>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2, gap="large")
        
        # Process Outgoing JNE
        if st.session_state.files_uploaded['outgoing_jne']:
            with col1:
                st.subheader("Verifikasi Outgoing JNE")
                if st.button("Proses Verifikasi JNE", key="verify_jne"):
                    with st.spinner("Memproses verifikasi..."):
                        everpro_df = st.session_state.get('everpro_df')
                        shopee_df = st.session_state.get('shopee_df')
                        outgoing_jne_df = st.session_state.get('outgoing_jne_df')
                        
                        verified_jne_df = verify_data(outgoing_jne_df, everpro_df, shopee_df)
                        st.session_state.verified_jne_df = verified_jne_df
                        
                        # Show verification summary
                        total_records = len(verified_jne_df)
                        verified_count = (verified_jne_df['Status_Verifikasi'] == "Terverifikasi").sum()
                        
                        st.success(f"**Verifikasi JNE Selesai!**  \n{verified_count} dari {total_records} data terverifikasi")
                        
                        # Show preview
                        st.subheader("Preview Hasil Verifikasi")
                        st.dataframe(verified_jne_df.head(), use_container_width=True)
        
        # Process Outgoing Non JNE
        if st.session_state.files_uploaded['outgoing_non_jne']:
            with col2:
                st.subheader("Verifikasi Outgoing Non JNE")
                if st.button("Proses Verifikasi Non JNE", key="verify_non_jne"):
                    with st.spinner("Memproses verifikasi..."):
                        everpro_df = st.session_state.get('everpro_df')
                        shopee_df = st.session_state.get('shopee_df')
                        outgoing_non_jne_df = st.session_state.get('outgoing_non_jne_df')
                        
                        verified_non_jne_df = verify_data(outgoing_non_jne_df, everpro_df, shopee_df)
                        st.session_state.verified_non_jne_df = verified_non_jne_df
                        
                        # Show verification summary
                        total_records = len(verified_non_jne_df)
                        verified_count = (verified_non_jne_df['Status_Verifikasi'] == "Terverifikasi").sum()
                        
                        st.success(f"**Verifikasi Non JNE Selesai!**  \n{verified_count} dari {total_records} data terverifikasi")
                        
                        # Show preview
                        st.subheader("Preview Hasil Verifikasi")
                        st.dataframe(verified_non_jne_df.head(), use_container_width=True)
        
        # Combined Download Section
        if 'verified_jne_df' in st.session_state or 'verified_non_jne_df' in st.session_state:
            st.markdown("<br><br>", unsafe_allow_html=True)
            st.markdown('<div class="section-header">Download Hasil Verifikasi</div>', unsafe_allow_html=True)
            
            jne_df = st.session_state.get('verified_jne_df')
            non_jne_df = st.session_state.get('verified_non_jne_df')
            
            excel_data, filename = create_combined_excel(jne_df, non_jne_df)
            
            # Show summary
            total_jne = len(jne_df) if jne_df is not None else 0
            total_non_jne = len(non_jne_df) if non_jne_df is not None else 0
            verified_jne = (jne_df['Status_Verifikasi'] == "Terverifikasi").sum() if jne_df is not None else 0
            verified_non_jne = (non_jne_df['Status_Verifikasi'] == "Terverifikasi").sum() if non_jne_df is not None else 0
            
            # Center the summary and download
            col_center1, col_center2, col_center3 = st.columns([1, 2, 1])
            with col_center2:
                st.markdown(f"""
                **Ringkasan Verifikasi:**
                - JNE: {verified_jne}/{total_jne} data terverifikasi
                - Non JNE: {verified_non_jne}/{total_non_jne} data terverifikasi
                - **Total: {verified_jne + verified_non_jne}/{total_jne + total_non_jne} data terverifikasi**
                """)
                
                st.download_button(
                    label="Download Tagihan Fulfillment (Excel)",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
    
    else:
        st.warning("""
        **Upload File Diperlukan**
        
        Silakan upload minimal:
        • 1 file database (Everpro atau Shopee JNE Surabaya)
        • 1 file yang akan diverifikasi (Outgoing JNE atau Outgoing Non JNE)
        """)
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div class="instructions-section">
        <h3>📋 Petunjuk Penggunaan</h3>
        <ol>
            <li><strong>Upload File Database</strong>: Upload file Everpro dan/atau Shopee JNE Surabaya sebagai referensi</li>
            <li><strong>Upload File Verifikasi</strong>: Upload file Outgoing JNE dan/atau Outgoing Non JNE yang akan diverifikasi</li>
            <li><strong>Proses Verifikasi</strong>: Klik tombol verifikasi untuk memulai pengecekan data</li>
            <li><strong>Download Hasil</strong>: Download file Excel gabungan dengan 2 sheet (JNE dan Non JNE)</li>
        </ol>
        
        <p style="margin-top: 1rem;"><strong>📁 Format File Download</strong>: YYYY-MM-DD-HH-Tagihan Fulfillment.xlsx</p>
        
        <p style="margin-top: 1rem;"><strong>📝 Catatan</strong>: Verifikasi dilakukan berdasarkan pencocokan nilai di Kolom D file Outgoing dengan Kolom C (Everpro) dan Kolom E (Shopee JNE Surabaya).</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
