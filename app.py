import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Analisis Stock Dashboard", layout="wide")

st.title("📦 Dashboard Analisis Stock")
st.markdown("Upload file Excel yang berisi sheet `SO_B2B` dan `Loct_F211`.")

# --- FUNGSI CACHING (Supaya Tidak Lemot) ---
@st.cache_data
def load_data(file):
    """Fungsi ini hanya akan jalan sekali saat file diupload."""
    xls = pd.ExcelFile(file)
    # Cek sheet
    required_sheets = ['SO_B2B', 'Loct_F211']
    missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
    
    if missing_sheets:
        return None, None, f"Sheet hilang: {', '.join(missing_sheets)}"
    
    df_so = pd.read_excel(file, sheet_name='SO_B2B')
    df_loct = pd.read_excel(file, sheet_name='Loct_F211')
    
    return df_so, df_loct, None

# Fungsi cleaning angka
def clean_number(x):
    if isinstance(x, str):
        x = x.replace(',', '') 
    return pd.to_numeric(x, errors='coerce')

# --- Bagian Upload File ---
st.sidebar.header("Upload File")
uploaded_file = st.sidebar.file_uploader("Upload File Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    # Panggil fungsi load_data dengan caching
    df_so, df_loct, error_msg = load_data(uploaded_file)
    
    if error_msg:
        st.error(error_msg)
    elif df_so is not None and df_loct is not None:
        try:
            # --- PREPROCESSING ---
            # Pastikan tipe data konsisten (String untuk Material biar bisa dicocokkan)
            df_so['Material'] = df_so['Material'].astype(str)
            df_loct['Material'] = df_loct['Material'].astype(str)

            # Bersihkan angka
            if df_so['Ordered Quantity'].dtype == 'object':
                df_so['Ordered Quantity'] = df_so['Ordered Quantity'].apply(clean_number)
            
            if df_loct['Unrestricted'].dtype == 'object':
                df_loct['Unrestricted'] = df_loct['Unrestricted'].apply(clean_number)

            # --- MEMBUAT TABS ---
            tab1, tab2 = st.tabs(["🚨 Analisis Defisit", "🔍 Cek Stock Gudang"])

            # =========================================
            # TAB 1: HASIL ANALISIS DEFISIT
            # =========================================
            with tab1:
                st.subheader("Analisis Batch Defisit & Shipment Terdampak")
                
                # Agregasi Data SO
                so_agg = df_so.groupby(['Material', 'Batch Number']).agg({
                    'Ordered Quantity': 'sum',
                    'Shipment Number': lambda x: ', '.join(x.astype(str).unique()) 
                }).reset_index()
                
                so_agg.rename(columns={
                    'Batch Number': 'Batch', 
                    'Ordered Quantity': 'Total_Ordered',
                    'Shipment Number': 'List_Shipment_Numbers'
                }, inplace=True)
                
                # Agregasi Data Stock
                loct_agg = df_loct.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                loct_agg.rename(columns={'Unrestricted': 'Stock_Onhand'}, inplace=True)
                
                # Merge
                merged_df = pd.merge(so_agg, loct_agg, on=['Material', 'Batch'], how='left')
                merged_df['Stock_Onhand'] = merged_df['Stock_Onhand'].fillna(0)
                merged_df['Balance'] = merged_df['Stock_Onhand'] - merged_df['Total_Ordered']
                
                # Filter Defisit
                deficit_df = merged_df[merged_df['Balance'] < 0].copy()
                
                if not deficit_df.empty:
                    st.error(f"Ditemukan {len(deficit_df)} Batch SKU yang defisit!")
                    
                    cols = ['Material', 'Batch', 'Total_Ordered', 'Stock_Onhand', 'Balance', 'List_Shipment_Numbers']
                    deficit_df = deficit_df[cols]
                    
                    st.dataframe(
                        deficit_df.style.format({
                            "Total_Ordered": "{:,.0f}", 
                            "Stock_Onhand": "{:,.0f}", 
                            "Balance": "{:,.0f}"
                        }), 
                        use_container_width=True
                    )
                    
                    csv_defisit = deficit_df.to_csv(index=False).encode('utf-8')
                    st.download_button("Download Data Defisit (CSV)", csv_defisit, "analisis_defisit.csv", "text/csv")
                else:
                    st.success("Aman! Tidak ada defisit stock pada batch yang diminta.")

            # =========================================
            # TAB 2: FILTER STOCK GUDANG (OPTIMIZED)
            # =========================================
            with tab2:
                st.subheader("Cek Detail Stock per SKU")
                
                # Ambil list unik Material (di-cache otomatis oleh Streamlit karena df_loct dari cache)
                all_materials = sorted(df_loct['Material'].unique())
                
                # Dropdown
                selected_material = st.selectbox("Pilih Material / SKU:", all_materials)
                
                if selected_material:
                    # 1. Hitung Total Stock di Gudang (Loct)
                    stock_filter = df_loct[df_loct['Material'] == selected_material]
                    total_stock_qty = stock_filter['Unrestricted'].sum()
                    
                    # 2. Hitung Total Order di SO (Request User)
                    so_filter = df_so[df_so['Material'] == selected_material]
                    total_so_qty = so_filter['Ordered Quantity'].sum()
                    
                    # Tampilkan Summary Cards (Metrik)
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total Order (SO)", f"{total_so_qty:,.0f}", help="Total yang diminta di sheet SO_B2B")
                    col2.metric("Total Stock (Gudang)", f"{total_stock_qty:,.0f}", help="Total yang ada di sheet Loct_F211")
                    
                    # Indikator warna Balance
                    balance_global = total_stock_qty - total_so_qty
                    col3.metric("Global Balance", f"{balance_global:,.0f}", delta_color="normal")

                    st.divider()
                    
                    # Tampilkan Tabel Batch Stock
                    st.write(f"**Detail Batch Stock Gudang untuk {selected_material}:**")
                    
                    stock_display = stock_filter.groupby(['Batch'])['Unrestricted'].sum().reset_index()
                    
                    if not stock_display.empty:
                        st.dataframe(
                            stock_display.style.format({"Unrestricted": "{:,.0f}"}), 
                            use_container_width=True
                        )
                    else:
                        st.warning("Material ini tidak ditemukan di list Stock Gudang (Loct_F211).")

        except Exception as e:
            st.error(f"Terjadi kesalahan logika: {e}")

else:
    st.info("Silakan upload file Excel di sidebar.")
