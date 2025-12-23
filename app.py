import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Analisis Stock Dashboard", layout="wide")

st.title("📦 Dashboard Analisis Stock")
st.markdown("Upload file Excel yang berisi sheet `SO_B2B` dan `Loct_F211`.")

# --- Bagian Upload File ---
st.sidebar.header("Upload File")
uploaded_file = st.sidebar.file_uploader("Upload File Excel (.xlsx)", type=['xlsx'])

# Fungsi cleaning angka
def clean_number(x):
    if isinstance(x, str):
        x = x.replace(',', '') 
    return pd.to_numeric(x, errors='coerce')

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # Cek kelengkapan sheet
        required_sheets = ['SO_B2B', 'Loct_F211']
        missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
        
        if missing_sheets:
            st.error(f"Sheet hilang: {', '.join(missing_sheets)}")
        else:
            # Load Data
            df_so = pd.read_excel(uploaded_file, sheet_name='SO_B2B')
            df_loct = pd.read_excel(uploaded_file, sheet_name='Loct_F211')
            
            # --- PREPROCESSING UMUM ---
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
                
                # 1. Agregasi Data SO (Group by Material & Batch)
                so_agg = df_so.groupby(['Material', 'Batch Number']).agg({
                    'Ordered Quantity': 'sum',
                    'Shipment Number': lambda x: ', '.join(x.astype(str).unique()) 
                }).reset_index()
                
                so_agg.rename(columns={
                    'Batch Number': 'Batch', 
                    'Ordered Quantity': 'Total_Ordered',
                    'Shipment Number': 'List_Shipment_Numbers'
                }, inplace=True)
                
                # 2. Agregasi Data Stock (Group by Material & Batch)
                loct_agg = df_loct.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                loct_agg.rename(columns={'Unrestricted': 'Stock_Onhand'}, inplace=True)
                
                # 3. Merge & Hitung
                merged_df = pd.merge(so_agg, loct_agg, on=['Material', 'Batch'], how='left')
                merged_df['Stock_Onhand'] = merged_df['Stock_Onhand'].fillna(0)
                merged_df['Balance'] = merged_df['Stock_Onhand'] - merged_df['Total_Ordered']
                
                # 4. Filter Defisit
                deficit_df = merged_df[merged_df['Balance'] < 0].copy()
                
                # Tampilan
                if not deficit_df.empty:
                    st.error(f"Ditemukan {len(deficit_df)} Batch SKU yang defisit (Minus)!")
                    
                    # Rapikan kolom
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
                    
                    # Download
                    csv_defisit = deficit_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Download Data Defisit (CSV)", 
                        data=csv_defisit, 
                        file_name="analisis_defisit_stock.csv", 
                        mime="text/csv"
                    )
                else:
                    st.success("Aman! Tidak ada defisit stock pada batch yang diminta.")

            # =========================================
            # TAB 2: FILTER STOCK GUDANG
            # =========================================
            with tab2:
                st.subheader("Cek Ketersediaan Stock di Loct_F211")
                st.markdown("Pilih Material Code untuk melihat semua Batch dan Stock yang tersedia.")
                
                # Ambil list unik Material dari file Loct
                all_materials = sorted(df_loct['Material'].astype(str).unique())
                
                # Dropdown Pilih Material
                selected_material = st.selectbox("Pilih Material / SKU:", all_materials)
                
                if selected_material:
                    # Filter data berdasarkan material yang dipilih
                    stock_filter = df_loct[df_loct['Material'].astype(str) == selected_material].copy()
                    
                    # Grouping by Batch untuk menjumlahkan stock (jika ada duplicate row per batch)
                    stock_display = stock_filter.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                    
                    # Tampilkan Total Stock Material Tersebut
                    total_stock = stock_display['Unrestricted'].sum()
                    st.info(f"Total Stock On-Hand untuk **{selected_material}**: {total_stock:,.0f}")
                    
                    # Tabel Detail Batch
                    st.dataframe(
                        stock_display.style.format({"Unrestricted": "{:,.0f}"}), 
                        use_container_width=True
                    )
                
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        st.warning("Pastikan format file Excel dan nama sheet sudah sesuai.")

else:
    st.info("Silakan upload file Excel di sidebar.")
