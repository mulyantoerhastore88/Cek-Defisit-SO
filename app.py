import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Analisis Defisit Stock", layout="wide")

st.title("📦 Dashboard Analisis Defisit Stock")
st.markdown("""
Aplikasi ini membandingkan data **Order** dengan **Stock Gudang** dari satu file Excel.
Pastikan file Excel memiliki sheet: `SO_B2B` dan `Loct_F211`.
""")

# --- Bagian Upload File ---
st.sidebar.header("Upload File")
uploaded_file = st.sidebar.file_uploader("Upload File Excel (.xlsx)", type=['xlsx'])

# Fungsi untuk membersihkan format angka (handle string dengan koma atau angka murni)
def clean_number(x):
    if isinstance(x, str):
        x = x.replace(',', '') 
    return pd.to_numeric(x, errors='coerce')

if uploaded_file:
    try:
        # Membaca file Excel
        xls = pd.ExcelFile(uploaded_file)
        
        # Cek ketersediaan sheet
        required_sheets = ['SO_B2B', 'Loct_F211']
        missing_sheets = [s for s in required_sheets if s not in xls.sheet_names]
        
        if missing_sheets:
            st.error(f"Sheet berikut tidak ditemukan dalam file Excel: {', '.join(missing_sheets)}")
        else:
            # 1. Load Data dari Sheet
            df_so = pd.read_excel(uploaded_file, sheet_name='SO_B2B')
            df_loct = pd.read_excel(uploaded_file, sheet_name='Loct_F211')
            
            # 2. Preprocessing Data SO (Permintaan)
            # Bersihkan angka jika formatnya text, jika sudah number biarkan
            if df_so['Ordered Quantity'].dtype == 'object':
                df_so['Ordered Quantity'] = df_so['Ordered Quantity'].apply(clean_number)
            
            # --- LOGIKA UTAMA ---
            # Group by Material & Batch Number:
            # - Sum Ordered Quantity
            # - Gabungkan List Shipment Number
            so_agg = df_so.groupby(['Material', 'Batch Number']).agg({
                'Ordered Quantity': 'sum',
                'Shipment Number': lambda x: ', '.join(x.astype(str).unique()) 
            }).reset_index()
            
            so_agg.rename(columns={
                'Batch Number': 'Batch', 
                'Ordered Quantity': 'Total_Ordered',
                'Shipment Number': 'List_Shipment_Numbers'
            }, inplace=True)
            
            # 3. Preprocessing Data Stock (Persediaan)
            if df_loct['Unrestricted'].dtype == 'object':
                df_loct['Unrestricted'] = df_loct['Unrestricted'].apply(clean_number)
            
            loct_agg = df_loct.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
            loct_agg.rename(columns={'Unrestricted': 'Stock_Onhand'}, inplace=True)
            
            # 4. Analisis Defisit (Merge Data)
            # Left join SO ke Stock
            merged_df = pd.merge(so_agg, loct_agg, on=['Material', 'Batch'], how='left')
            
            # Isi NaN dengan 0 (artinya batch tidak ada di gudang sama sekali)
            merged_df['Stock_Onhand'] = merged_df['Stock_Onhand'].fillna(0)
            
            # Hitung Balance
            merged_df['Balance'] = merged_df['Stock_Onhand'] - merged_df['Total_Ordered']
            
            # Filter Defisit
            deficit_df = merged_df[merged_df['Balance'] < 0].copy()
            
            # Rapikan urutan kolom
            cols = ['Material', 'Batch', 'Total_Ordered', 'Stock_Onhand', 'Balance', 'List_Shipment_Numbers']
            deficit_df = deficit_df[cols]

            # --- TAMPILAN OUTPUT ---
            st.subheader("🚨 Hasil Analisis: Batch yang Defisit & Shipment Terdampak")
            
            if not deficit_df.empty:
                st.error(f"Ditemukan {len(deficit_df)} Batch SKU yang defisit!")
                
                # Tampilkan tabel defisit
                st.dataframe(
                    deficit_df.style.format({
                        "Total_Ordered": "{:,.0f}", 
                        "Stock_Onhand": "{:,.0f}", 
                        "Balance": "{:,.0f}"
                    }), 
                    use_container_width=True
                )
                
                # Download Button
                csv_defisit = deficit_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="Download Data Defisit (CSV)", 
                    data=csv_defisit, 
                    file_name="defisit_stock_analysis.csv", 
                    mime="text/csv"
                )
                
                # --- Tampilkan Stock Tersedia Lainnya ---
                st.divider()
                st.subheader("🔍 Alternatif Stock Lain yang Tersedia")
                st.info("Tabel ini menampilkan SEMUA Batch yang tersedia di Gudang untuk Material yang sedang defisit.")
                
                # Ambil list material yang bermasalah
                deficit_materials = deficit_df['Material'].unique()
                
                # Filter data master gudang hanya untuk material tsb
                available_batches = df_loct[df_loct['Material'].isin(deficit_materials)].copy()
                
                # Agregasi
                available_display = available_batches.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                
                st.dataframe(available_display.style.format({"Unrestricted": "{:,.0f}"}), use_container_width=True)
                
            else:
                st.success("Mantap Bro! Tidak ada defisit stock untuk orderan ini.")
                
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
        st.warning("Pastikan file Excel tidak corrupt dan memiliki nama kolom yang sesuai.")

else:
    st.info("Silakan upload file Excel (.xlsx) di sidebar.")
