import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Analisis Defisit Stock", layout="wide")

st.title("📦 Dashboard Analisis Defisit Stock")
st.markdown("""
Aplikasi ini membandingkan data **Order (SO_B2B)** dengan **Stock Gudang (Loct_F211)** untuk menemukan Batch yang defisit.
**Update:** Menampilkan List Shipment Number yang terdampak.
""")

# --- Bagian Upload File ---
st.sidebar.header("Upload File")
file_so = st.sidebar.file_uploader("Upload SO_B2B (CSV)", type=['csv'])
file_loct = st.sidebar.file_uploader("Upload Loct_F211 (CSV)", type=['csv'])
file_sap = st.sidebar.file_uploader("Upload Sap Code (CSV) - Opsional", type=['csv'])

# Fungsi untuk membersihkan format angka
def clean_number(x):
    if isinstance(x, str):
        x = x.replace(',', '') 
    return pd.to_numeric(x, errors='coerce')

if file_so and file_loct:
    try:
        # 1. Load Data
        df_so = pd.read_csv(file_so)
        df_loct = pd.read_csv(file_loct)
        
        # Mapping Nama Produk (Optional)
        product_map = {}
        if file_sap:
            df_sap = pd.read_csv(file_sap)
            sap_cols = df_sap.columns.tolist()
            # Logika mapping sku sap ke deskripsi
            if 'SKU SAP' in sap_cols and 'Product Description' in sap_cols:
                 product_map = dict(zip(df_sap['SKU SAP'], df_sap['Product Description']))
        
        # 2. Preprocessing Data SO (Permintaan)
        df_so['Ordered Quantity'] = df_so['Ordered Quantity'].apply(clean_number)
        
        # --- LOGIKA BARU DI SINI ---
        # Group by Material & Batch:
        # - Sum Ordered Quantity
        # - List unik Shipment Number
        so_agg = df_so.groupby(['Material', 'Batch Number']).agg({
            'Ordered Quantity': 'sum',
            'Shipment Number': lambda x: ', '.join(x.astype(str).unique()) # Gabungkan nomor shipment
        }).reset_index()
        
        so_agg.rename(columns={
            'Batch Number': 'Batch', 
            'Ordered Quantity': 'Total_Ordered',
            'Shipment Number': 'List_Shipment_Numbers' # Nama kolom baru
        }, inplace=True)
        
        # 3. Preprocessing Data Stock (Persediaan)
        df_loct['Unrestricted'] = df_loct['Unrestricted'].apply(clean_number)
        
        loct_agg = df_loct.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
        loct_agg.rename(columns={'Unrestricted': 'Stock_Onhand'}, inplace=True)
        
        # 4. Analisis Defisit (Merge Data)
        merged_df = pd.merge(so_agg, loct_agg, on=['Material', 'Batch'], how='left')
        
        # Isi NaN dengan 0 untuk stock yang tidak ditemukan
        merged_df['Stock_Onhand'] = merged_df['Stock_Onhand'].fillna(0)
        
        # Hitung Selisih
        merged_df['Balance'] = merged_df['Stock_Onhand'] - merged_df['Total_Ordered']
        
        # Filter Defisit
        deficit_df = merged_df[merged_df['Balance'] < 0].copy()
        
        # Rapikan Kolom
        if product_map:
            deficit_df['Product Name'] = deficit_df['Material'].map(product_map)
            # Urutan kolom dengan Product Name
            cols = ['Material', 'Product Name', 'Batch', 'Total_Ordered', 'Stock_Onhand', 'Balance', 'List_Shipment_Numbers']
            # Pastikan hanya mengambil kolom yang ada (jaga-jaga)
            final_cols = [c for c in cols if c in deficit_df.columns]
            deficit_df = deficit_df[final_cols]
        else:
             # Urutan kolom tanpa Product Name
            cols = ['Material', 'Batch', 'Total_Ordered', 'Stock_Onhand', 'Balance', 'List_Shipment_Numbers']
            deficit_df = deficit_df[cols]

        # --- TAMPILAN OUTPUT ---
        st.subheader("🚨 Hasil Analisis: Batch yang Defisit & Shipment Terdampak")
        
        if not deficit_df.empty:
            st.error(f"Ditemukan {len(deficit_df)} Batch SKU yang defisit!")
            
            # Tampilkan dataframe
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
            st.download_button("Download Data Defisit (CSV)", csv_defisit, "defisit_stock_with_shipment.csv", "text/csv")
            
            # --- Tampilkan Stock Tersedia Lainnya ---
            st.divider()
            st.subheader("🔍 Stock Lain yang Tersedia (Untuk Substitusi)")
            st.info("Tabel ini menampilkan SEMUA Batch yang tersedia di Gudang untuk Material yang defisit di atas.")
            
            deficit_materials = deficit_df['Material'].unique()
            available_batches = df_loct[df_loct['Material'].isin(deficit_materials)].copy()
            
            available_display = available_batches.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
            
            if product_map:
                available_display['Product Name'] = available_display['Material'].map(product_map)
                cols = ['Material', 'Product Name', 'Batch', 'Unrestricted']
                available_display = available_display[[c for c in cols if c in available_display.columns]]
            
            st.dataframe(available_display.style.format({"Unrestricted": "{:,.0f}"}), use_container_width=True)
            
        else:
            st.success("Aman Bro! Tidak ada defisit stock.")
            
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        st.warning("Cek kembali nama kolom di file CSV. Pastikan ada 'Shipment Number' di file SO.")

else:
    st.info("Silakan upload file SO_B2B dan Loct_F211.")
