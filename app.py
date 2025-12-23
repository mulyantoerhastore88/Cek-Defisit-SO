import streamlit as st
import pandas as pd

# Konfigurasi Halaman
st.set_page_config(page_title="Analisis Stock Dashboard", layout="wide")

st.title("📦 Dashboard Analisis Stock")
st.markdown("Upload file Excel yang berisi sheet `SO_B2B` dan `Loct_F211`.")

# --- FUNGSI CACHING ---
@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
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

# --- MAIN APP ---
st.sidebar.header("Upload File")
uploaded_file = st.sidebar.file_uploader("Upload File Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    df_so, df_loct, error_msg = load_data(uploaded_file)
    
    if error_msg:
        st.error(error_msg)
    elif df_so is not None and df_loct is not None:
        try:
            # --- PREPROCESSING ---
            df_so['Material'] = df_so['Material'].astype(str)
            df_loct['Material'] = df_loct['Material'].astype(str)
            
            # Cleaning Data
            if df_so['Ordered Quantity'].dtype == 'object':
                df_so['Ordered Quantity'] = df_so['Ordered Quantity'].apply(clean_number)
            
            if df_loct['Unrestricted'].dtype == 'object':
                df_loct['Unrestricted'] = df_loct['Unrestricted'].apply(clean_number)

            # --- TABS ---
            tab1, tab2 = st.tabs(["🚨 Analisis Defisit", "🔍 Cek Stock & Alokasi"])

            # =========================================
            # TAB 1: HASIL ANALISIS DEFISIT
            # =========================================
            with tab1:
                st.subheader("Analisis Batch Defisit")
                
                # Grouping SO
                so_agg = df_so.groupby(['Material', 'Batch Number']).agg({
                    'Ordered Quantity': 'sum',
                    'Shipment Number': lambda x: ', '.join(x.astype(str).unique()) 
                }).reset_index()
                
                so_agg.rename(columns={
                    'Batch Number': 'Batch', 
                    'Ordered Quantity': 'Total_Ordered',
                    'Shipment Number': 'List_Shipment_Numbers'
                }, inplace=True)
                
                # Grouping Stock
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
                    st.dataframe(deficit_df[cols].style.format({
                        "Total_Ordered": "{:,.0f}", 
                        "Stock_Onhand": "{:,.0f}", 
                        "Balance": "{:,.0f}"
                    }), use_container_width=True)
                    
                    csv_defisit = deficit_df.to_csv(index=False).encode('utf-8')
                    st.download_button("Download Data Defisit (CSV)", csv_defisit, "defisit_stock.csv", "text/csv")
                else:
                    st.success("Aman! Tidak ada defisit stock.")

            # =========================================
            # TAB 2: CEK DETAIL STOCK & SO (Updated)
            # =========================================
            with tab2:
                st.subheader("Cek Ketersediaan & Alokasi Stock")
                st.markdown("Pilih SKU untuk melihat perbandingan **Stock Gudang** vs **Order SO** per Batch.")
                
                all_materials = sorted(df_loct['Material'].unique())
                selected_material = st.selectbox("Pilih Material / SKU:", all_materials)
                
                if selected_material:
                    # 1. Siapkan Data Stock (Gudang) untuk Material ini
                    loct_subset = df_loct[df_loct['Material'] == selected_material]
                    # Group by Batch (antisipasi jika ada double row per batch di file asli)
                    loct_grouped = loct_subset.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                    loct_grouped.rename(columns={'Unrestricted': 'Stock_Gudang'}, inplace=True)
                    
                    # 2. Siapkan Data Order (SO) untuk Material ini
                    so_subset = df_so[df_so['Material'] == selected_material]
                    so_grouped = so_subset.groupby(['Material', 'Batch Number'])['Ordered Quantity'].sum().reset_index()
                    so_grouped.rename(columns={'Batch Number': 'Batch', 'Ordered Quantity': 'Qty_SO'}, inplace=True)
                    
                    # 3. Gabungkan (Left Join ke Stock Gudang)
                    # Kita pakai Left Join ke Loct karena tujuannya mau liat "Stock yg tersedia"
                    final_view = pd.merge(loct_grouped, so_grouped, on=['Material', 'Batch'], how='left')
                    
                    # Isi NaN dengan 0 (karena ada stock tapi gak ada order)
                    final_view['Qty_SO'] = final_view['Qty_SO'].fillna(0)
                    
                    # Hitung Sisa Stock (Balance)
                    final_view['Sisa_Stock'] = final_view['Stock_Gudang'] - final_view['Qty_SO']
                    
                    # Summary Cards
                    tot_stock = final_view['Stock_Gudang'].sum()
                    tot_so = final_view['Qty_SO'].sum()
                    global_bal = tot_stock - tot_so
                    
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total Stock (Gudang)", f"{tot_stock:,.0f}")
                    col2.metric("Total Order (SO)", f"{tot_so:,.0f}")
                    col3.metric("Global Balance", f"{global_bal:,.0f}", delta_color="normal")
                    
                    st.divider()
                    st.write(f"**Detail Alokasi Batch untuk {selected_material}:**")
                    
                    # Tampilkan Tabel
                    # Kolom: Material, Batch, Stock Gudang, Qty SO, Sisa Stock
                    st.dataframe(
                        final_view.style.format({
                            "Stock_Gudang": "{:,.0f}", 
                            "Qty_SO": "{:,.0f}", 
                            "Sisa_Stock": "{:,.0f}"
                        }), 
                        use_container_width=True
                    )
                    
                    st.caption("Catatan: Tabel di atas menampilkan batch yang **fisiknya ada di gudang** (Loct_F211).")

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

else:
    st.info("Silakan upload file Excel di sidebar.")
