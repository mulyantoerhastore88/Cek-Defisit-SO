import streamlit as st
import pandas as pd
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Dashboard Analisis Defisit Stock SO - Mulyanto Demand Planner", layout="wide")

st.title("📦 Dashboard Analisis Defisit Stock SO - Mulyanto Demand Planner ")
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

def clean_number(x):
    if isinstance(x, str):
        x = x.replace(',', '') 
    return pd.to_numeric(x, errors='coerce')

# --- FUNGSI EXPORT EXCEL ---
def to_excel(df_defisit, df_substitusi):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_defisit.to_excel(writer, index=False, sheet_name='Data Defisit (Action Needed)')
        df_substitusi.to_excel(writer, index=False, sheet_name='Opsi Substitusi (Stock Tersedia)')
    processed_data = output.getvalue()
    return processed_data

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
            
            if df_so['Ordered Quantity'].dtype == 'object':
                df_so['Ordered Quantity'] = df_so['Ordered Quantity'].apply(clean_number)
            
            if df_loct['Unrestricted'].dtype == 'object':
                df_loct['Unrestricted'] = df_loct['Unrestricted'].apply(clean_number)

            tab1, tab2 = st.tabs(["🚨 Analisis Defisit & Download", "🔍 Cek Detail per SKU"])

            # =========================================
            # TAB 1: HASIL ANALISIS & DOWNLOAD
            # =========================================
            with tab1:
                st.subheader("Analisis Batch Defisit")
                
                # --- STEP 1: Hitung Defisit Utama ---
                so_agg = df_so.groupby(['Material', 'Batch Number']).agg({
                    'Ordered Quantity': 'sum',
                    'Shipment Number': lambda x: ', '.join(x.astype(str).unique()) 
                }).reset_index()
                
                so_agg.rename(columns={
                    'Batch Number': 'Batch', 
                    'Ordered Quantity': 'Total_Ordered',
                    'Shipment Number': 'List_Shipment_Numbers'
                }, inplace=True)
                
                loct_agg = df_loct.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                loct_agg.rename(columns={'Unrestricted': 'Stock_Onhand'}, inplace=True)
                
                merged_df = pd.merge(so_agg, loct_agg, on=['Material', 'Batch'], how='left')
                merged_df['Stock_Onhand'] = merged_df['Stock_Onhand'].fillna(0)
                merged_df['Balance'] = merged_df['Stock_Onhand'] - merged_df['Total_Ordered']
                
                # Filter hanya yang defisit
                deficit_df = merged_df[merged_df['Balance'] < 0].copy()
                
                if not deficit_df.empty:
                    cols = ['Material', 'Batch', 'Total_Ordered', 'Stock_Onhand', 'Balance', 'List_Shipment_Numbers']
                    deficit_df_clean = deficit_df[cols]

                    st.error(f"Ditemukan {len(deficit_df_clean)} Batch SKU yang defisit!")
                    st.dataframe(deficit_df_clean.style.format({
                        "Total_Ordered": "{:,.0f}", 
                        "Stock_Onhand": "{:,.0f}", 
                        "Balance": "{:,.0f}"
                    }), use_container_width=True)

                    # --- STEP 2: Siapkan Data Substitusi (Updated Logic) ---
                    # Ambil list material yang bermasalah
                    list_material_defisit = deficit_df['Material'].unique()
                    
                    # 2a. Ambil Stock Gudang untuk material tsb
                    loct_subset = df_loct[df_loct['Material'].isin(list_material_defisit)]
                    loct_avail = loct_subset.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                    loct_avail.rename(columns={'Unrestricted': 'Stock_Gudang'}, inplace=True)
                    
                    # 2b. Ambil Total SO untuk material tsb (Biar tau batch ini dipake siapa aja)
                    so_subset = df_so[df_so['Material'].isin(list_material_defisit)]
                    so_avail = so_subset.groupby(['Material', 'Batch Number'])['Ordered Quantity'].sum().reset_index()
                    so_avail.rename(columns={'Batch Number': 'Batch', 'Ordered Quantity': 'Qty_SO_Terpakai'}, inplace=True)
                    
                    # 2c. Gabungkan (Left Join ke Stock Gudang)
                    # Kenapa Left Join ke Stock? Karena kita mau cari barang yang ADA fisiknya buat subtitusi.
                    substitusi_df = pd.merge(loct_avail, so_avail, on=['Material', 'Batch'], how='left')
                    substitusi_df['Qty_SO_Terpakai'] = substitusi_df['Qty_SO_Terpakai'].fillna(0)
                    
                    # 2d. Hitung Sisa Stock yang Beneran Free
                    substitusi_df['Sisa_Stock_Bisa_Pakai'] = substitusi_df['Stock_Gudang'] - substitusi_df['Qty_SO_Terpakai']
                    
                    # Urutkan biar enak bacanya (Material dulu, lalu Sisa Stock terbesar di atas)
                    substitusi_df = substitusi_df.sort_values(by=['Material', 'Sisa_Stock_Bisa_Pakai'], ascending=[True, False])

                    # Tampilkan Preview
                    with st.expander("Lihat Preview Opsi Substitusi (Batch dengan Stock Positif)"):
                        st.dataframe(substitusi_df.style.format({
                            "Stock_Gudang": "{:,.0f}",
                            "Qty_SO_Terpakai": "{:,.0f}",
                            "Sisa_Stock_Bisa_Pakai": "{:,.0f}"
                        }), use_container_width=True)

                    # --- STEP 3: Generate File Excel ---
                    excel_data = to_excel(deficit_df_clean, substitusi_df)
                    
                    st.download_button(
                        label="📥 Download Report Lengkap (.xlsx)",
                        data=excel_data,
                        file_name='Laporan_Analisis_Stock_Defisit.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        help="Sheet 1: List Defisit. Sheet 2: Stock Tersedia lengkap dengan info pemakaian SO."
                    )
                    
                else:
                    st.success("Aman! Tidak ada defisit stock.")

            # =========================================
            # TAB 2: CEK DETAIL STOCK & SO
            # =========================================
            with tab2:
                st.subheader("Cek Ketersediaan & Alokasi Stock per SKU")
                
                all_materials = sorted(df_loct['Material'].unique())
                selected_material = st.selectbox("Pilih Material / SKU:", all_materials)
                
                if selected_material:
                    loct_subset = df_loct[df_loct['Material'] == selected_material]
                    loct_grouped = loct_subset.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                    loct_grouped.rename(columns={'Unrestricted': 'Stock_Gudang'}, inplace=True)
                    
                    so_subset = df_so[df_so['Material'] == selected_material]
                    so_grouped = so_subset.groupby(['Material', 'Batch Number'])['Ordered Quantity'].sum().reset_index()
                    so_grouped.rename(columns={'Batch Number': 'Batch', 'Ordered Quantity': 'Qty_SO'}, inplace=True)
                    
                    final_view = pd.merge(loct_grouped, so_grouped, on=['Material', 'Batch'], how='left')
                    final_view['Qty_SO'] = final_view['Qty_SO'].fillna(0)
                    final_view['Sisa_Stock'] = final_view['Stock_Gudang'] - final_view['Qty_SO']
                    
                    tot_stock = final_view['Stock_Gudang'].sum()
                    tot_so = final_view['Qty_SO'].sum()
                    
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total Stock", f"{tot_stock:,.0f}")
                    col2.metric("Total Order", f"{tot_so:,.0f}")
                    col3.metric("Balance Global", f"{tot_stock - tot_so:,.0f}")
                    
                    st.dataframe(final_view.style.format({
                        "Stock_Gudang": "{:,.0f}", "Qty_SO": "{:,.0f}", "Sisa_Stock": "{:,.0f}"
                    }), use_container_width=True)

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

else:
    st.info("Silakan upload file Excel di sidebar.")
