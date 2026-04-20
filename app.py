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

def to_excel_detail_so(df_detail):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_detail.to_excel(writer, index=False, sheet_name='Detail SKU per SO')
    processed_data = output.getvalue()
    return processed_data

def to_excel_batch_suggestion(df_suggestion):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_suggestion.to_excel(writer, index=False, sheet_name='Saran Batch untuk SO Tanpa Batch')
    processed_data = output.getvalue()
    return processed_data

# --- FUNGSI COLOR CODING UNTUK STATUS ---
def highlight_status(val):
    if val == '❌ DEFISIT':
        return 'background-color: #ffcccc'
    elif val == '⚠️ PAS':
        return 'background-color: #ffffcc'
    elif val == '✅ SURPLUS':
        return 'background-color: #ccffcc'
    elif val == '⚠️ TANPA BATCH':
        return 'background-color: #ffe6cc'
    return ''

def highlight_kecukupan(val):
    if val == '✅ CUKUP':
        return 'background-color: #ccffcc'
    elif val == '⚠️ KURANG':
        return 'background-color: #ffffcc'
    elif val == '❌ TIDAK ADA STOCK':
        return 'background-color: #ffcccc'
    return ''

def highlight_total_status(val):
    if val == '✅ TOTAL STOCK CUKUP':
        return 'background-color: #ccffcc'
    elif val == '⚠️ TOTAL STOCK KURANG':
        return 'background-color: #ffffcc'
    return ''

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

            # Cek nama kolom Batch di df_so
            if 'Batch Number' not in df_so.columns:
                batch_col = [col for col in df_so.columns if 'batch' in col.lower()]
                if batch_col:
                    batch_col_name = batch_col[0]
                    df_so.rename(columns={batch_col_name: 'Batch Number'}, inplace=True)
                else:
                    st.error("Kolom 'Batch Number' tidak ditemukan di sheet SO_B2B")
                    st.stop()

            # Buat dataframe detail dengan status stock
            loct_batch = df_loct.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
            loct_batch.rename(columns={'Unrestricted': 'Stock_Batch'}, inplace=True)
            
            # Merge dengan penanganan khusus untuk NaN di Batch Number
            df_so_detail = df_so.copy()
            df_so_detail['Batch Number'] = df_so_detail['Batch Number'].fillna('TANPA BATCH')
            
            df_so_detail = df_so_detail.merge(
                loct_batch, 
                left_on=['Material', 'Batch Number'], 
                right_on=['Material', 'Batch'], 
                how='left'
            )
            df_so_detail.drop('Batch_y', axis=1, errors='ignore', inplace=True)
            df_so_detail.rename(columns={'Batch_x': 'Batch Number'}, inplace=True, errors='ignore')
            
            df_so_detail['Stock_Batch'] = df_so_detail['Stock_Batch'].fillna(0)
            df_so_detail['Balance_Per_Line'] = df_so_detail['Stock_Batch'] - df_so_detail['Ordered Quantity']
            
            # Tambah kolom Status
            def get_status_detail(row):
                if row['Batch Number'] == 'TANPA BATCH':
                    return "⚠️ TANPA BATCH"
                elif row['Balance_Per_Line'] < 0:
                    return "❌ DEFISIT"
                elif row['Balance_Per_Line'] == 0:
                    return "⚠️ PAS"
                else:
                    return "✅ SURPLUS"
            
            df_so_detail['Status_Stock'] = df_so_detail.apply(get_status_detail, axis=1)
            
            # Tambah kolom global stock per material
            loct_material = df_loct.groupby('Material')['Unrestricted'].sum().reset_index()
            loct_material.rename(columns={'Unrestricted': 'Total_Stock_Material'}, inplace=True)
            df_so_detail = df_so_detail.merge(loct_material, on='Material', how='left')
            df_so_detail['Total_Stock_Material'] = df_so_detail['Total_Stock_Material'].fillna(0)

            tab1, tab2, tab3 = st.tabs(["🚨 Analisis Defisit & Download", "📋 Detail SKU per SO", "🔍 Cek Detail per SKU"])

            # =========================================
            # TAB 1: HASIL ANALISIS & DOWNLOAD
            # =========================================
            with tab1:
                st.subheader("Analisis Batch Defisit")
                
                # Filter SO yang memiliki batch number saja untuk analisis defisit
                df_so_with_batch = df_so[df_so['Batch Number'].notna()].copy()
                
                if not df_so_with_batch.empty:
                    so_agg = df_so_with_batch.groupby(['Material', 'Batch Number']).agg({
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

                        list_material_defisit = deficit_df['Material'].unique()
                        
                        loct_subset = df_loct[df_loct['Material'].isin(list_material_defisit)]
                        loct_avail = loct_subset.groupby(['Material', 'Batch'])['Unrestricted'].sum().reset_index()
                        loct_avail.rename(columns={'Unrestricted': 'Stock_Gudang'}, inplace=True)
                        
                        so_subset = df_so_with_batch[df_so_with_batch['Material'].isin(list_material_defisit)]
                        so_avail = so_subset.groupby(['Material', 'Batch Number'])['Ordered Quantity'].sum().reset_index()
                        so_avail.rename(columns={'Batch Number': 'Batch', 'Ordered Quantity': 'Qty_SO_Terpakai'}, inplace=True)
                        
                        substitusi_df = pd.merge(loct_avail, so_avail, on=['Material', 'Batch'], how='outer')
                        substitusi_df['Stock_Gudang'] = substitusi_df['Stock_Gudang'].fillna(0)
                        substitusi_df['Qty_SO_Terpakai'] = substitusi_df['Qty_SO_Terpakai'].fillna(0)
                        substitusi_df['Sisa_Stock_Bisa_Pakai'] = substitusi_df['Stock_Gudang'] - substitusi_df['Qty_SO_Terpakai']
                        
                        def get_status(row):
                            if row['Sisa_Stock_Bisa_Pakai'] < 0:
                                return "❌ DEFISIT"
                            elif row['Sisa_Stock_Bisa_Pakai'] == 0:
                                return "⚠️ PAS"
                            else:
                                return "✅ SURPLUS"
                        
                        substitusi_df['Status'] = substitusi_df.apply(get_status, axis=1)
                        
                        status_order = {"❌ DEFISIT": 1, "⚠️ PAS": 2, "✅ SURPLUS": 3}
                        substitusi_df['Sort_Status'] = substitusi_df['Status'].map(status_order)
                        substitusi_df = substitusi_df.sort_values(
                            by=['Material', 'Sort_Status', 'Sisa_Stock_Bisa_Pakai'], 
                            ascending=[True, True, True]
                        ).drop('Sort_Status', axis=1)

                        with st.expander("📊 Lihat Preview Opsi Substitusi (Semua Batch Material Terkait)"):
                            st.caption("Tabel ini menampilkan semua batch dari material yang defisit.")
                            
                            styled_df = substitusi_df.style.map(
                                highlight_status, 
                                subset=['Status']
                            ).format({
                                "Stock_Gudang": "{:,.0f}",
                                "Qty_SO_Terpakai": "{:,.0f}",
                                "Sisa_Stock_Bisa_Pakai": "{:,.0f}"
                            })
                            
                            st.dataframe(styled_df, use_container_width=True)

                        excel_data = to_excel(deficit_df_clean, substitusi_df)
                        
                        st.download_button(
                            label="📥 Download Report Lengkap (.xlsx)",
                            data=excel_data,
                            file_name='Laporan_Analisis_Stock_Defisit.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                        
                    else:
                        st.success("Aman! Tidak ada defisit stock untuk SO yang memiliki Batch Number.")
                else:
                    st.info("Tidak ada SO dengan Batch Number untuk dianalisis.")

            # =========================================
            # TAB 2: DETAIL SKU PER SO
            # =========================================
            with tab2:
                st.subheader("📋 Detail SKU Material per Shipment Number")
                st.caption("Filter berdasarkan Shipment Number untuk melihat detail SKU dan status stock-nya.")
                
                list_so = sorted(df_so_detail['Shipment Number'].unique())
                
                col1, col2 = st.columns([2, 1])
                with col1:
                    selected_so = st.multiselect(
                        "Pilih Shipment Number (bisa pilih lebih dari satu):",
                        options=list_so,
                        help="Kosongkan untuk menampilkan semua SO"
                    )
                
                with col2:
                    status_filter = st.multiselect(
                        "Filter Status Stock:",
                        options=["❌ DEFISIT", "⚠️ PAS", "✅ SURPLUS", "⚠️ TANPA BATCH"],
                        default=["❌ DEFISIT", "⚠️ TANPA BATCH"],
                        help="Pilih status stock yang ingin ditampilkan"
                    )
                
                # Filter data berdasarkan pilihan
                if selected_so:
                    df_filtered = df_so_detail[df_so_detail['Shipment Number'].isin(selected_so)].copy()
                else:
                    df_filtered = df_so_detail.copy()
                
                if status_filter:
                    df_filtered = df_filtered[df_filtered['Status_Stock'].isin(status_filter)]
                
                if not df_filtered.empty:
                    # Tampilkan ringkasan
                    total_lines = len(df_filtered)
                    total_qty = df_filtered['Ordered Quantity'].sum()
                    deficit_lines = len(df_filtered[df_filtered['Status_Stock'] == '❌ DEFISIT'])
                    tanpa_batch_lines = len(df_filtered[df_filtered['Status_Stock'] == '⚠️ TANPA BATCH'])
                    
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Total Line Items", total_lines)
                    col2.metric("Total Quantity", f"{total_qty:,.0f}")
                    col3.metric("Line Defisit", deficit_lines)
                    col4.metric("Tanpa Batch", tanpa_batch_lines)
                    
                    # Siapkan kolom yang ingin ditampilkan
                    cols_display = [
                        'Shipment Number', 
                        'Material', 
                        'Batch Number', 
                        'Ordered Quantity',
                        'Stock_Batch',
                        'Balance_Per_Line',
                        'Total_Stock_Material',
                        'Status_Stock'
                    ]
                    
                    df_display = df_filtered[cols_display].copy()
                    df_display = df_display.sort_values(['Shipment Number', 'Status_Stock', 'Material'])
                    
                    # Tampilkan tabel dengan styling
                    styled_detail = df_display.style.map(
                        highlight_status,
                        subset=['Status_Stock']
                    ).format({
                        "Ordered Quantity": "{:,.0f}",
                        "Stock_Batch": "{:,.0f}",
                        "Balance_Per_Line": "{:,.0f}",
                        "Total_Stock_Material": "{:,.0f}"
                    })
                    
                    st.dataframe(styled_detail, use_container_width=True)
                    
                    # Download button untuk data yang difilter
                    st.download_button(
                        label="📥 Download Detail SKU (Filtered).xlsx",
                        data=to_excel_detail_so(df_display),
                        file_name=f'Detail_SKU_SO_{len(df_display)}_items.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                    
                    # ===== FITUR BARU: SARAN BATCH UNTUK SO TANPA BATCH =====
                    # Ambil data SO yang tidak memiliki batch number dari hasil filter
                    df_tanpa_batch = df_filtered[df_filtered['Batch Number'] == 'TANPA BATCH'].copy()
                    
                    if not df_tanpa_batch.empty:
                        st.markdown("---")
                        st.subheader("🎯 Saran Batch untuk SO yang Belum Ada Batch Number")
                        st.caption("Berikut adalah rekomendasi batch yang available di F211 untuk material yang belum ditentukan batchnya.")
                        
                        # Buat list saran batch
                        list_saran = []
                        
                        for idx, row in df_tanpa_batch.iterrows():
                            material = row['Material']
                            qty_needed = row['Ordered Quantity']
                            so_number = row['Shipment Number']
                            
                            # Ambil batch yang available untuk material tersebut
                            batch_available = df_loct[
                                (df_loct['Material'] == material) & 
                                (df_loct['Unrestricted'] > 0)
                            ].groupby('Batch')['Unrestricted'].sum().reset_index()
                            
                            if not batch_available.empty:
                                # Hitung total stock available
                                batch_available['Stock_Available'] = batch_available['Unrestricted']
                                batch_available['Qty_Dibutuhkan'] = qty_needed
                                batch_available['Shipment_Number'] = so_number
                                batch_available['Material'] = material
                                batch_available['Status_Kecukupan'] = batch_available['Stock_Available'].apply(
                                    lambda x: '✅ CUKUP' if x >= qty_needed else '⚠️ KURANG'
                                )
                                
                                batch_available = batch_available[[
                                    'Shipment_Number', 'Material', 'Batch', 'Stock_Available', 
                                    'Qty_Dibutuhkan', 'Status_Kecukupan'
                                ]]
                                
                                list_saran.append(batch_available)
                            else:
                                # Jika tidak ada stock sama sekali
                                no_stock_row = pd.DataFrame([{
                                    'Shipment_Number': so_number,
                                    'Material': material,
                                    'Batch': 'TIDAK ADA STOCK',
                                    'Stock_Available': 0,
                                    'Qty_Dibutuhkan': qty_needed,
                                    'Status_Kecukupan': '❌ TIDAK ADA STOCK'
                                }])
                                list_saran.append(no_stock_row)
                        
                        if list_saran:
                            df_saran = pd.concat(list_saran, ignore_index=True)
                            
                            # Tampilkan tabel saran - GUNAKAN .map() BUKAN .applymap()
                            styled_saran = df_saran.style.map(
                                highlight_kecukupan,
                                subset=['Status_Kecukupan']
                            ).format({
                                "Stock_Available": "{:,.0f}",
                                "Qty_Dibutuhkan": "{:,.0f}"
                            })
                            
                            st.dataframe(styled_saran, use_container_width=True)
                            
                            # Download button untuk saran batch
                            st.download_button(
                                label="📥 Download Saran Batch untuk SO Tanpa Batch.xlsx",
                                data=to_excel_batch_suggestion(df_saran),
                                file_name=f'Saran_Batch_SO_Tanpa_Batch_{len(df_saran)}_items.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                help="Download rekomendasi batch untuk SO yang belum memiliki batch number"
                            )
                            
                            # Ringkasan per SO
                            with st.expander("📊 Lihat Ringkasan per SO (Tanpa Batch)"):
                                summary_so = df_saran.groupby(['Shipment_Number', 'Material']).agg({
                                    'Qty_Dibutuhkan': 'first',
                                    'Batch': lambda x: ', '.join(x.unique()),
                                    'Stock_Available': 'sum'
                                }).reset_index()
                                
                                summary_so['Status'] = summary_so.apply(
                                    lambda row: '✅ TOTAL STOCK CUKUP' if row['Stock_Available'] >= row['Qty_Dibutuhkan'] 
                                    else '⚠️ TOTAL STOCK KURANG', axis=1
                                )
                                
                                styled_summary = summary_so.style.map(
                                    highlight_total_status,
                                    subset=['Status']
                                ).format({
                                    "Qty_Dibutuhkan": "{:,.0f}",
                                    "Stock_Available": "{:,.0f}"
                                })
                                
                                st.dataframe(styled_summary, use_container_width=True)
                    
                    # Summary per Material (untuk semua data)
                    with st.expander("📊 Lihat Summary per Material"):
                        summary = df_filtered.groupby('Material').agg({
                            'Ordered Quantity': 'sum',
                            'Shipment Number': lambda x: ', '.join(x.unique())
                        }).reset_index()
                        summary.columns = ['Material', 'Total_Qty_SO', 'List_SO']
                        
                        stock_summary = df_filtered.groupby('Material')['Total_Stock_Material'].first().reset_index()
                        summary = summary.merge(stock_summary, on='Material')
                        summary['Balance_Global'] = summary['Total_Stock_Material'] - summary['Total_Qty_SO']
                        
                        def get_status_summary(row):
                            if row['Balance_Global'] < 0:
                                return "❌ DEFISIT"
                            elif row['Balance_Global'] == 0:
                                return "⚠️ PAS"
                            else:
                                return "✅ SURPLUS"
                        
                        summary['Status_Global'] = summary.apply(get_status_summary, axis=1)
                        
                        styled_summary = summary.style.map(
                            highlight_status,
                            subset=['Status_Global']
                        ).format({
                            "Total_Qty_SO": "{:,.0f}",
                            "Total_Stock_Material": "{:,.0f}",
                            "Balance_Global": "{:,.0f}"
                        })
                        
                        st.dataframe(styled_summary, use_container_width=True)
                        
                else:
                    st.warning("Tidak ada data yang sesuai dengan filter yang dipilih.")

            # =========================================
            # TAB 3: CEK DETAIL STOCK & SO
            # =========================================
            with tab3:
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
                    
                    final_view = pd.merge(loct_grouped, so_grouped, on=['Material', 'Batch'], how='outer')
                    final_view['Stock_Gudang'] = final_view['Stock_Gudang'].fillna(0)
                    final_view['Qty_SO'] = final_view['Qty_SO'].fillna(0)
                    final_view['Sisa_Stock'] = final_view['Stock_Gudang'] - final_view['Qty_SO']
                    
                    def get_status_final(row):
                        if row['Sisa_Stock'] < 0:
                            return "❌ DEFISIT"
                        elif row['Sisa_Stock'] == 0:
                            return "⚠️ PAS"
                        else:
                            return "✅ SURPLUS"
                    
                    final_view['Status'] = final_view.apply(get_status_final, axis=1)
                    
                    tot_stock = final_view['Stock_Gudang'].sum()
                    tot_so = final_view['Qty_SO'].sum()
                    
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total Stock", f"{tot_stock:,.0f}")
                    col2.metric("Total Order", f"{tot_so:,.0f}")
                    col3.metric("Balance Global", f"{tot_stock - tot_so:,.0f}")
                    
                    styled_final = final_view.style.map(
                        highlight_status,
                        subset=['Status']
                    ).format({
                        "Stock_Gudang": "{:,.0f}", 
                        "Qty_SO": "{:,.0f}", 
                        "Sisa_Stock": "{:,.0f}"
                    })
                    
                    st.dataframe(styled_final, use_container_width=True)
                    
                    with st.expander("📋 Lihat Detail SO untuk Material ini"):
                        detail_material = df_so_detail[df_so_detail['Material'] == selected_material][
                            ['Shipment Number', 'Batch Number', 'Ordered Quantity', 'Stock_Batch', 'Balance_Per_Line', 'Status_Stock']
                        ].sort_values('Shipment Number')
                        
                        styled_detail_mat = detail_material.style.map(
                            highlight_status,
                            subset=['Status_Stock']
                        ).format({
                            "Ordered Quantity": "{:,.0f}",
                            "Stock_Batch": "{:,.0f}",
                            "Balance_Per_Line": "{:,.0f}"
                        })
                        
                        st.dataframe(styled_detail_mat, use_container_width=True)

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
            import traceback
            st.code(traceback.format_exc())

else:
    st.info("Silakan upload file Excel di sidebar.")
