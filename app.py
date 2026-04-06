import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import warnings
warnings.filterwarnings('ignore')

# --- Konfigurasi Halaman ---
st.set_page_config(
    page_title="SKU Evaluator Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom CSS Premium (Diadopsi dari dashboard sebelumnya) ---
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 900;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 1rem;
        text-align: center;
    }
    .sku-header {
        background-color: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        border-left: 6px solid #6366F1;
        margin-bottom: 1.5rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
    }
    .sku-title-box { flex: 2; min-width: 300px; }
    .sku-title { font-size: 1.4rem; font-weight: 800; color: #1F2937; margin-bottom: 0.5rem; }
    .sku-badges { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 10px; }
    .badge { padding: 4px 12px; border-radius: 20px; font-size: 0.8rem; font-weight: 600; display: flex; align-items: center; gap: 5px; }
    .badge-blue { background: #E0E7FF; color: #4338CA; }
    .badge-purple { background: #F3E8FF; color: #7E22CE; }
    .badge-gray { background: #F3F4F6; color: #4B5563; }
    .sku-fin-box { flex: 1; text-align: right; min-width: 200px; border-left: 1px solid #E5E7EB; padding-left: 20px; }
    .fin-label { font-size: 0.8rem; color: #6B7280; font-weight: 600; text-transform: uppercase; }
    .fin-val-big { font-size: 1.5rem; font-weight: 800; color: #10B981; }
    .metric-card {
        border-radius: 12px; padding: 1.2rem; color: white;
        box-shadow: 0 4px 10px rgba(0,0,0,0.05); transition: transform 0.3s;
    }
    .metric-card:hover { transform: translateY(-3px); }
    .metric-label { font-size: 0.8rem; font-weight: 700; opacity: 0.9; text-transform: uppercase; margin-bottom: 5px; }
    .metric-value { font-size: 1.6rem; font-weight: 800; margin: 5px 0; text-shadow: 0 1px 2px rgba(0,0,0,0.1); }
    .metric-sub { font-size: 0.85rem; font-weight: 500; opacity: 0.95; }
    .diagnostic-box {
        background: #F8FAFC; border-left: 4px solid #EF4444; padding: 12px; border-radius: 8px; margin-bottom: 10px;
    }
    .diagnostic-title { font-weight: 700; color: #374151; display: flex; align-items: center; gap: 8px; }
    .diagnostic-desc { font-size: 0.9rem; color: #4B5563; margin-left: 32px; }
</style>
""", unsafe_allow_html=True)

# --- Judul Dashboard ---
st.markdown('<h1 class="main-header">📊 SKU 360° Evaluator Pro</h1>', unsafe_allow_html=True)
st.caption(f"🚀 Evaluasi Performa SKU: Sales vs PO | Terakhir diperbarui: {datetime.now().strftime('%d %B %Y %H:%M')}")

# --- Koneksi ke Google Sheets Baru ---
@st.cache_resource(show_spinner=False)
def init_gsheet_connection():
    try:
        skey = st.secrets["gcp_service_account"]
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        credentials = Credentials.from_service_account_info(skey, scopes=scopes)
        client = gspread.authorize(credentials)
        return client
    except Exception as e:
        st.error(f"❌ Koneksi Gagal: {str(e)}")
        return None

@st.cache_data(ttl=300, show_spinner=False)
def load_new_data(_client):
    """
    Load data dari 4 sheet sesuai spesifikasi:
    - Product_Master
    - Data_PO
    - Sales_2025
    - Sales_2026
    """
    gsheet_url = "https://docs.google.com/spreadsheets/d/1REhZBDsFXLlCgKJbKKilRKrIEPcBzvA6EQRPY_XhZPg/edit?gid=2062248078#gid=2062248078"
    data = {}
    
    try:
        # 1. Product Master
        ws_prod = _client.open_by_url(gsheet_url).worksheet("Product_Master")
        df_product = pd.DataFrame(ws_prod.get_all_records())
        df_product.columns = [col.strip().replace(' ', '_') for col in df_product.columns]
        
        # Pastikan kolom harga numerik
        for col in ['Floor_Price', 'Purchase_Order_Price']:
            if col in df_product.columns:
                df_product[col] = pd.to_numeric(df_product[col], errors='coerce').fillna(0)
        
        # Filter Active SKUs
        if 'Status' in df_product.columns:
            df_product['Status'] = df_product['Status'].astype(str).str.upper()
            df_product_active = df_product[df_product['Status'] == 'ACTIVE'].copy()
        else:
            df_product_active = df_product.copy()
            df_product_active['Status'] = 'ACTIVE'
            
        data['product'] = df_product
        data['product_active'] = df_product_active
        
        # 2. Data PO (Purchase Order)
        ws_po = _client.open_by_url(gsheet_url).worksheet("Data_PO")
        df_po_raw = pd.DataFrame(ws_po.get_all_records())
        df_po_raw.columns = [col.strip() for col in df_po_raw.columns]
        
        # Identifikasi kolom bulan (semua kolom setelah SKU_ID)
        month_cols_po = [c for c in df_po_raw.columns if c != 'SKU_ID' and any(m in c.upper() for m in ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'])]
        
        if month_cols_po and 'SKU_ID' in df_po_raw.columns:
            df_po_long = df_po_raw.melt(id_vars=['SKU_ID'], value_vars=month_cols_po, var_name='Month_Label', value_name='PO_Qty')
            df_po_long['PO_Qty'] = pd.to_numeric(df_po_long['PO_Qty'], errors='coerce').fillna(0)
            df_po_long['Month'] = df_po_long['Month_Label'].apply(parse_month)
            df_po_long = df_po_long[df_po_long['SKU_ID'].isin(df_product_active['SKU_ID'])]
            data['po'] = df_po_long
        else:
            data['po'] = pd.DataFrame()
        
        # 3. Sales 2025
        ws_sales25 = _client.open_by_url(gsheet_url).worksheet("Sales_2025")
        df_sales25_raw = pd.DataFrame(ws_sales25.get_all_records())
        df_sales25_raw.columns = [col.strip() for col in df_sales25_raw.columns]
        
        # Kolom bulan di Sales 2025
        month_cols_25 = [c for c in df_sales25_raw.columns if any(m in c.upper() for m in ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'])]
        
        if month_cols_25:
            id_cols = ['OLD_Material', 'SKU Name', 'Group', 'Brand', 'SKU Tier']
            id_cols = [c for c in id_cols if c in df_sales25_raw.columns]
            df_sales25_long = df_sales25_raw.melt(id_vars=id_cols, value_vars=month_cols_25, var_name='Month_Label', value_name='Sales_Qty')
            df_sales25_long['Sales_Qty'] = pd.to_numeric(df_sales25_long['Sales_Qty'], errors='coerce').fillna(0)
            df_sales25_long['Month'] = df_sales25_long['Month_Label'].apply(parse_month)
            df_sales25_long['Year'] = 2025
            
            # Mapping OLD_Material ke SKU_ID (jika perlu)
            if 'OLD_Material' in df_sales25_long.columns:
                # Coba mapping dari Product Master
                if 'OLD_Material' in df_product.columns:
                    sku_mapping = df_product[['OLD_Material', 'SKU_ID']].drop_duplicates()
                    df_sales25_long = pd.merge(df_sales25_long, sku_mapping, on='OLD_Material', how='left')
                else:
                    df_sales25_long['SKU_ID'] = df_sales25_long['OLD_Material']
            
            data['sales_2025'] = df_sales25_long
        else:
            data['sales_2025'] = pd.DataFrame()
        
        # 4. Sales 2026
        ws_sales26 = _client.open_by_url(gsheet_url).worksheet("Sales_2026")
        df_sales26_raw = pd.DataFrame(ws_sales26.get_all_records())
        df_sales26_raw.columns = [col.strip() for col in df_sales26_raw.columns]
        
        month_cols_26 = [c for c in df_sales26_raw.columns if any(m in c.upper() for m in ['JAN','FEB','MAR'])]
        
        if month_cols_26:
            id_cols = ['OLD_Material', 'SKU Name', 'Group', 'Brand', 'SKU Tier']
            id_cols = [c for c in id_cols if c in df_sales26_raw.columns]
            df_sales26_long = df_sales26_raw.melt(id_vars=id_cols, value_vars=month_cols_26, var_name='Month_Label', value_name='Sales_Qty')
            df_sales26_long['Sales_Qty'] = pd.to_numeric(df_sales26_long['Sales_Qty'], errors='coerce').fillna(0)
            df_sales26_long['Month'] = df_sales26_long['Month_Label'].apply(parse_month)
            df_sales26_long['Year'] = 2026
            
            if 'OLD_Material' in df_sales26_long.columns:
                if 'OLD_Material' in df_product.columns:
                    sku_mapping = df_product[['OLD_Material', 'SKU_ID']].drop_duplicates()
                    df_sales26_long = pd.merge(df_sales26_long, sku_mapping, on='OLD_Material', how='left')
                else:
                    df_sales26_long['SKU_ID'] = df_sales26_long['OLD_Material']
            
            data['sales_2026'] = df_sales26_long
        else:
            data['sales_2026'] = pd.DataFrame()
        
        # Gabungkan Sales 2025 & 2026
        sales_list = []
        if not data['sales_2025'].empty:
            sales_list.append(data['sales_2025'])
        if not data['sales_2026'].empty:
            sales_list.append(data['sales_2026'])
        
        if sales_list:
            data['sales'] = pd.concat(sales_list, ignore_index=True)
        else:
            data['sales'] = pd.DataFrame()
        
        return data
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return {}

def parse_month(month_str):
    """Parse berbagai format bulan ke datetime"""
    if pd.isna(month_str):
        return datetime.now()
    
    month_str = str(month_str).strip()
    
    # Mapping bulan
    month_map = {
        'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
        'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
    }
    
    # Coba format standar
    formats = ['%b %Y', '%b-%Y', '%b-%y', '%b %y', '%B %Y']
    for fmt in formats:
        try:
            return datetime.strptime(month_str, fmt)
        except:
            continue
    
    # Coba ekstrak manual
    for month_name, month_num in month_map.items():
        if month_name in month_str.upper():
            # Cari tahun
            year = datetime.now().year
            numbers = ''.join([c for c in month_str if c.isdigit()])
            if numbers:
                if len(numbers) == 2:
                    year = 2000 + int(numbers)
                else:
                    year = int(numbers)
            return datetime(year, month_num, 1)
    
    return datetime.now()

def calculate_sku_metrics(df_sales, df_po, df_product, sku_id):
    """Hitung semua metrik untuk SKU tertentu"""
    
    metrics = {
        'sales_data': pd.DataFrame(),
        'po_data': pd.DataFrame(),
        'total_sales': 0,
        'total_po': 0,
        'avg_monthly_sales': 0,
        'coverage_months': 0,
        'last_sales_month': None,
        'last_po_month': None,
        'months_with_sales': 0,
        'months_with_po': 0
    }
    
    # Filter data untuk SKU ini
    sales_sku = df_sales[df_sales['SKU_ID'] == sku_id].copy()
    po_sku = df_po[df_po['SKU_ID'] == sku_id].copy()
    
    if sales_sku.empty and po_sku.empty:
        return metrics
    
    # Aggregasi per bulan
    if not sales_sku.empty:
        sales_monthly = sales_sku.groupby('Month')['Sales_Qty'].sum().reset_index()
        sales_monthly = sales_monthly.sort_values('Month')
        metrics['sales_data'] = sales_monthly
        metrics['total_sales'] = sales_monthly['Sales_Qty'].sum()
        metrics['months_with_sales'] = len(sales_monthly)
        if not sales_monthly.empty:
            metrics['last_sales_month'] = sales_monthly['Month'].max()
            metrics['avg_monthly_sales'] = sales_monthly['Sales_Qty'].mean()
    
    if not po_sku.empty:
        po_monthly = po_sku.groupby('Month')['PO_Qty'].sum().reset_index()
        po_monthly = po_monthly.sort_values('Month')
        metrics['po_data'] = po_monthly
        metrics['total_po'] = po_monthly['PO_Qty'].sum()
        metrics['months_with_po'] = len(po_monthly)
        if not po_monthly.empty:
            metrics['last_po_month'] = po_monthly['Month'].max()
    
    # Hitung coverage (asumsi: stok = total PO, kecepatan jual = avg monthly sales)
    if metrics['avg_monthly_sales'] > 0:
        metrics['coverage_months'] = metrics['total_po'] / metrics['avg_monthly_sales']
    else:
        metrics['coverage_months'] = 999 if metrics['total_po'] > 0 else 0
    
    return metrics

def format_rupiah(value):
    """Format Rupiah dengan masking sederhana"""
    if pd.isna(value) or value == 0:
        return "Rp 0"
    if value >= 1_000_000_000:
        return f"Rp {value/1e9:,.1f} M"
    elif value >= 1_000_000:
        return f"Rp {value/1e6:,.1f} Jt"
    else:
        return f"Rp {value:,.0f}"

# --- MAIN APP ---
client = init_gsheet_connection()
if client is None:
    st.stop()

with st.spinner('🔄 Loading data dari Google Sheets...'):
    all_data = load_new_data(client)
    
    df_product = all_data.get('product', pd.DataFrame())
    df_product_active = all_data.get('product_active', pd.DataFrame())
    df_sales = all_data.get('sales', pd.DataFrame())
    df_po = all_data.get('po', pd.DataFrame())

# --- CEK DATA ---
if df_product_active.empty:
    st.error("❌ Tidak ada data Product Master. Periksa sheet 'Product_Master'.")
    st.stop()

if df_sales.empty:
    st.warning("⚠️ Data Sales kosong. Analisis akan terbatas.")
if df_po.empty:
    st.warning("⚠️ Data PO kosong. Analisis akan terbatas.")

# --- SIDEBAR CONTROLS ---
with st.sidebar:
    st.markdown("### ⚙️ Kontrol Dashboard")
    
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    
    st.markdown("---")
    st.markdown("### 📊 Filter SKU")
    
    # Pilihan filter
    all_brands = ['Semua Brand'] + sorted(df_product_active['Brand'].dropna().unique().tolist()) if 'Brand' in df_product_active.columns else ['Semua Brand']
    selected_brand = st.selectbox("Brand", all_brands)
    
    all_tiers = ['Semua Tier'] + sorted(df_product_active['SKU_Tier'].dropna().unique().tolist()) if 'SKU_Tier' in df_product_active.columns else ['Semua Tier']
    selected_tier = st.selectbox("SKU Tier", all_tiers)
    
    # Filter SKU berdasarkan pilihan
    filtered_skus = df_product_active.copy()
    if selected_brand != 'Semua Brand':
        filtered_skus = filtered_skus[filtered_skus['Brand'] == selected_brand]
    if selected_tier != 'Semua Tier':
        filtered_skus = filtered_skus[filtered_skus['SKU_Tier'] == selected_tier]
    
    # Dropdown SKU
    sku_options = filtered_skus.apply(
        lambda x: f"{x['SKU_ID']} - {x.get('Product_Name', '')}", axis=1
    ).tolist()
    
    selected_sku_display = st.selectbox("Pilih SKU", sorted(sku_options))
    
    st.markdown("---")
    st.markdown("### 📈 Statistik Dataset")
    st.metric("Total Active SKU", len(df_product_active))
    st.metric("SKU dengan Sales", df_sales['SKU_ID'].nunique() if not df_sales.empty else 0)
    st.metric("SKU dengan PO", df_po['SKU_ID'].nunique() if not df_po.empty else 0)

# --- MAIN CONTENT ---
if selected_sku_display:
    selected_sku = selected_sku_display.split(" - ")[0].strip()
    
    # Ambil data SKU dari Product Master
    sku_master = df_product_active[df_product_active['SKU_ID'] == selected_sku]
    if sku_master.empty:
        st.error(f"SKU {selected_sku} tidak ditemukan di Product Master")
        st.stop()
    
    sku_master = sku_master.iloc[0]
    product_name = sku_master.get('Product_Name', 'Unknown')
    brand = sku_master.get('Brand', 'Unknown')
    tier = sku_master.get('SKU_Tier', 'Standard')
    floor_price = sku_master.get('Floor_Price', 0)
    purchase_price = sku_master.get('Purchase_Order_Price', 0)
    moq = sku_master.get('MOQ', 0)
    
    # Hitung metrik SKU
    metrics = calculate_sku_metrics(df_sales, df_po, df_product_active, selected_sku)
    
    # --- HEADER SKU ---
    st.markdown(f"""
    <div class="sku-header">
        <div class="sku-title-box">
            <div class="sku-title">{product_name} <span style="font-weight:400; font-size:1rem; color:#6B7280;">({selected_sku})</span></div>
            <div class="sku-badges">
                <span class="badge badge-blue">🏷️ {brand}</span>
                <span class="badge badge-purple">💎 {tier}</span>
                <span class="badge badge-gray">MOQ: {moq:,.0f}</span>
            </div>
        </div>
        <div class="sku-fin-box">
            <div class="fin-label">Financial Projection</div>
            <div class="fin-val-big">{format_rupiah(metrics['total_po'] * floor_price)}</div>
            <div class="fin-margin">Potential Revenue</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # --- METRIC CARDS ---
    col1, col2, col3, col4 = st.columns(4)
    
    # Warna coverage
    cover_color = "linear-gradient(135deg, #10B981 0%, #059669 100%)"
    if metrics['coverage_months'] < 1.5:
        cover_color = "linear-gradient(135deg, #EF4444 0%, #B91C1C 100%)"
    elif metrics['coverage_months'] > 6:
        cover_color = "linear-gradient(135deg, #F59E0B 0%, #D97706 100%)"
    
    with col1:
        st.markdown(f"""
        <div class="metric-card" style="background: linear-gradient(135deg, #6366F1 0%, #4338CA 100%);">
            <div class="metric-label">Total Sales (All Time)</div>
            <div class="metric-value">{metrics['total_sales']:,.0f}</div>
            <div class="metric-sub">Units Terjual</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card" style="background: linear-gradient(135deg, #F59E0B 0%, #D97706 100%);">
            <div class="metric-label">Total PO (Order)</div>
            <div class="metric-value">{metrics['total_po']:,.0f}</div>
            <div class="metric-sub">Units Dipesan</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card" style="background: {cover_color};">
            <div class="metric-label">Stock Coverage</div>
            <div class="metric-value">{metrics['coverage_months']:.1f} Mo</div>
            <div class="metric-sub">Dengan rata-rata sales</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        avg_sales = metrics['avg_monthly_sales']
        st.markdown(f"""
        <div class="metric-card" style="background: linear-gradient(135deg, #0EA5E9 0%, #0284C7 100%);">
            <div class="metric-label">Avg Monthly Sales</div>
            <div class="metric-value">{avg_sales:.0f}</div>
            <div class="metric-sub">Rata-rata per bulan</div>
        </div>
        """, unsafe_allow_html=True)
    
    # --- TREND CHART (Sales vs PO) ---
    st.subheader("📈 Tren Sales vs Purchase Order")
    
    # Gabungkan data sales dan PO per bulan
    combined_data = []
    
    # Sales data
    if not metrics['sales_data'].empty:
        for _, row in metrics['sales_data'].iterrows():
            combined_data.append({
                'Month': row['Month'],
                'Month_Label': row['Month'].strftime('%b-%y'),
                'Sales': row['Sales_Qty'],
                'PO': 0
            })
    
    # PO data
    if not metrics['po_data'].empty:
        for _, row in metrics['po_data'].iterrows():
            # Cek apakah bulan sudah ada
            existing = next((x for x in combined_data if x['Month'] == row['Month']), None)
            if existing:
                existing['PO'] = row['PO_Qty']
            else:
                combined_data.append({
                    'Month': row['Month'],
                    'Month_Label': row['Month'].strftime('%b-%y'),
                    'Sales': 0,
                    'PO': row['PO_Qty']
                })
    
    if combined_data:
        df_trend = pd.DataFrame(combined_data)
        df_trend = df_trend.sort_values('Month')
        
        fig = go.Figure()
        
        # Bar untuk Sales
        fig.add_trace(go.Bar(
            x=df_trend['Month_Label'],
            y=df_trend['Sales'],
            name='Sales Aktual',
            marker_color='#10B981',
            text=df_trend['Sales'].apply(lambda x: f"{x:,.0f}"),
            textposition='auto'
        ))
        
        # Bar untuk PO
        fig.add_trace(go.Bar(
            x=df_trend['Month_Label'],
            y=df_trend['PO'],
            name='Purchase Order',
            marker_color='#F59E0B',
            text=df_trend['PO'].apply(lambda x: f"{x:,.0f}"),
            textposition='auto'
        ))
        
        fig.update_layout(
            height=450,
            barmode='group',
            xaxis_title='Bulan',
            yaxis_title='Quantity (Units)',
            hovermode='x unified',
            plot_bgcolor='white',
            legend=dict(orientation='h', y=1.1)
        )
        
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("📊 Tidak ada data Sales atau PO untuk SKU ini")
    
    # --- DIAGNOSTIK & REKOMENDASI ---
    st.subheader("🩺 Smart Diagnostics & Rekomendasi")
    
    diagnoses = []
    
    # 1. Analisis Coverage
    if metrics['coverage_months'] < 1.5 and metrics['total_po'] > 0:
        diagnoses.append(("🔴", "High Stockout Risk", f"Stock hanya cukup untuk {metrics['coverage_months']:.1f} bulan dengan sales rata-rata. Segera lakukan replenishment."))
    elif metrics['coverage_months'] > 6 and metrics['total_po'] > 0:
        diagnoses.append(("🟡", "Overstock Alert", f"Stock mencukupi hingga {metrics['coverage_months']:.1f} bulan. Pertimbangkan hold PO atau jalankan promosi."))
    elif metrics['total_po'] > 0:
        diagnoses.append(("🟢", "Healthy Inventory", "Level stok optimal. Lanjutkan monitoring rutin."))
    else:
        diagnoses.append(("⚪", "No PO Data", "Tidak ada data Purchase Order untuk SKU ini."))
    
    # 2. Analisis Sales Velocity
    if metrics['months_with_sales'] >= 3:
        # Cek tren 3 bulan terakhir vs sebelumnya
        sales_df = metrics['sales_data'].sort_values('Month')
        if len(sales_df) >= 6:
            recent_3 = sales_df.tail(3)['Sales_Qty'].mean()
            prev_3 = sales_df.iloc[-6:-3]['Sales_Qty'].mean() if len(sales_df) >= 6 else recent_3
            if prev_3 > 0:
                growth = (recent_3 - prev_3) / prev_3 * 100
                if growth > 30:
                    diagnoses.append(("🚀", "Surging Demand", f"Sales meningkat {growth:.1f}% dalam 3 bulan terakhir. Siapkan stok tambahan."))
                elif growth < -30:
                    diagnoses.append(("📉", "Declining Sales", f"Sales menurun {abs(growth):.1f}%. Evaluasi ulang forecast dan stok."))
    elif metrics['months_with_sales'] == 0:
        diagnoses.append(("⚠️", "No Sales History", "SKU ini belum pernah terjual. Order dengan hati-hati (trial order)."))
    
    # 3. Analisis Gap Sales vs PO
    if metrics['total_po'] > 0 and metrics['total_sales'] > 0:
        sell_through = (metrics['total_sales'] / metrics['total_po'] * 100)
        if sell_through < 50:
            diagnoses.append(("📦", "Low Sell-Through", f"Hanya {sell_through:.1f}% dari total PO yang terjual. Risiko dead stock tinggi."))
        elif sell_through > 100:
            diagnoses.append(("🔥", "High Demand", f"Sales melebihi PO ({sell_through:.0f}%). Potensi lost sales."))
    
    # Tampilkan diagnosa
    for icon, title, desc in diagnoses:
        bg_color = "#F0FDF4" if icon == "🟢" else "#FEF2F2" if icon == "🔴" else "#FFFBEB"
        border_color = "#22C55E" if icon == "🟢" else "#EF4444" if icon == "🔴" else "#F59E0B"
        st.markdown(f"""
        <div class="diagnostic-box" style="background:{bg_color}; border-left-color:{border_color};">
            <div class="diagnostic-title">
                <span style="font-size:1.2rem;">{icon}</span> {title}
            </div>
            <div class="diagnostic-desc">{desc}</div>
        </div>
        """, unsafe_allow_html=True)
    
    # --- DETAIL TABEL ---
    st.divider()
    with st.expander("📋 Lihat Detail Data per Bulan", expanded=False):
        if combined_data:
            detail_df = pd.DataFrame(combined_data).sort_values('Month')
            detail_df['Month'] = detail_df['Month'].dt.strftime('%b %Y')
            detail_df['Gap (PO - Sales)'] = detail_df['PO'] - detail_df['Sales']
            detail_df['Sell Through %'] = detail_df.apply(
                lambda x: (x['Sales'] / x['PO'] * 100) if x['PO'] > 0 else 0, axis=1
            )
            
            # Format
            detail_df['Gap (PO - Sales)'] = detail_df['Gap (PO - Sales)'].apply(lambda x: f"{x:+,.0f}")
            detail_df['Sell Through %'] = detail_df['Sell Through %'].apply(lambda x: f"{x:.1f}%")
            
            st.dataframe(detail_df, use_container_width=True, hide_index=True)
    
    # --- FINANCIAL SUMMARY ---
    st.divider()
    st.subheader("💰 Financial Summary")
    
    fin_col1, fin_col2, fin_col3 = st.columns(3)
    
    potential_revenue = metrics['total_po'] * floor_price
    potential_cogs = metrics['total_po'] * purchase_price
    potential_margin = potential_revenue - potential_cogs
    
    with fin_col1:
        st.metric("Potential Revenue (dari PO)", format_rupiah(potential_revenue))
    with fin_col2:
        st.metric("Estimated COGS", format_rupiah(potential_cogs))
    with fin_col3:
        margin_pct = (potential_margin / potential_revenue * 100) if potential_revenue > 0 else 0
        st.metric("Potential Gross Margin", format_rupiah(potential_margin), delta=f"{margin_pct:.1f}%")
    
    # Peringatan jika data harga tidak lengkap
    if floor_price == 0 or purchase_price == 0:
        st.warning("⚠️ Data harga (Floor_Price atau Purchase_Order_Price) tidak lengkap di Product Master. Hitungan finansial bersifat estimasi.")

else:
    st.info("👈 Silakan pilih SKU dari sidebar untuk memulai analisis.")

# --- FOOTER ---
st.divider()
st.markdown("""
<div style="text-align: center; color: #666; font-size: 0.9rem; padding: 1rem;">
    <p>🚀 <strong>SKU 360° Evaluator Pro v1.0</strong> | Analisis Performa SKU berbasis Sales vs PO</p>
    <p>📊 Data Sales 2025-2026 | Data PO hingga Mar 2026</p>
</div>
""", unsafe_allow_html=True)
