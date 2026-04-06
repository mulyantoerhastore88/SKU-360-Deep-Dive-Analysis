import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import warnings
warnings.filterwarnings('ignore')

# --- Konfigurasi Halaman ---
st.set_page_config(
    page_title="SKU Evaluator Pro",
    page_icon="📊",
    layout="wide"
)

# --- Custom CSS ---
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 900;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.5rem;
        text-align: center;
    }
    .sub-header {
        text-align: center;
        color: #666;
        margin-bottom: 2rem;
        font-size: 0.9rem;
    }
    .sku-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 16px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        color: white;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
    }
    .sku-title { font-size: 1.6rem; font-weight: 800; margin-bottom: 0.5rem; }
    .sku-badges { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 8px; }
    .badge {
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        background: rgba(255,255,255,0.2);
        backdrop-filter: blur(4px);
    }
    .sku-stats { text-align: right; }
    .stat-label { font-size: 0.7rem; opacity: 0.8; text-transform: uppercase; }
    .stat-value { font-size: 1.4rem; font-weight: 700; }
    
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border-top: 3px solid;
        text-align: center;
        transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-3px); box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .metric-value { font-size: 1.8rem; font-weight: 800; }
    .metric-label { font-size: 0.75rem; color: #666; text-transform: uppercase; letter-spacing: 1px; }
    
    .filter-panel {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 1rem;
        margin-bottom: 1.5rem;
        display: flex;
        gap: 1rem;
        flex-wrap: wrap;
        align-items: flex-end;
    }
    .filter-item { flex: 1; min-width: 150px; }
    .filter-label { font-size: 0.7rem; font-weight: 600; color: #666; margin-bottom: 4px; text-transform: uppercase; }
    
    .diagnostic-box {
        background: #F8FAFC;
        border-radius: 10px;
        padding: 1rem;
        margin-bottom: 0.75rem;
        border-left: 4px solid;
    }
    .diagnostic-title { font-weight: 700; display: flex; align-items: center; gap: 8px; margin-bottom: 4px; }
    .diagnostic-desc { font-size: 0.85rem; color: #4B5563; margin-left: 28px; }
    
    .status-active {
        background: #10B981;
        color: white;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.7rem;
        font-weight: 600;
        display: inline-block;
    }
    .status-inactive {
        background: #EF4444;
        color: white;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.7rem;
        font-weight: 600;
        display: inline-block;
    }
    hr { margin: 1.5rem 0; }
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown('<h1 class="main-header">📊 SKU 360° Evaluator Pro</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Analisis Performa SKU: Perbandingan Sales vs Purchase Order</p>', unsafe_allow_html=True)

# --- Koneksi Google Sheets ---
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
def load_data(_client):
    gsheet_url = "https://docs.google.com/spreadsheets/d/1REhZBDsFXLlCgKJbKKilRKrIEPcBzvA6EQRPY_XhZPg/edit?gid=2062248078#gid=2062248078"
    data = {}
    
    try:
        # 1. Product Master
        ws_prod = _client.open_by_url(gsheet_url).worksheet("Product_Master")
        df_product = pd.DataFrame(ws_prod.get_all_records())
        df_product.columns = [col.strip().replace(' ', '_') for col in df_product.columns]
        
        for col in ['Floor_Price', 'Purchase_Order_Price']:
            if col in df_product.columns:
                df_product[col] = pd.to_numeric(df_product[col], errors='coerce').fillna(0)
        
        if 'Status' in df_product.columns:
            df_product['Status'] = df_product['Status'].astype(str).str.upper()
        else:
            df_product['Status'] = 'ACTIVE'
        
        data['product'] = df_product
        
        # 2. Data PO
        ws_po = _client.open_by_url(gsheet_url).worksheet("Data_PO")
        df_po_raw = pd.DataFrame(ws_po.get_all_records())
        df_po_raw.columns = [col.strip() for col in df_po_raw.columns]
        
        month_cols_po = [c for c in df_po_raw.columns if c != 'SKU_ID' and any(m in c.upper() for m in ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'])]
        
        if month_cols_po and 'SKU_ID' in df_po_raw.columns:
            df_po_long = df_po_raw.melt(id_vars=['SKU_ID'], value_vars=month_cols_po, var_name='Month_Label', value_name='PO_Qty')
            df_po_long['PO_Qty'] = pd.to_numeric(df_po_long['PO_Qty'], errors='coerce').fillna(0)
            df_po_long['Month'] = df_po_long['Month_Label'].apply(parse_month)
            data['po'] = df_po_long
        else:
            data['po'] = pd.DataFrame()
        
        # 3. Sales 2025
        ws_sales25 = _client.open_by_url(gsheet_url).worksheet("Sales_2025")
        df_sales25_raw = pd.DataFrame(ws_sales25.get_all_records())
        df_sales25_raw.columns = [col.strip() for col in df_sales25_raw.columns]
        
        month_cols_25 = [c for c in df_sales25_raw.columns if any(m in c.upper() for m in ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'])]
        
        if month_cols_25:
            id_cols = ['OLD_Material', 'SKU Name', 'Group', 'Brand', 'SKU Tier']
            id_cols = [c for c in id_cols if c in df_sales25_raw.columns]
            df_sales25_long = df_sales25_raw.melt(id_vars=id_cols, value_vars=month_cols_25, var_name='Month_Label', value_name='Sales_Qty')
            df_sales25_long['Sales_Qty'] = pd.to_numeric(df_sales25_long['Sales_Qty'], errors='coerce').fillna(0)
            df_sales25_long['Month'] = df_sales25_long['Month_Label'].apply(parse_month)
            df_sales25_long['Year'] = 2025
            
            if 'OLD_Material' in df_sales25_long.columns:
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
        
        # Gabungkan Sales
        sales_list = []
        if not data['sales_2025'].empty:
            sales_list.append(data['sales_2025'])
        if not data['sales_2026'].empty:
            sales_list.append(data['sales_2026'])
        
        data['sales'] = pd.concat(sales_list, ignore_index=True) if sales_list else pd.DataFrame()
        
        return data
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return {}

def parse_month(month_str):
    if pd.isna(month_str):
        return datetime.now()
    
    month_str = str(month_str).strip()
    
    month_map = {
        'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
        'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
    }
    
    formats = ['%b %Y', '%b-%Y', '%b-%y', '%b %y', '%B %Y']
    for fmt in formats:
        try:
            return datetime.strptime(month_str, fmt)
        except:
            continue
    
    for month_name, month_num in month_map.items():
        if month_name in month_str.upper():
            numbers = ''.join([c for c in month_str if c.isdigit()])
            if numbers:
                year = 2000 + int(numbers) if len(numbers) == 2 else int(numbers)
            else:
                year = datetime.now().year
            return datetime(year, month_num, 1)
    
    return datetime.now()

def format_rupiah(value):
    if pd.isna(value) or value == 0:
        return "Rp 0"
    if value >= 1_000_000_000:
        return f"Rp {value/1e9:,.1f} M"
    elif value >= 1_000_000:
        return f"Rp {value/1e6:,.1f} Jt"
    else:
        return f"Rp {value:,.0f}"

def calculate_sku_metrics(df_sales, df_po, sku_id):
    metrics = {
        'sales_data': pd.DataFrame(),
        'po_data': pd.DataFrame(),
        'total_sales': 0,
        'total_po': 0,
        'avg_monthly_sales': 0,
        'avg_monthly_po': 0,
        'last_sales_month': None,
        'last_po_month': None,
        'months_with_sales': 0,
        'months_with_po': 0
    }
    
    sales_sku = df_sales[df_sales['SKU_ID'] == sku_id].copy() if not df_sales.empty else pd.DataFrame()
    po_sku = df_po[df_po['SKU_ID'] == sku_id].copy() if not df_po.empty else pd.DataFrame()
    
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
            metrics['avg_monthly_po'] = po_monthly['PO_Qty'].mean()
    
    return metrics

# --- Load Data ---
client = init_gsheet_connection()
if client is None:
    st.stop()

with st.spinner('🔄 Loading data...'):
    all_data = load_data(client)
    
    df_product = all_data.get('product', pd.DataFrame())
    df_sales = all_data.get('sales', pd.DataFrame())
    df_po = all_data.get('po', pd.DataFrame())

if df_product.empty:
    st.error("❌ Data Product Master kosong.")
    st.stop()

# --- FILTER PANEL (di halaman utama, bukan sidebar) ---
st.markdown('<div class="filter-panel">', unsafe_allow_html=True)

col_filter1, col_filter2, col_filter3, col_filter4, col_filter5 = st.columns([1.5, 1.5, 2, 1, 0.8])

with col_filter1:
    st.markdown('<div class="filter-label">🏷️ FILTER BRAND</div>', unsafe_allow_html=True)
    brands = ['Semua'] + sorted(df_product['Brand'].dropna().unique().tolist()) if 'Brand' in df_product.columns else ['Semua']
    selected_brand = st.selectbox("Brand", brands, label_visibility="collapsed")

with col_filter2:
    st.markdown('<div class="filter-label">💎 FILTER TIER</div>', unsafe_allow_html=True)
    tiers = ['Semua'] + sorted(df_product['SKU_Tier'].dropna().unique().tolist()) if 'SKU_Tier' in df_product.columns else ['Semua']
    selected_tier = st.selectbox("Tier", tiers, label_visibility="collapsed")

with col_filter3:
    st.markdown('<div class="filter-label">🔍 CARI SKU</div>', unsafe_allow_html=True)
    search_term = st.text_input("Cari SKU atau Nama Produk", placeholder="Ketik SKU ID atau nama...", label_visibility="collapsed")

with col_filter4:
    st.markdown('<div class="filter-label">📊 TIPE CHART</div>', unsafe_allow_html=True)
    chart_type = st.selectbox("Chart Type", ["Bar Chart", "Line Chart"], label_visibility="collapsed")

with col_filter5:
    st.markdown('<div class="filter-label" style="opacity:0;">Refresh</div>', unsafe_allow_html=True)
    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# --- Filter SKU ---
filtered_skus = df_product.copy()

if selected_brand != 'Semua':
    filtered_skus = filtered_skus[filtered_skus['Brand'] == selected_brand]

if selected_tier != 'Semua':
    filtered_skus = filtered_skus[filtered_skus['SKU_Tier'] == selected_tier]

if search_term:
    filtered_skus = filtered_skus[
        filtered_skus['SKU_ID'].astype(str).str.contains(search_term, case=False, na=False) |
        filtered_skus['Product_Name'].astype(str).str.contains(search_term, case=False, na=False)
    ]

# --- Dropdown SKU Utama ---
sku_options = filtered_skus.apply(
    lambda x: f"{x['SKU_ID']} - {x.get('Product_Name', '')}", axis=1
).tolist()

if not sku_options:
    st.warning("⚠️ Tidak ada SKU yang sesuai dengan filter.")
    st.stop()

selected_sku_display = st.selectbox("📦 Pilih SKU untuk Dianalisis", sku_options, index=0)
selected_sku = selected_sku_display.split(" - ")[0].strip()

# --- Ambil Data SKU dari Product Master ---
sku_master = df_product[df_product['SKU_ID'] == selected_sku]
if sku_master.empty:
    st.error(f"SKU {selected_sku} tidak ditemukan")
    st.stop()

sku_master = sku_master.iloc[0]
product_name = sku_master.get('Product_Name', 'Unknown')
brand = sku_master.get('Brand', 'Unknown')
tier = sku_master.get('SKU_Tier', 'Standard')
floor_price = sku_master.get('Floor_Price', 0)
purchase_price = sku_master.get('Purchase_Order_Price', 0)
moq = sku_master.get('MOQ', 0)
status = sku_master.get('Status', 'ACTIVE')

# --- Hitung Metrik SKU ---
metrics = calculate_sku_metrics(df_sales, df_po, selected_sku)

# =============================================================================
# HEADER SKU (dengan Status)
# =============================================================================
status_badge = '<span class="status-active">ACTIVE</span>' if status == 'ACTIVE' else '<span class="status-inactive">INACTIVE</span>'

st.markdown(f"""
<div class="sku-header">
    <div>
        <div class="sku-title">{product_name} <span style="font-size:0.9rem;">({selected_sku})</span></div>
        <div class="sku-badges">
            <span class="badge">🏷️ {brand}</span>
            <span class="badge">💎 {tier}</span>
            <span class="badge">📦 MOQ: {moq:,.0f}</span>
            <span class="badge">💰 Harga: {format_rupiah(floor_price)}</span>
            {status_badge}
        </div>
    </div>
    <div class="sku-stats">
        <div class="stat-label">STATUS</div>
        <div class="stat-value">{status}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# METRIC CARDS (4 buah: Total Sales, Total PO, Avg Monthly Sales, Avg Monthly PO)
# =============================================================================
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="metric-card" style="border-top-color: #10B981;">
        <div class="metric-value">{metrics['total_sales']:,.0f}</div>
        <div class="metric-label">📈 TOTAL SALES</div>
        <div style="font-size:0.7rem; color:#666;">{metrics['months_with_sales']} bulan aktif</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card" style="border-top-color: #F59E0B;">
        <div class="metric-value">{metrics['total_po']:,.0f}</div>
        <div class="metric-label">📦 TOTAL PO</div>
        <div style="font-size:0.7rem; color:#666;">{metrics['months_with_po']} bulan order</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card" style="border-top-color: #6366F1;">
        <div class="metric-value">{metrics['avg_monthly_sales']:.0f}</div>
        <div class="metric-label">📊 AVG MONTHLY SALES</div>
        <div style="font-size:0.7rem; color:#666;">Rata-rata per bulan</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="metric-card" style="border-top-color: #8B5CF6;">
        <div class="metric-value">{metrics['avg_monthly_po']:.0f}</div>
        <div class="metric-label">🎯 AVG MONTHLY PO</div>
        <div style="font-size:0.7rem; color:#666;">Rata-rata order per bulan</div>
    </div>
    """, unsafe_allow_html=True)

# =============================================================================
# TREND CHART (Sales vs PO) dengan pilihan Line atau Bar
# =============================================================================
st.markdown("---")
st.subheader("📈 Tren Sales vs Purchase Order")

# Gabungkan data
combined_data = []
if not metrics['sales_data'].empty:
    for _, row in metrics['sales_data'].iterrows():
        combined_data.append({
            'Month': row['Month'],
            'Month_Label': row['Month'].strftime('%b %Y'),
            'Sales': row['Sales_Qty'],
            'PO': 0,
            'Year': row['Month'].year
        })

if not metrics['po_data'].empty:
    for _, row in metrics['po_data'].iterrows():
        existing = next((x for x in combined_data if x['Month'] == row['Month']), None)
        if existing:
            existing['PO'] = row['PO_Qty']
        else:
            combined_data.append({
                'Month': row['Month'],
                'Month_Label': row['Month'].strftime('%b %Y'),
                'Sales': 0,
                'PO': row['PO_Qty'],
                'Year': row['Month'].year
            })

if combined_data:
    df_trend = pd.DataFrame(combined_data)
    df_trend = df_trend.sort_values('Month')
    
    fig = go.Figure()
    
    if chart_type == "Bar Chart":
        fig.add_trace(go.Bar(
            x=df_trend['Month_Label'],
            y=df_trend['Sales'],
            name='Sales Aktual',
            marker_color='#10B981',
            text=df_trend['Sales'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='outside'
        ))
        fig.add_trace(go.Bar(
            x=df_trend['Month_Label'],
            y=df_trend['PO'],
            name='Purchase Order',
            marker_color='#F59E0B',
            text=df_trend['PO'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='outside'
        ))
        fig.update_layout(barmode='group')
    else:  # Line Chart
        fig.add_trace(go.Scatter(
            x=df_trend['Month_Label'],
            y=df_trend['Sales'],
            name='Sales Aktual',
            mode='lines+markers',
            line=dict(color='#10B981', width=3),
            marker=dict(size=8, color='#10B981'),
            text=df_trend['Sales'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='top center'
        ))
        fig.add_trace(go.Scatter(
            x=df_trend['Month_Label'],
            y=df_trend['PO'],
            name='Purchase Order',
            mode='lines+markers',
            line=dict(color='#F59E0B', width=3, dash='dash'),
            marker=dict(size=8, color='#F59E0B', symbol='diamond'),
            text=df_trend['PO'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='top center'
        ))
    
    fig.update_layout(
        height=450,
        xaxis_title='Periode',
        yaxis_title='Quantity (Units)',
        hovermode='x unified',
        plot_bgcolor='white',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        margin=dict(t=50, b=40)
    )
    
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("📊 Tidak ada data Sales atau PO untuk SKU ini")

# =============================================================================
# SMART DIAGNOSTICS (Fokus ke Trend Sales)
# =============================================================================
st.markdown("---")
st.subheader("🩺 Smart Diagnostics & Rekomendasi")

diagnoses = []

# Analisis Trend Sales (naik/turun)
if metrics['months_with_sales'] >= 4:
    sales_df = metrics['sales_data'].sort_values('Month')
    if len(sales_df) >= 4:
        # Bandingkan 3 bulan terakhir vs 3 bulan sebelumnya
        recent_3 = sales_df.tail(3)['Sales_Qty'].mean()
        prev_3 = sales_df.iloc[-6:-3]['Sales_Qty'].mean() if len(sales_df) >= 6 else recent_3
        
        if prev_3 > 0:
            growth = (recent_3 - prev_3) / prev_3 * 100
            if growth > 30:
                diagnoses.append(("🚀", "Surging Demand (Meningkat Drastis)", 
                                 f"Sales rata-rata 3 bulan terakhir naik {growth:.1f}% dibanding periode sebelumnya. 🔥 Siapkan stok tambahan!", "#10B981"))
            elif growth > 10:
                diagnoses.append(("📈", "Positive Growth (Tren Naik)", 
                                 f"Sales meningkat {growth:.1f}% dalam 3 bulan terakhir. Pertahankan momentum!", "#3B82F6"))
            elif growth < -30:
                diagnoses.append(("📉", "Declining Demand (Tren Turun Drastis)", 
                                 f"Sales turun {abs(growth):.1f}% dalam 3 bulan terakhir. ⚠️ Evaluasi ulang strategi!", "#EF4444"))
            elif growth < -10:
                diagnoses.append(("🔻", "Negative Growth (Tren Turun)", 
                                 f"Sales turun {abs(growth):.1f}%. Perlu investigasi penyebab penurunan.", "#F59E0B"))
            else:
                diagnoses.append(("🟢", "Stable Demand (Stabil)", 
                                 f"Sales relatif stabil (perubahan {growth:+.1f}%). Lanjutkan strategi saat ini.", "#10B981"))
elif metrics['months_with_sales'] > 0:
    diagnoses.append(("ℹ️", "Data Terbatas", 
                     f"Hanya {metrics['months_with_sales']} bulan data sales. Pantau secara rutin untuk melihat trend.", "#6B7280"))
else:
    diagnoses.append(("⚠️", "No Sales Data", 
                     "SKU ini belum memiliki riwayat penjualan. Order dengan hati-hati (trial order).", "#F59E0B"))

# Analisis Gap Sales vs PO (jika ada data)
if metrics['total_po'] > 0 and metrics['total_sales'] > 0:
    sell_through = (metrics['total_sales'] / metrics['total_po'] * 100)
    if sell_through < 40:
        diagnoses.append(("📦", "Low Sell-Through (Penjualan Rendah)", 
                         f"Hanya {sell_through:.1f}% dari total PO yang terjual. Risiko dead stock tinggi!", "#EF4444"))
    elif sell_through > 100:
        diagnoses.append(("🔥", "High Demand (Melebihi PO)", 
                         f"Sales {sell_through:.0f}% melebihi PO. Potensi lost sales! Segera tambah order.", "#F59E0B"))
    elif sell_through < 80:
        diagnoses.append(("⚠️", "Moderate Sell-Through", 
                         f"{sell_through:.1f}% terjual. Masih ada stok tersisa untuk dijual.", "#F59E0B"))

# Tampilkan diagnosa
for icon, title, desc, color in diagnoses:
    bg_color = "#F0FDF4" if "🟢" in icon else "#FEF2F2" if "🔴" in icon or "📉" in icon else "#FFFBEB"
    st.markdown(f"""
    <div class="diagnostic-box" style="background:{bg_color}; border-left-color:{color};">
        <div class="diagnostic-title">
            <span style="font-size:1.2rem;">{icon}</span> {title}
        </div>
        <div class="diagnostic-desc">{desc}</div>
    </div>
    """, unsafe_allow_html=True)

# =============================================================================
# FINANCIAL SUMMARY (Total PO Value vs Total Sales Value)
# =============================================================================
st.markdown("---")
st.subheader("💰 Financial Summary")

total_po_value = metrics['total_po'] * floor_price
total_sales_value = metrics['total_sales'] * floor_price
gap_value = total_po_value - total_sales_value

col_fin1, col_fin2, col_fin3 = st.columns(3)

with col_fin1:
    st.metric("Total PO Value", format_rupiah(total_po_value), 
              help="Nilai total Purchase Order berdasarkan Floor Price")

with col_fin2:
    st.metric("Total Sales Value", format_rupiah(total_sales_value),
              help="Nilai total Penjualan berdasarkan Floor Price")

with col_fin3:
    delta_color = "normal" if gap_value >= 0 else "inverse"
    st.metric("Gap (PO - Sales)", format_rupiah(gap_value), 
              delta=f"{gap_value/total_po_value*100:.1f}% dari PO" if total_po_value > 0 else None,
              delta_color=delta_color)

# Peringatan jika harga 0
if floor_price == 0:
    st.warning("⚠️ Floor_Price = 0, nilai finansial bersifat estimasi. Periksa data Product Master.")

# =============================================================================
# DETAIL DATA PER BULAN (Tabel)
# =============================================================================
st.markdown("---")
with st.expander("📋 Lihat Detail Data per Bulan", expanded=False):
    if combined_data:
        detail_df = pd.DataFrame(combined_data).sort_values('Month')
        detail_df['Bulan'] = detail_df['Month'].dt.strftime('%b %Y')
        detail_df['Sales'] = detail_df['Sales'].apply(lambda x: f"{x:,.0f}")
        detail_df['PO'] = detail_df['PO'].apply(lambda x: f"{x:,.0f}")
        detail_df['Sales Value'] = (detail_df['Sales'].str.replace(',', '').astype(float) * floor_price).apply(format_rupiah)
        detail_df['PO Value'] = (detail_df['PO'].str.replace(',', '').astype(float) * floor_price).apply(format_rupiah)
        
        display_cols = ['Bulan', 'Sales', 'Sales Value', 'PO', 'PO Value']
        st.dataframe(detail_df[display_cols], use_container_width=True, hide_index=True)
    else:
        st.info("Tidak ada data")

# =============================================================================
# FOOTER
# =============================================================================
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #888; font-size: 0.75rem; padding: 1rem;">
    <p>📊 SKU 360° Evaluator Pro | Data Sales 2025-2026 | Data PO hingga Mar 2026</p>
</div>
""", unsafe_allow_html=True)
