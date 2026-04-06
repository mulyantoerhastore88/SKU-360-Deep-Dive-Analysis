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

# --- Custom CSS (Compact Version) ---
st.markdown("""
<style>
    .main-header {
        font-size: 1.8rem;
        font-weight: 900;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.25rem;
        text-align: center;
    }
    .sub-header {
        text-align: center;
        color: #666;
        margin-bottom: 1rem;
        font-size: 0.8rem;
    }
    .sku-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 0.75rem 1.25rem;
        margin-bottom: 1rem;
        color: white;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
    }
    .sku-title { font-size: 1.2rem; font-weight: 700; margin-bottom: 0.25rem; }
    .sku-badges { display: flex; gap: 6px; flex-wrap: wrap; margin-top: 4px; }
    .badge {
        padding: 2px 8px;
        border-radius: 16px;
        font-size: 0.65rem;
        font-weight: 600;
        background: rgba(255,255,255,0.2);
        backdrop-filter: blur(4px);
    }
    .sku-stats { text-align: right; }
    .stat-label { font-size: 0.6rem; opacity: 0.8; text-transform: uppercase; }
    .stat-value { font-size: 1rem; font-weight: 700; }
    
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 0.6rem 0.8rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        border-top: 2px solid;
        text-align: center;
    }
    .metric-value { font-size: 1.3rem; font-weight: 800; line-height: 1.2; }
    .metric-label { font-size: 0.65rem; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 4px; }
    .metric-sub { font-size: 0.6rem; color: #999; margin-top: 2px; }
    
    .control-panel {
        background: #f8f9fa;
        border-radius: 10px;
        padding: 0.75rem 1rem;
        margin-bottom: 1rem;
        display: flex;
        gap: 1rem;
        flex-wrap: wrap;
        align-items: flex-end;
    }
    .control-item { flex: 1; min-width: 200px; }
    .control-label { font-size: 0.65rem; font-weight: 600; color: #666; margin-bottom: 2px; text-transform: uppercase; }
    
    .diagnostic-box {
        background: #F8FAFC;
        border-radius: 8px;
        padding: 0.6rem 1rem;
        margin-bottom: 0.5rem;
        border-left: 3px solid;
    }
    .diagnostic-title { font-weight: 700; font-size: 0.85rem; display: flex; align-items: center; gap: 6px; }
    .diagnostic-desc { font-size: 0.75rem; color: #4B5563; margin-left: 24px; }
    
    .status-active {
        background: #10B981;
        color: white;
        padding: 2px 8px;
        border-radius: 16px;
        font-size: 0.65rem;
        font-weight: 600;
        display: inline-block;
    }
    .status-inactive {
        background: #EF4444;
        color: white;
        padding: 2px 8px;
        border-radius: 16px;
        font-size: 0.65rem;
        font-weight: 600;
        display: inline-block;
    }
    .status-notfound {
        background: #9CA3AF;
        color: white;
        padding: 2px 8px;
        border-radius: 16px;
        font-size: 0.65rem;
        font-weight: 600;
        display: inline-block;
    }
    hr { margin: 1rem 0; }
    .small-text { font-size: 0.7rem; color: #666; }
    .compare-badge {
        background: #8B5CF6;
        color: white;
        padding: 2px 8px;
        border-radius: 16px;
        font-size: 0.7rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown('<h1 class="main-header">📊 SKU 360° Evaluator Pro</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Analisis Performa SKU: Perbandingan Sales vs Purchase Order | Support Perbandingan 2 SKU</p>', unsafe_allow_html=True)

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
            
            # Map OLD_Material ke SKU_ID
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

def calculate_sku_metrics(df_sales, df_po, sku_id, df_product):
    """Hitung metrik untuk SKU tertentu"""
    metrics = {
        'sales_data': pd.DataFrame(),
        'po_data': pd.DataFrame(),
        'total_sales': 0,
        'total_po': 0,
        'avg_monthly_sales': 0,
        'avg_monthly_po': 0,
        'months_with_sales': 0,
        'months_with_po': 0,
        'floor_price': 0,
        'purchase_price': 0,
        'status': 'NOT_FOUND',
        'product_name': sku_id,
        'brand': '-',
        'tier': '-',
        'moq': 0
    }
    
    # Ambil data dari Product Master jika ada
    product_data = df_product[df_product['SKU_ID'] == sku_id]
    if not product_data.empty:
        metrics['floor_price'] = product_data.iloc[0].get('Floor_Price', 0)
        metrics['purchase_price'] = product_data.iloc[0].get('Purchase_Order_Price', 0)
        metrics['status'] = product_data.iloc[0].get('Status', 'NOT_FOUND')
        metrics['product_name'] = product_data.iloc[0].get('Product_Name', sku_id)
        metrics['brand'] = product_data.iloc[0].get('Brand', '-')
        metrics['tier'] = product_data.iloc[0].get('SKU_Tier', '-')
        metrics['moq'] = product_data.iloc[0].get('MOQ', 0)
    
    # Sales data
    sales_sku = df_sales[df_sales['SKU_ID'] == sku_id].copy() if not df_sales.empty else pd.DataFrame()
    if not sales_sku.empty:
        sales_monthly = sales_sku.groupby('Month')['Sales_Qty'].sum().reset_index()
        sales_monthly = sales_monthly.sort_values('Month')
        metrics['sales_data'] = sales_monthly
        metrics['total_sales'] = sales_monthly['Sales_Qty'].sum()
        metrics['months_with_sales'] = len(sales_monthly)
        if not sales_monthly.empty:
            metrics['avg_monthly_sales'] = sales_monthly['Sales_Qty'].mean()
    
    # PO data
    po_sku = df_po[df_po['SKU_ID'] == sku_id].copy() if not df_po.empty else pd.DataFrame()
    if not po_sku.empty:
        po_monthly = po_sku.groupby('Month')['PO_Qty'].sum().reset_index()
        po_monthly = po_monthly.sort_values('Month')
        metrics['po_data'] = po_monthly
        metrics['total_po'] = po_monthly['PO_Qty'].sum()
        metrics['months_with_po'] = len(po_monthly)
        if not po_monthly.empty:
            metrics['avg_monthly_po'] = po_monthly['PO_Qty'].mean()
    
    return metrics

def get_all_skus_from_data(df_sales, df_po, df_product):
    """Dapatkan semua SKU unik dari Sales, PO, dan Product Master"""
    sku_set = set()
    
    if not df_sales.empty:
        sku_set.update(df_sales['SKU_ID'].dropna().unique())
    
    if not df_po.empty:
        sku_set.update(df_po['SKU_ID'].dropna().unique())
    
    if not df_product.empty:
        sku_set.update(df_product['SKU_ID'].dropna().unique())
    
    # Filter out NaN
    sku_set = {s for s in sku_set if pd.notna(s) and str(s).strip() != ''}
    
    return sorted(list(sku_set))

# --- Load Data ---
client = init_gsheet_connection()
if client is None:
    st.stop()

with st.spinner('🔄 Loading data...'):
    all_data = load_data(client)
    
    df_product = all_data.get('product', pd.DataFrame())
    df_sales = all_data.get('sales', pd.DataFrame())
    df_po = all_data.get('po', pd.DataFrame())

# --- Dapatkan semua SKU unik ---
all_skus = get_all_skus_from_data(df_sales, df_po, df_product)

if not all_skus:
    st.error("❌ Tidak ada data SKU ditemukan.")
    st.stop()

# --- Control Panel ---
st.markdown('<div class="control-panel">', unsafe_allow_html=True)

col_sku1, col_sku2, col_chart, col_refresh = st.columns([2, 2, 1.5, 0.8])

with col_sku1:
    st.markdown('<div class="control-label">📦 SKU UTAMA (Referensi)</div>', unsafe_allow_html=True)
    # Buat mapping display name
    sku_display_map = {}
    sku_display_list = []
    for sku in all_skus:
        # Cari nama produk dari Product Master
        product_row = df_product[df_product['SKU_ID'] == sku]
        if not product_row.empty:
            product_name = product_row.iloc[0].get('Product_Name', '')
            display = f"{sku} - {product_name}" if product_name else sku
        else:
            display = f"{sku} (No Product Data)"
        sku_display_map[display] = sku
        sku_display_list.append(display)
    
    selected_main_display = st.selectbox("SKU Utama", sku_display_list, key="main_sku")

with col_sku2:
    st.markdown('<div class="control-label">🔄 SKU PEMBANDING (Opsional)</div>', unsafe_allow_html=True)
    compare_options = ["[Tidak ada perbandingan]"] + sku_display_list
    selected_compare_display = st.selectbox("SKU Pembanding", compare_options, key="compare_sku")

with col_chart:
    st.markdown('<div class="control-label">📊 TIPE CHART</div>', unsafe_allow_html=True)
    chart_type = st.selectbox("Chart Type", ["Bar Chart", "Line Chart"], label_visibility="collapsed")

with col_refresh:
    st.markdown('<div class="control-label" style="opacity:0;">Refresh</div>', unsafe_allow_html=True)
    if st.button("🔄 Refresh", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# --- Ambil SKU ID ---
main_sku = sku_display_map[selected_main_display]
compare_sku = sku_display_map[selected_compare_display] if selected_compare_display != "[Tidak ada perbandingan]" else None

# --- Hitung Metrik untuk SKU Utama ---
main_metrics = calculate_sku_metrics(df_sales, df_po, main_sku, df_product)

# =============================================================================
# HEADER SKU UTAMA (Compact)
# =============================================================================
status_class = "status-active" if main_metrics['status'] == 'ACTIVE' else "status-inactive" if main_metrics['status'] == 'INACTIVE' else "status-notfound"
status_text = main_metrics['status'] if main_metrics['status'] != 'NOT_FOUND' else "TIDAK ADA DI PRODUCT MASTER"

st.markdown(f"""
<div class="sku-header">
    <div>
        <div class="sku-title">{main_metrics['product_name']} <span style="font-size:0.8rem;">({main_sku})</span></div>
        <div class="sku-badges">
            <span class="badge">🏷️ {main_metrics['brand']}</span>
            <span class="badge">💎 {main_metrics['tier']}</span>
            <span class="badge">📦 MOQ: {main_metrics['moq']:,.0f}</span>
            <span class="{status_class}">{status_text}</span>
        </div>
    </div>
    <div class="sku-stats">
        <div class="stat-label">FLOOR PRICE</div>
        <div class="stat-value">{format_rupiah(main_metrics['floor_price'])}</div>
        <div class="stat-label" style="margin-top:4px;">PURCHASE PRICE</div>
        <div class="stat-value">{format_rupiah(main_metrics['purchase_price'])}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# METRIC CARDS (Compact - 4 kolom)
# =============================================================================
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class="metric-card" style="border-top-color: #10B981;">
        <div class="metric-value">{main_metrics['total_sales']:,.0f}</div>
        <div class="metric-label">📈 TOTAL SALES</div>
        <div class="metric-sub">{main_metrics['months_with_sales']} bulan aktif</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="metric-card" style="border-top-color: #F59E0B;">
        <div class="metric-value">{main_metrics['total_po']:,.0f}</div>
        <div class="metric-label">📦 TOTAL PO</div>
        <div class="metric-sub">{main_metrics['months_with_po']} bulan order</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card" style="border-top-color: #6366F1;">
        <div class="metric-value">{main_metrics['avg_monthly_sales']:.0f}</div>
        <div class="metric-label">📊 AVG MONTHLY SALES</div>
        <div class="metric-sub">Rata-rata per bulan</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="metric-card" style="border-top-color: #8B5CF6;">
        <div class="metric-value">{main_metrics['avg_monthly_po']:.0f}</div>
        <div class="metric-label">🎯 AVG MONTHLY PO</div>
        <div class="metric-sub">Rata-rata order per bulan</div>
    </div>
    """, unsafe_allow_html=True)

# =============================================================================
# TREND CHART (Support 2 SKU)
# =============================================================================
st.markdown("---")
st.subheader("📈 Tren Sales vs Purchase Order")

# Fungsi untuk menyiapkan data SKU
def prepare_chart_data(sku_id, metrics, label, color_sales, color_po):
    combined = []
    if not metrics['sales_data'].empty:
        for _, row in metrics['sales_data'].iterrows():
            combined.append({
                'Month': row['Month'],
                'Month_Label': row['Month'].strftime('%b %Y'),
                f'{label}_Sales': row['Sales_Qty'],
                f'{label}_PO': 0
            })
    
    if not metrics['po_data'].empty:
        for _, row in metrics['po_data'].iterrows():
            existing = next((x for x in combined if x['Month'] == row['Month']), None)
            if existing:
                existing[f'{label}_PO'] = row['PO_Qty']
            else:
                combined.append({
                    'Month': row['Month'],
                    'Month_Label': row['Month'].strftime('%b %Y'),
                    f'{label}_Sales': 0,
                    f'{label}_PO': row['PO_Qty']
                })
    
    df = pd.DataFrame(combined) if combined else pd.DataFrame()
    return df.sort_values('Month') if not df.empty else df

# Prepare data
main_df = prepare_chart_data(main_sku, main_metrics, 'Main', '#10B981', '#F59E0B')
compare_df = None
if compare_sku:
    compare_metrics = calculate_sku_metrics(df_sales, df_po, compare_sku, df_product)
    compare_df = prepare_chart_data(compare_sku, compare_metrics, 'Compare', '#8B5CF6', '#EC4899')

# Buat chart
if not main_df.empty:
    fig = go.Figure()
    
    # SKU Utama - Sales
    if chart_type == "Bar Chart":
        fig.add_trace(go.Bar(
            x=main_df['Month_Label'], y=main_df['Main_Sales'],
            name=f'{main_sku} - Sales',
            marker_color='#10B981',
            text=main_df['Main_Sales'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='outside'
        ))
        fig.add_trace(go.Bar(
            x=main_df['Month_Label'], y=main_df['Main_PO'],
            name=f'{main_sku} - PO',
            marker_color='#F59E0B',
            text=main_df['Main_PO'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='outside'
        ))
    else:
        fig.add_trace(go.Scatter(
            x=main_df['Month_Label'], y=main_df['Main_Sales'],
            name=f'{main_sku} - Sales',
            mode='lines+markers',
            line=dict(color='#10B981', width=2.5),
            marker=dict(size=6, color='#10B981'),
            text=main_df['Main_Sales'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='top center'
        ))
        fig.add_trace(go.Scatter(
            x=main_df['Month_Label'], y=main_df['Main_PO'],
            name=f'{main_sku} - PO',
            mode='lines+markers',
            line=dict(color='#F59E0B', width=2.5, dash='dash'),
            marker=dict(size=6, color='#F59E0B', symbol='diamond'),
            text=main_df['Main_PO'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='top center'
        ))
    
    # SKU Pembanding (jika ada)
    if compare_df is not None and not compare_df.empty:
        if chart_type == "Bar Chart":
            fig.add_trace(go.Bar(
                x=compare_df['Month_Label'], y=compare_df['Compare_Sales'],
                name=f'{compare_sku} - Sales',
                marker_color='#8B5CF6',
                text=compare_df['Compare_Sales'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
                textposition='outside'
            ))
            fig.add_trace(go.Bar(
                x=compare_df['Month_Label'], y=compare_df['Compare_PO'],
                name=f'{compare_sku} - PO',
                marker_color='#EC4899',
                text=compare_df['Compare_PO'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
                textposition='outside'
            ))
            fig.update_layout(barmode='group')
        else:
            fig.add_trace(go.Scatter(
                x=compare_df['Month_Label'], y=compare_df['Compare_Sales'],
                name=f'{compare_sku} - Sales',
                mode='lines+markers',
                line=dict(color='#8B5CF6', width=2.5),
                marker=dict(size=6, color='#8B5CF6'),
                text=compare_df['Compare_Sales'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
                textposition='top center'
            ))
            fig.add_trace(go.Scatter(
                x=compare_df['Month_Label'], y=compare_df['Compare_PO'],
                name=f'{compare_sku} - PO',
                mode='lines+markers',
                line=dict(color='#EC4899', width=2.5, dash='dash'),
                marker=dict(size=6, color='#EC4899', symbol='diamond'),
                text=compare_df['Compare_PO'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
                textposition='top center'
            ))
    
    fig.update_layout(
        height=400,
        xaxis_title='Periode',
        yaxis_title='Quantity (Units)',
        hovermode='x unified',
        plot_bgcolor='white',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        margin=dict(t=40, b=40, l=20, r=20)
    )
    
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("📊 Tidak ada data Sales atau PO untuk SKU ini")

# =============================================================================
# SMART DIAGNOSTICS (Fokus ke Trend Sales)
# =============================================================================
st.markdown("---")
st.subheader("🩺 Smart Diagnostics & Rekomendasi")

def get_diagnostics(metrics, sku_label=""):
    diagnoses = []
    prefix = f"**{sku_label}** " if sku_label else ""
    
    if metrics['months_with_sales'] >= 4:
        sales_df = metrics['sales_data'].sort_values('Month')
        if len(sales_df) >= 4:
            recent_3 = sales_df.tail(3)['Sales_Qty'].mean()
            prev_3 = sales_df.iloc[-6:-3]['Sales_Qty'].mean() if len(sales_df) >= 6 else recent_3
            
            if prev_3 > 0:
                growth = (recent_3 - prev_3) / prev_3 * 100
                if growth > 30:
                    diagnoses.append(("🚀", f"{prefix}Surging Demand", f"Sales naik {growth:.1f}% dalam 3 bulan terakhir. Siapkan stok!", "#10B981"))
                elif growth > 10:
                    diagnoses.append(("📈", f"{prefix}Positive Growth", f"Sales meningkat {growth:.1f}%. Pertahankan momentum!", "#3B82F6"))
                elif growth < -30:
                    diagnoses.append(("📉", f"{prefix}Declining Demand", f"Sales turun {abs(growth):.1f}%. Evaluasi strategi!", "#EF4444"))
                elif growth < -10:
                    diagnoses.append(("🔻", f"{prefix}Negative Growth", f"Sales turun {abs(growth):.1f}%. Perlu investigasi.", "#F59E0B"))
                else:
                    diagnoses.append(("🟢", f"{prefix}Stable Demand", f"Sales stabil (perubahan {growth:+.1f}%).", "#10B981"))
    elif metrics['months_with_sales'] > 0:
        diagnoses.append(("ℹ️", f"{prefix}Limited Data", f"Hanya {metrics['months_with_sales']} bulan data. Pantau rutin.", "#6B7280"))
    else:
        diagnoses.append(("⚠️", f"{prefix}No Sales Data", "Belum ada riwayat penjualan. Order trial.", "#F59E0B"))
    
    if metrics['total_po'] > 0 and metrics['total_sales'] > 0:
        sell_through = (metrics['total_sales'] / metrics['total_po'] * 100)
        if sell_through < 40:
            diagnoses.append(("📦", f"{prefix}Low Sell-Through", f"Hanya {sell_through:.1f}% PO terjual. Risiko dead stock!", "#EF4444"))
        elif sell_through > 100:
            diagnoses.append(("🔥", f"{prefix}High Demand", f"Sales {sell_through:.0f}% > PO. Potensi lost sales!", "#F59E0B"))
    
    return diagnoses

# Tampilkan diagnosa SKU Utama
diagnoses = get_diagnostics(main_metrics, "")
for icon, title, desc, color in diagnoses:
    bg_color = "#F0FDF4" if "🟢" in icon else "#FEF2F2" if "🔴" in icon or "📉" in icon else "#FFFBEB"
    st.markdown(f"""
    <div class="diagnostic-box" style="background:{bg_color}; border-left-color:{color};">
        <div class="diagnostic-title">
            <span style="font-size:1rem;">{icon}</span> {title}
        </div>
        <div class="diagnostic-desc">{desc}</div>
    </div>
    """, unsafe_allow_html=True)

# Jika ada SKU pembanding, tampilkan juga
if compare_sku:
    st.markdown(f'<div class="small-text" style="margin-top:8px;">🔍 <strong>Perbandingan dengan {compare_sku}</strong></div>', unsafe_allow_html=True)
    compare_metrics = calculate_sku_metrics(df_sales, df_po, compare_sku, df_product)
    compare_diagnoses = get_diagnostics(compare_metrics, f"{compare_sku}")
    for icon, title, desc, color in compare_diagnoses:
        bg_color = "#F0FDF4" if "🟢" in icon else "#FEF2F2" if "🔴" in icon or "📉" in icon else "#FFFBEB"
        st.markdown(f"""
        <div class="diagnostic-box" style="background:{bg_color}; border-left-color:{color}; padding:0.4rem 1rem;">
            <div class="diagnostic-title" style="font-size:0.8rem;">
                <span style="font-size:0.9rem;">{icon}</span> {title}
            </div>
            <div class="diagnostic-desc" style="font-size:0.7rem;">{desc}</div>
        </div>
        """, unsafe_allow_html=True)

# =============================================================================
# FINANCIAL SUMMARY (PO Value = PO × Purchase_Order_Price, Sales Value = Sales × Floor_Price)
# =============================================================================
st.markdown("---")
st.subheader("💰 Financial Summary")

# SKU Utama
po_value_main = main_metrics['total_po'] * main_metrics['purchase_price']
sales_value_main = main_metrics['total_sales'] * main_metrics['floor_price']
gap_main = po_value_main - sales_value_main

col_fin1, col_fin2, col_fin3 = st.columns(3)

with col_fin1:
    st.metric(f"📦 {main_sku} - Total PO Value", format_rupiah(po_value_main),
              help=f"PO Qty × Purchase Price ({format_rupiah(main_metrics['purchase_price'])}/unit)")

with col_fin2:
    st.metric(f"💰 {main_sku} - Total Sales Value", format_rupiah(sales_value_main),
              help=f"Sales Qty × Floor Price ({format_rupiah(main_metrics['floor_price'])}/unit)")

with col_fin3:
    delta_color = "normal" if gap_main >= 0 else "inverse"
    st.metric(f"⚖️ {main_sku} - Gap", format_rupiah(gap_main),
              delta=f"{gap_main/po_value_main*100:.1f}% dari PO" if po_value_main > 0 else None,
              delta_color=delta_color)

# Jika ada SKU pembanding, tampilkan ringkasan perbandingan
if compare_sku:
    compare_metrics = calculate_sku_metrics(df_sales, df_po, compare_sku, df_product)
    po_value_comp = compare_metrics['total_po'] * compare_metrics['purchase_price']
    sales_value_comp = compare_metrics['total_sales'] * compare_metrics['floor_price']
    
    st.markdown(f'<div class="small-text" style="margin-top:8px;">📊 <strong>Perbandingan dengan {compare_sku}</strong></div>', unsafe_allow_html=True)
    
    col_c1, col_c2, col_c3, col_c4 = st.columns(4)
    with col_c1:
        st.metric("Total PO", f"{compare_metrics['total_po']:,.0f}", 
                  delta=f"{compare_metrics['total_po'] - main_metrics['total_po']:+,.0f}")
    with col_c2:
        st.metric("Total Sales", f"{compare_metrics['total_sales']:,.0f}",
                  delta=f"{compare_metrics['total_sales'] - main_metrics['total_sales']:+,.0f}")
    with col_c3:
        st.metric("PO Value", format_rupiah(po_value_comp),
                  delta=f"{format_rupiah(po_value_comp - po_value_main)}")
    with col_c4:
        st.metric("Sales Value", format_rupiah(sales_value_comp),
                  delta=f"{format_rupiah(sales_value_comp - sales_value_main)}")

# Peringatan harga
if main_metrics['floor_price'] == 0 or main_metrics['purchase_price'] == 0:
    st.warning("⚠️ Harga (Floor_Price atau Purchase_Order_Price) = 0. Periksa data Product Master.")

# =============================================================================
# DETAIL DATA PER BULAN
# =============================================================================
st.markdown("---")
with st.expander("📋 Lihat Detail Data per Bulan", expanded=False):
    if not main_df.empty:
        detail_df = main_df.copy()
        detail_df['Bulan'] = detail_df['Month'].dt.strftime('%b %Y')
        detail_df['Sales Qty'] = detail_df['Main_Sales'].apply(lambda x: f"{x:,.0f}")
        detail_df['PO Qty'] = detail_df['Main_PO'].apply(lambda x: f"{x:,.0f}")
        detail_df['Sales Value'] = (detail_df['Main_Sales'] * main_metrics['floor_price']).apply(format_rupiah)
        detail_df['PO Value'] = (detail_df['Main_PO'] * main_metrics['purchase_price']).apply(format_rupiah)
        
        display_cols = ['Bulan', 'Sales Qty', 'Sales Value', 'PO Qty', 'PO Value']
        st.dataframe(detail_df[display_cols], use_container_width=True, hide_index=True)
        
        # Jika ada SKU pembanding, tampilkan juga datanya
        if compare_sku and compare_df is not None and not compare_df.empty:
            st.markdown(f'<div class="small-text" style="margin-top:12px;">📊 <strong>Data {compare_sku}</strong></div>', unsafe_allow_html=True)
            compare_detail = compare_df.copy()
            compare_detail['Bulan'] = compare_detail['Month'].dt.strftime('%b %Y')
            compare_detail['Sales Qty'] = compare_detail['Compare_Sales'].apply(lambda x: f"{x:,.0f}")
            compare_detail['PO Qty'] = compare_detail['Compare_PO'].apply(lambda x: f"{x:,.0f}")
            st.dataframe(compare_detail[['Bulan', 'Sales Qty', 'PO Qty']], use_container_width=True, hide_index=True)
    else:
        st.info("Tidak ada data")

# =============================================================================
# FOOTER
# =============================================================================
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #888; font-size: 0.7rem; padding: 0.5rem;">
    <p>📊 SKU 360° Evaluator Pro | Data Sales 2025-2026 | Data PO hingga Mar 2026 | Support Perbandingan 2 SKU</p>
</div>
""", unsafe_allow_html=True)
