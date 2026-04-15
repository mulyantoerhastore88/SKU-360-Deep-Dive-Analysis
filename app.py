"""
SKU 360° Evaluator Pro with ECharts
Powerful SKU Performance Dashboard with Advanced Visualizations
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
from streamlit_echarts import st_echarts
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
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
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
    
    .insight-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border-radius: 10px;
        padding: 1rem;
        color: white;
        margin-bottom: 1rem;
    }
    .insight-title { font-size: 0.7rem; text-transform: uppercase; opacity: 0.7; letter-spacing: 1px; }
    .insight-value { font-size: 1.4rem; font-weight: 700; margin: 5px 0; }
    .insight-desc { font-size: 0.7rem; opacity: 0.8; }
    
    hr { margin: 1rem 0; }
    .small-text { font-size: 0.7rem; color: #666; }
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown('<h1 class="main-header">📊 SKU 360° Evaluator Pro</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Analisis Performa SKU | Perbandingan Sales vs PO vs Inbound | Deep Dive Analytics | Stock Management</p>', unsafe_allow_html=True)

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
        
        # 2. Data PO (Purchase Order)
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
        
        # 3. Data PO Delivered (Inbound)
        try:
            ws_po_delivered = _client.open_by_url(gsheet_url).worksheet("Data_PO_Delivered(Inbound)")
            df_po_delivered_raw = pd.DataFrame(ws_po_delivered.get_all_records())
            df_po_delivered_raw.columns = [col.strip() for col in df_po_delivered_raw.columns]
            
            month_cols_po_del = [c for c in df_po_delivered_raw.columns if c != 'SKU_ID' and any(m in c.upper() for m in ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'])]
            
            if month_cols_po_del and 'SKU_ID' in df_po_delivered_raw.columns:
                df_po_delivered_long = df_po_delivered_raw.melt(id_vars=['SKU_ID'], value_vars=month_cols_po_del, var_name='Month_Label', value_name='PO_Delivered_Qty')
                df_po_delivered_long['PO_Delivered_Qty'] = pd.to_numeric(df_po_delivered_long['PO_Delivered_Qty'], errors='coerce').fillna(0)
                df_po_delivered_long['Month'] = df_po_delivered_long['Month_Label'].apply(parse_month)
                data['po_delivered'] = df_po_delivered_long
            else:
                data['po_delivered'] = pd.DataFrame()
        except:
            data['po_delivered'] = pd.DataFrame()
        
        # 4. Sales 2025
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
            
            if 'OLD_Material' in df_sales25_long.columns and 'OLD_Material' in df_product.columns:
                sku_mapping = df_product[['OLD_Material', 'SKU_ID']].drop_duplicates()
                df_sales25_long = pd.merge(df_sales25_long, sku_mapping, on='OLD_Material', how='left')
            else:
                df_sales25_long['SKU_ID'] = df_sales25_long['OLD_Material']
            
            data['sales_2025'] = df_sales25_long
        
        # 5. Sales 2026
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
            
            if 'OLD_Material' in df_sales26_long.columns and 'OLD_Material' in df_product.columns:
                sku_mapping = df_product[['OLD_Material', 'SKU_ID']].drop_duplicates()
                df_sales26_long = pd.merge(df_sales26_long, sku_mapping, on='OLD_Material', how='left')
            else:
                df_sales26_long['SKU_ID'] = df_sales26_long['OLD_Material']
            
            data['sales_2026'] = df_sales26_long
        
        # Gabungkan Sales
        sales_list = []
        if 'sales_2025' in data:
            sales_list.append(data['sales_2025'])
        if 'sales_2026' in data:
            sales_list.append(data['sales_2026'])
        
        data['sales'] = pd.concat(sales_list, ignore_index=True) if sales_list else pd.DataFrame()
        
        # 6. Stock Onhand
        try:
            ws_stock = _client.open_by_url(gsheet_url).worksheet("Stock_Onhand")
            df_stock_raw = pd.DataFrame(ws_stock.get_all_records())
            df_stock_raw.columns = [col.strip().replace(' ', '_') for col in df_stock_raw.columns]
            
            stock_cols = {}
            if 'OLD_Material' in df_stock_raw.columns:
                stock_cols['OLD_Material'] = 'OLD_Material'
            if 'Batch_Number' in df_stock_raw.columns:
                stock_cols['Batch_Number'] = 'Batch_Number'
            if 'Product_Description' in df_stock_raw.columns:
                stock_cols['Product_Description'] = 'Product_Description'
            if 'Expiry_Date' in df_stock_raw.columns:
                stock_cols['Expiry_Date'] = 'Expiry_Date'
            if 'Physical_Stock' in df_stock_raw.columns:
                stock_cols['Physical_Stock'] = 'Physical_Stock'
            
            if stock_cols:
                df_stock = df_stock_raw[list(stock_cols.keys())].copy()
                df_stock = df_stock.rename(columns=stock_cols)
                df_stock['Physical_Stock'] = pd.to_numeric(df_stock['Physical_Stock'], errors='coerce').fillna(0)
                df_stock['Expiry_Date'] = pd.to_datetime(df_stock['Expiry_Date'], errors='coerce', dayfirst=True)
                
                if 'OLD_Material' in df_stock.columns and 'OLD_Material' in df_product.columns:
                    sku_mapping = df_product[['OLD_Material', 'SKU_ID']].drop_duplicates()
                    df_stock = pd.merge(df_stock, sku_mapping, on='OLD_Material', how='left')
                else:
                    df_stock['SKU_ID'] = df_stock['OLD_Material']
                
                df_stock = df_stock[df_stock['Physical_Stock'] > 0]
                data['stock'] = df_stock
        except:
            data['stock'] = pd.DataFrame()
        
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

def calculate_sku_metrics(df_sales, df_po, df_po_delivered, sku_id, df_product):
    metrics = {
        'sales_data': pd.DataFrame(),
        'po_data': pd.DataFrame(),
        'po_delivered_data': pd.DataFrame(),
        'total_sales': 0,
        'total_po': 0,
        'total_po_delivered': 0,
        'avg_monthly_sales': 0,
        'floor_price': 0,
        'purchase_price': 0,
        'status': 'NOT_FOUND',
        'product_name': sku_id,
        'brand': '-',
        'tier': '-',
        'moq': 0,
    }
    
    product_data = df_product[df_product['SKU_ID'] == sku_id]
    if not product_data.empty:
        metrics['floor_price'] = product_data.iloc[0].get('Floor_Price', 0)
        metrics['purchase_price'] = product_data.iloc[0].get('Purchase_Order_Price', 0)
        metrics['status'] = product_data.iloc[0].get('Status', 'NOT_FOUND')
        metrics['product_name'] = product_data.iloc[0].get('Product_Name', sku_id)
        metrics['brand'] = product_data.iloc[0].get('Brand', '-')
        metrics['tier'] = product_data.iloc[0].get('SKU_Tier', '-')
        metrics['moq'] = product_data.iloc[0].get('MOQ', 0)
    
    sales_sku = df_sales[df_sales['SKU_ID'] == sku_id].copy() if not df_sales.empty else pd.DataFrame()
    if not sales_sku.empty:
        sales_monthly = sales_sku.groupby('Month')['Sales_Qty'].sum().reset_index().sort_values('Month')
        metrics['sales_data'] = sales_monthly
        metrics['total_sales'] = sales_monthly['Sales_Qty'].sum()
        metrics['avg_monthly_sales'] = sales_monthly['Sales_Qty'].mean()
    
    po_sku = df_po[df_po['SKU_ID'] == sku_id].copy() if not df_po.empty else pd.DataFrame()
    if not po_sku.empty:
        po_monthly = po_sku.groupby('Month')['PO_Qty'].sum().reset_index().sort_values('Month')
        metrics['po_data'] = po_monthly
        metrics['total_po'] = po_monthly['PO_Qty'].sum()
    
    po_delivered_sku = df_po_delivered[df_po_delivered['SKU_ID'] == sku_id].copy() if not df_po_delivered.empty else pd.DataFrame()
    if not po_delivered_sku.empty:
        po_delivered_monthly = po_delivered_sku.groupby('Month')['PO_Delivered_Qty'].sum().reset_index().sort_values('Month')
        metrics['po_delivered_data'] = po_delivered_monthly
        metrics['total_po_delivered'] = po_delivered_monthly['PO_Delivered_Qty'].sum()
    
    return metrics

def calculate_stock_metrics(df_stock, df_product, sku_id):
    stock_metrics = {
        'total_stock': 0,
        'batch_count': 0,
        'expired_stock': 0,
        'expiring_soon': 0,
        'expiring_3months': 0,
        'expiring_6months': 0,
        'fresh_stock': 0,
        'batch_details': pd.DataFrame(),
        'has_stock': False
    }
    
    if df_stock.empty:
        return stock_metrics
    
    sku_stock = df_stock[df_stock['SKU_ID'] == sku_id].copy()
    if sku_stock.empty:
        return stock_metrics
    
    stock_metrics['has_stock'] = True
    stock_metrics['total_stock'] = sku_stock['Physical_Stock'].sum()
    stock_metrics['batch_count'] = len(sku_stock)
    
    today = datetime.now().date()
    
    for _, row in sku_stock.iterrows():
        expiry = row['Expiry_Date']
        qty = row['Physical_Stock']
        
        if pd.isna(expiry):
            stock_metrics['fresh_stock'] += qty
        else:
            expiry_date = expiry.date() if hasattr(expiry, 'date') else expiry
            days_to_expiry = (expiry_date - today).days
            
            if days_to_expiry < 0:
                stock_metrics['expired_stock'] += qty
            elif days_to_expiry <= 30:
                stock_metrics['expiring_soon'] += qty
            elif days_to_expiry <= 90:
                stock_metrics['expiring_3months'] += qty
            elif days_to_expiry <= 180:
                stock_metrics['expiring_6months'] += qty
            else:
                stock_metrics['fresh_stock'] += qty
    
    stock_metrics['batch_details'] = sku_stock[['Batch_Number', 'Physical_Stock', 'Expiry_Date']].copy()
    
    return stock_metrics

def get_all_skus_from_data(df_sales, df_po, df_po_delivered, df_product):
    sku_set = set()
    if not df_product.empty:
        sku_set.update(df_product['SKU_ID'].dropna().unique())
    if not df_sales.empty:
        sku_set.update(df_sales['SKU_ID'].dropna().unique())
    if not df_po.empty:
        sku_set.update(df_po['SKU_ID'].dropna().unique())
    if not df_po_delivered.empty:
        sku_set.update(df_po_delivered['SKU_ID'].dropna().unique())
    
    sku_set = {s for s in sku_set if pd.notna(s) and str(s).strip() != ''}
    return sorted(list(sku_set))

def prepare_chart_data(sku_id, metrics):
    combined = []
    all_months = set()
    
    if not metrics['sales_data'].empty:
        all_months.update(metrics['sales_data']['Month'])
    if not metrics['po_data'].empty:
        all_months.update(metrics['po_data']['Month'])
    if not metrics['po_delivered_data'].empty:
        all_months.update(metrics['po_delivered_data']['Month'])
    
    if not all_months:
        return pd.DataFrame()
    
    sales_dict = {row['Month']: row['Sales_Qty'] for _, row in metrics['sales_data'].iterrows()} if not metrics['sales_data'].empty else {}
    po_dict = {row['Month']: row['PO_Qty'] for _, row in metrics['po_data'].iterrows()} if not metrics['po_data'].empty else {}
    po_delivered_dict = {row['Month']: row['PO_Delivered_Qty'] for _, row in metrics['po_delivered_data'].iterrows()} if not metrics['po_delivered_data'].empty else {}
    
    sorted_months = sorted(all_months)
    for month in sorted_months:
        combined.append({
            'Month': month,
            'Month_Label': month.strftime('%b %Y'),
            'Sales': sales_dict.get(month, 0),
            'PO': po_dict.get(month, 0),
            'PO_Delivered': po_delivered_dict.get(month, 0)
        })
    
    return pd.DataFrame(combined)

# =============================================================================
# ECHARTS VISUALIZATION FUNCTIONS
# =============================================================================

ddef create_trend_chart_echarts(main_df, compare_df=None, main_sku="", compare_sku=""):
    """Create ECharts trend chart with Sales & Inbound as Bar, PO as Line"""
    
    if main_df.empty:
        return None
    
    months = main_df['Month_Label'].tolist()
    
    # PERBAIKAN: Fungsi helper untuk membersihkan NaN
    def clean_data(series):
        return [0 if pd.isna(x) else int(x) for x in series]
    
    series = []
    
    # Main SKU - Sales (Bar)
    sales_data = clean_data(main_df['Sales'])
    series.append({
        "name": f"{main_sku} - Sales",
        "type": "bar",
        "data": sales_data,
        "itemStyle": {"color": "#10B981", "borderRadius": [4, 4, 0, 0]},
        "barWidth": "30%",
        "label": {
            "show": True,
            "position": "top",
            "fontSize": 10
        }
    })
    
    # Main SKU - Inbound (Bar)
    inbound_data = clean_data(main_df['PO_Delivered'])
    series.append({
        "name": f"{main_sku} - Inbound",
        "type": "bar",
        "data": inbound_data,
        "itemStyle": {"color": "#3B82F6", "borderRadius": [4, 4, 0, 0]},
        "barWidth": "30%",
        "label": {
            "show": True,
            "position": "top",
            "fontSize": 10
        }
    })
    
    # Main SKU - PO (Line)
    po_data = clean_data(main_df['PO'])
    series.append({
        "name": f"{main_sku} - PO",
        "type": "line",
        "data": po_data,
        "lineStyle": {"color": "#F59E0B", "width": 3, "type": "dashed"},
        "symbol": "diamond",
        "symbolSize": 8,
        "itemStyle": {"color": "#F59E0B"},
        "label": {
            "show": True,
            "position": "top",
            "fontSize": 10
        }
    })
    
    # Compare SKU (if exists)
    if compare_df is not None and not compare_df.empty:
        compare_sales = clean_data(compare_df['Sales'])
        series.append({
            "name": f"{compare_sku} - Sales",
            "type": "bar",
            "data": compare_sales,
            "itemStyle": {"color": "#A7F3D0", "borderRadius": [4, 4, 0, 0]},
            "barWidth": "30%",
            "label": {"show": True, "position": "top", "fontSize": 10}
        })
        
        compare_inbound = clean_data(compare_df['PO_Delivered'])
        series.append({
            "name": f"{compare_sku} - Inbound",
            "type": "bar",
            "data": compare_inbound,
            "itemStyle": {"color": "#BFDBFE", "borderRadius": [4, 4, 0, 0]},
            "barWidth": "30%",
            "label": {"show": True, "position": "top", "fontSize": 10}
        })
        
        compare_po = clean_data(compare_df['PO'])
        series.append({
            "name": f"{compare_sku} - PO",
            "type": "line",
            "data": compare_po,
            "lineStyle": {"color": "#FDE68A", "width": 3, "type": "dashed"},
            "symbol": "diamond",
            "symbolSize": 8,
            "itemStyle": {"color": "#FDE68A"},
            "label": {"show": True, "position": "top", "fontSize": 10}
        })
    
    options = {
        "title": {
            "text": "📈 Tren Sales vs PO vs Inbound",
            "left": "center",
            "textStyle": {"fontSize": 16, "fontWeight": "bold"}
        },
        "tooltip": {
            "trigger": "axis",
            "axisPointer": {"type": "shadow"}
        },
        "legend": {
            "bottom": 0,
            "type": "scroll",
            "pageIconColor": "#667eea"
        },
        "grid": {
            "left": "8%",
            "right": "5%",
            "bottom": "20%",
            "top": "15%",
            "containLabel": True
        },
        "xAxis": {
            "type": "category",
            "data": months,
            "axisLabel": {"rotate": 45, "fontSize": 11}
        },
        "yAxis": {
            "type": "value",
            "name": "Quantity (Units)",
            "axisLabel": {"fontSize": 11}
        },
        "series": series,
        "dataZoom": [
            {
                "type": "slider",
                "start": 0,
                "end": 100,
                "height": 20,
                "bottom": 25
            },
            {
                "type": "inside",
                "start": 0,
                "end": 100
            }
        ]
    }
    
    return options

def create_multi_brand_chart_echarts(df_filtered, display_metric="Quantity"):
    """Create multi-brand comparison chart with ECharts"""
    
    if df_filtered.empty:
        return None
    
    # Aggregate per brand per month
    brand_monthly = df_filtered.groupby(['Brand', df_filtered['Month'].dt.to_period('M')]).agg({
        'Sales_Qty': 'sum',
        'Sales_Value': 'sum'
    }).reset_index()
    brand_monthly['Month'] = brand_monthly['Month'].dt.to_timestamp()
    
    value_col = 'Sales_Qty' if display_metric == "Quantity" else 'Sales_Value'
    
    # Pivot data
    pivot_data = brand_monthly.pivot(index='Month', columns='Brand', values=value_col).fillna(0)
    pivot_data = pivot_data.sort_index()
    months = pivot_data.index.strftime('%b %Y').tolist()
    
    colors = ['#5470c6', '#91cc75', '#fac858', '#ee6666', '#73c0de', '#3ba272', '#fc8452', '#9a60b4', '#ea7ccc']
    
    series = []
    for i, brand in enumerate(pivot_data.columns):
        # PERBAIKAN: Bersihkan NaN dan konversi ke int
        brand_data = pivot_data[brand].fillna(0).round(0).astype(int).tolist()
        
        series.append({
            "name": brand,
            "type": "bar",
            "data": brand_data,
            "itemStyle": {
                "color": colors[i % len(colors)],
                "borderRadius": [4, 4, 0, 0]
            },
            "label": {
                "show": True,
                "position": "top",
                "fontSize": 10
            }
        })
    
    options = {
        "title": {
            "text": f"📊 Multi-Brand Sales Comparison ({display_metric})",
            "left": "center",
            "textStyle": {"fontSize": 16, "fontWeight": "bold"}
        },
        "tooltip": {
            "trigger": "axis",
            "axisPointer": {"type": "shadow"}
        },
        "legend": {
            "data": pivot_data.columns.tolist(),
            "bottom": 0,
            "type": "scroll"
        },
        "grid": {
            "left": "8%",
            "right": "5%",
            "bottom": "20%",
            "top": "15%",
            "containLabel": True
        },
        "xAxis": {
            "type": "category",
            "data": months,
            "axisLabel": {"rotate": 45, "fontSize": 11}
        },
        "yAxis": {
            "type": "value",
            "name": "Sales Quantity" if display_metric == "Quantity" else "Sales Value (Rp)",
            "axisLabel": {"fontSize": 11}
        },
        "series": series,
        "dataZoom": [
            {"type": "slider", "start": 0, "end": 100, "height": 20, "bottom": 25},
            {"type": "inside", "start": 0, "end": 100}
        ]
    }
    
    return options

def create_tier_chart_echarts(tier_performance):
    """Create tier performance chart with ECharts"""
    
    if tier_performance.empty:
        return None
    
    # PERBAIKAN: Bersihkan NaN
    tier_data = tier_performance['Total_Sales_Qty'].fillna(0).astype(int).tolist()
    tier_labels = tier_performance['SKU_Tier'].fillna('Unknown').tolist()
    
    options = {
        "title": {
            "text": "💎 Sales Quantity by SKU Tier",
            "left": "center",
            "textStyle": {"fontSize": 16, "fontWeight": "bold"}
        },
        "tooltip": {
            "trigger": "axis",
            "axisPointer": {"type": "shadow"}
        },
        "grid": {
            "left": "15%",
            "right": "5%",
            "bottom": "10%",
            "top": "15%"
        },
        "xAxis": {
            "type": "category",
            "data": tier_labels,
            "axisLabel": {"fontSize": 12, "fontWeight": "bold"}
        },
        "yAxis": {
            "type": "value",
            "name": "Sales Quantity",
            "axisLabel": {"fontSize": 11}
        },
        "series": [{
            "name": "Sales Qty",
            "type": "bar",
            "data": tier_data,
            "itemStyle": {
                "color": {
                    "type": "linear",
                    "x": 0, "y": 0, "x2": 0, "y2": 1,
                    "colorStops": [
                        {"offset": 0, "color": "#F59E0B"},
                        {"offset": 1, "color": "#EF4444"}
                    ]
                },
                "borderRadius": [4, 4, 0, 0]
            },
            "label": {
                "show": True,
                "position": "top",
                "fontSize": 11,
                "fontWeight": "bold"
            }
        }]
    }
    
    return options

def create_status_pie_chart_echarts(status_performance):
    """Create status distribution pie chart with ECharts"""
    
    if status_performance.empty:
        return None
    
    pie_data = []
    color_map = {'ACTIVE': '#10B981', 'INACTIVE': '#EF4444', 'UNKNOWN': '#9CA3AF'}
    
    for _, row in status_performance.iterrows():
        status = str(row['Status']) if pd.notna(row['Status']) else 'UNKNOWN'
        qty = int(row['Total_Sales_Qty']) if pd.notna(row['Total_Sales_Qty']) else 0
        
        pie_data.append({
            "value": qty,
            "name": status,
            "itemStyle": {"color": color_map.get(status.upper(), '#9CA3AF')}
        })
    
    options = {
        "title": {
            "text": "🥧 Sales Contribution by Status",
            "left": "center",
            "textStyle": {"fontSize": 16, "fontWeight": "bold"}
        },
        "tooltip": {
            "trigger": "item",
            "formatter": "{b}: {c} ({d}%)"
        },
        "legend": {
            "orient": "vertical",
            "left": "left",
            "top": "center"
        },
        "series": [{
            "name": "Sales by Status",
            "type": "pie",
            "radius": ["40%", "70%"],
            "avoidLabelOverlap": True,
            "itemStyle": {
                "borderRadius": 8,
                "borderColor": "#fff",
                "borderWidth": 2
            },
            "label": {
                "show": True,
                "position": "outside",
                "formatter": "{b}\n{d}%",
                "fontSize": 11
            },
            "emphasis": {
                "label": {"show": True},
                "scale": True,
                "scaleSize": 10
            },
            "data": pie_data
        }]
    }
    
    return options

def create_seasonality_chart_echarts(seasonality):
    """Create seasonality pattern chart with ECharts"""
    
    if seasonality.empty:
        return None
    
    # PERBAIKAN: Bersihkan NaN
    month_labels = seasonality['Month_Name'].fillna('Unknown').tolist()
    seasonal_data = seasonality['Seasonal_Index'].fillna(1.0).round(2).tolist()
    
    options = {
        "title": {
            "text": "🌙 Seasonal Index (Average Sales by Month)",
            "left": "center",
            "textStyle": {"fontSize": 16, "fontWeight": "bold"}
        },
        "tooltip": {
            "trigger": "axis",
            "axisPointer": {"type": "shadow"},
            "valueFormatter": "{value} x"
        },
        "grid": {
            "left": "10%",
            "right": "5%",
            "bottom": "10%",
            "top": "15%"
        },
        "xAxis": {
            "type": "category",
            "data": month_labels,
            "axisLabel": {"fontSize": 12}
        },
        "yAxis": {
            "type": "value",
            "name": "Seasonal Index",
            "axisLabel": {"fontSize": 11}
        },
        "series": [{
            "name": "Seasonal Index",
            "type": "bar",
            "data": seasonal_data,
            "itemStyle": {
                "color": {
                    "type": "linear",
                    "x": 0, "y": 0, "x2": 0, "y2": 1,
                    "colorStops": [
                        {"offset": 0, "color": "#3B82F6"},
                        {"offset": 1, "color": "#10B981"}
                    ]
                },
                "borderRadius": [4, 4, 0, 0]
            },
            "label": {
                "show": True,
                "position": "top",
                "formatter": "{c}x",
                "fontSize": 11
            },
            "markLine": {
                "data": [{"yAxis": 1, "name": "Average"}],
                "lineStyle": {"color": "#EF4444", "type": "dashed", "width": 2},
                "label": {"show": True, "position": "end", "formatter": "Average: 1.0"}
            }
        }]
    }
    
    return options

def create_pareto_chart_echarts(all_skus_pareto):
    """Create Pareto chart with ECharts"""
    
    if all_skus_pareto.empty:
        return None
    
    top_20 = all_skus_pareto.head(20).copy()
    
    # PERBAIKAN: Bersihkan NaN
    sku_labels = top_20['SKU_ID'].astype(str).fillna('Unknown').tolist()
    sales_data = top_20['Sales_Qty'].fillna(0).astype(int).tolist()
    cum_data = top_20['Cumulative_Percent'].fillna(0).round(1).tolist()
    
    options = {
        "title": {
            "text": "🎯 Top 20 SKUs by Sales (Pareto Analysis)",
            "left": "center",
            "textStyle": {"fontSize": 16, "fontWeight": "bold"}
        },
        "tooltip": {
            "trigger": "axis",
            "axisPointer": {"type": "shadow"}
        },
        "legend": {
            "data": ["Sales Qty", "Cumulative %"],
            "bottom": 0
        },
        "grid": {
            "left": "10%",
            "right": "8%",
            "bottom": "20%",
            "top": "15%"
        },
        "xAxis": {
            "type": "category",
            "data": sku_labels,
            "axisLabel": {"rotate": 45, "fontSize": 10}
        },
        "yAxis": [
            {
                "type": "value",
                "name": "Sales Quantity",
                "position": "left"
            },
            {
                "type": "value",
                "name": "Cumulative %",
                "position": "right",
                "min": 0,
                "max": 100,
                "axisLabel": {"formatter": "{value}%"}
            }
        ],
        "series": [
            {
                "name": "Sales Qty",
                "type": "bar",
                "data": sales_data,
                "itemStyle": {
                    "color": "#10B981",
                    "borderRadius": [4, 4, 0, 0]
                },
                "label": {
                    "show": True,
                    "position": "top",
                    "fontSize": 10
                }
            },
            {
                "name": "Cumulative %",
                "type": "line",
                "yAxisIndex": 1,
                "data": cum_data,
                "smooth": True,
                "lineStyle": {"color": "#F59E0B", "width": 3},
                "symbol": "circle",
                "symbolSize": 8,
                "itemStyle": {"color": "#F59E0B"},
                "label": {
                    "show": True,
                    "position": "top",
                    "formatter": "{c}%",
                    "fontSize": 10
                }
            }
        ],
        "dataZoom": [
            {"type": "slider", "start": 0, "end": 100, "height": 20, "bottom": 25}
        ]
    }
    
    return options

def create_mom_growth_chart_echarts(growth_metrics):
    """Create MoM growth chart with ECharts"""
    
    if growth_metrics.empty:
        return None
    
    months = growth_metrics['Month'].dt.strftime('%b %Y').tolist()
    
    # PERBAIKAN: Ganti NaN dengan 0
    growth_data = growth_metrics['MoM_Growth_Qty'].fillna(0).round(1).tolist()
    
    # Color based on positive/negative
    bar_colors = ['#10B981' if (g if g is not None else 0) >= 0 else '#EF4444' for g in growth_data]
    
    # Buat series data dengan itemStyle individual
    series_data = []
    for i, (growth, color) in enumerate(zip(growth_data, bar_colors)):
        # PERBAIKAN: Pastikan value bukan NaN
        value = growth if growth is not None and not pd.isna(growth) else 0
        series_data.append({
            "value": value,
            "itemStyle": {"color": color, "borderRadius": [4, 4, 0, 0]}
        })
    
    options = {
        "title": {
            "text": "📉 Month-over-Month Sales Growth",
            "left": "center",
            "textStyle": {"fontSize": 16, "fontWeight": "bold"}
        },
        "tooltip": {
            "trigger": "axis",
            "axisPointer": {"type": "shadow"},
            "valueFormatter": "{value}%"
        },
        "grid": {
            "left": "8%",
            "right": "5%",
            "bottom": "15%",
            "top": "15%"
        },
        "xAxis": {
            "type": "category",
            "data": months,
            "axisLabel": {"rotate": 45, "fontSize": 11}
        },
        "yAxis": {
            "type": "value",
            "name": "Growth (%)",
            "axisLabel": {"formatter": "{value}%"}
        },
        "series": [{
            "name": "MoM Growth",
            "type": "bar",
            "data": series_data,
            "label": {
                "show": True,
                "position": "top",
                "formatter": "{c}%",
                "fontSize": 11,
                "fontWeight": "bold"
            }
        }]
    }
    
    return options

def create_expiry_pie_echarts(expiry_df):
    """Create expiry distribution pie chart with ECharts"""
    
    if expiry_df.empty:
        return None
    
    color_map = {
        'Expired (< 0 days)': '#EF4444',
        'Expiring Soon (0-30 days)': '#F59E0B',
        'Expiring in 1-3 months': '#FBBF24',
        'Expiring in 3-6 months': '#60A5FA',
        'Fresh (> 6 months)': '#10B981'
    }
    
    pie_data = []
    for _, row in expiry_df.iterrows():
        category = str(row['Category']) if pd.notna(row['Category']) else 'Unknown'
        qty = int(row['Qty']) if pd.notna(row['Qty']) else 0
        
        pie_data.append({
            "value": qty,
            "name": category,
            "itemStyle": {"color": color_map.get(category, '#9CA3AF')}
        })
    
    options = {
        "title": {
            "text": "📦 Stock Distribution by Expiry Status",
            "left": "center",
            "textStyle": {"fontSize": 16, "fontWeight": "bold"}
        },
        "tooltip": {
            "trigger": "item",
            "formatter": "{b}: {c} units ({d}%)"
        },
        "legend": {
            "orient": "vertical",
            "left": "left",
            "top": "center"
        },
        "series": [{
            "name": "Expiry Distribution",
            "type": "pie",
            "radius": ["40%", "70%"],
            "avoidLabelOverlap": True,
            "itemStyle": {
                "borderRadius": 8,
                "borderColor": "#fff",
                "borderWidth": 2
            },
            "label": {
                "show": True,
                "position": "outside",
                "formatter": "{b}\n{c} ({d}%)",
                "fontSize": 10
            },
            "emphasis": {
                "label": {"show": True},
                "scale": True
            },
            "data": pie_data
        }]
    }
    
    return options

def prepare_sales_analysis_data(df_sales, df_product):
    """Prepare sales analysis data with product info"""
    if df_sales.empty:
        return pd.DataFrame()
    
    df = df_sales.copy()
    df['SKU_ID'] = df['SKU_ID'].astype(str)
    
    df['Brand'] = 'Unknown'
    df['SKU_Tier'] = 'Unknown'
    df['Status'] = 'UNKNOWN'
    df['Floor_Price'] = 0
    
    if not df_product.empty:
        df_product['SKU_ID'] = df_product['SKU_ID'].astype(str)
        
        product_cols = ['SKU_ID']
        if 'Brand' in df_product.columns:
            product_cols.append('Brand')
        if 'SKU_Tier' in df_product.columns:
            product_cols.append('SKU_Tier')
        if 'Status' in df_product.columns:
            product_cols.append('Status')
        if 'Floor_Price' in df_product.columns:
            product_cols.append('Floor_Price')
        
        df_product_subset = df_product[product_cols].copy()
        df = pd.merge(df, df_product_subset, on='SKU_ID', how='left', suffixes=('', '_prod'))
        
        if 'Brand_prod' in df.columns:
            df['Brand'] = df['Brand_prod'].fillna('Unknown')
        if 'SKU_Tier_prod' in df.columns:
            df['SKU_Tier'] = df['SKU_Tier_prod'].fillna('Unknown')
        if 'Status_prod' in df.columns:
            df['Status'] = df['Status_prod'].fillna('UNKNOWN')
        if 'Floor_Price_prod' in df.columns:
            df['Floor_Price'] = df['Floor_Price_prod'].fillna(0)
    
    df['Sales_Value'] = df['Sales_Qty'] * df['Floor_Price']
    df['Year'] = df['Month'].dt.year
    df['Month_Num'] = df['Month'].dt.month
    df['Month_Name'] = df['Month'].dt.strftime('%b')
    
    return df

def get_tier_performance(df):
    if df.empty:
        return pd.DataFrame()
    
    tier_stats = df.groupby('SKU_Tier').agg({
        'SKU_ID': 'nunique',
        'Sales_Qty': 'sum',
        'Sales_Value': 'sum'
    }).reset_index()
    tier_stats.columns = ['SKU_Tier', 'SKU_Count', 'Total_Sales_Qty', 'Total_Sales_Value']
    
    return tier_stats.sort_values('Total_Sales_Qty', ascending=False)

def get_status_performance(df):
    if df.empty:
        return pd.DataFrame()
    
    status_stats = df.groupby('Status').agg({
        'SKU_ID': 'nunique',
        'Sales_Qty': 'sum',
        'Sales_Value': 'sum'
    }).reset_index()
    status_stats.columns = ['Status', 'SKU_Count', 'Total_Sales_Qty', 'Total_Sales_Value']
    
    return status_stats

def calculate_pareto(df, top_percent=80):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()
    
    group_cols = ['SKU_ID']
    if 'Product_Name' in df.columns:
        group_cols.append('Product_Name')
    if 'Brand' in df.columns:
        group_cols.append('Brand')
    
    sku_sales = df.groupby(group_cols)['Sales_Qty'].sum().reset_index()
    sku_sales = sku_sales.sort_values('Sales_Qty', ascending=False)
    sku_sales['Cumulative_Qty'] = sku_sales['Sales_Qty'].cumsum()
    sku_sales['Cumulative_Percent'] = (sku_sales['Cumulative_Qty'] / sku_sales['Sales_Qty'].sum()) * 100
    
    pareto_skus = sku_sales[sku_sales['Cumulative_Percent'] <= top_percent]
    return pareto_skus, sku_sales

def get_seasonality_pattern(df):
    if df.empty:
        return pd.DataFrame()
    
    monthly_pattern = df.groupby('Month_Num').agg({
        'Sales_Qty': 'mean'
    }).reset_index()
    
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    monthly_pattern['Month_Name'] = monthly_pattern['Month_Num'].apply(lambda x: month_names[x-1] if 1 <= x <= 12 else 'Unknown')
    
    avg_qty = monthly_pattern['Sales_Qty'].mean()
    monthly_pattern['Seasonal_Index'] = monthly_pattern['Sales_Qty'] / avg_qty
    
    return monthly_pattern

def calculate_growth_metrics(df):
    if df.empty:
        return pd.DataFrame()
    
    monthly = df.groupby(df['Month'].dt.to_period('M')).agg({
        'Sales_Qty': 'sum'
    }).reset_index()
    monthly['Month'] = monthly['Month'].dt.to_timestamp()
    monthly = monthly.sort_values('Month')
    monthly['MoM_Growth_Qty'] = monthly['Sales_Qty'].pct_change() * 100
    
    return monthly

def get_diagnostics(metrics, sku_label=""):
    diagnoses = []
    prefix = f"**{sku_label}** " if sku_label else ""
    
    if not metrics['sales_data'].empty:
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
    elif not metrics['sales_data'].empty:
        diagnoses.append(("ℹ️", f"{prefix}Limited Data", f"Hanya {len(metrics['sales_data'])} bulan data. Pantau rutin.", "#6B7280"))
    else:
        diagnoses.append(("⚠️", f"{prefix}No Sales Data", "Belum ada riwayat penjualan. Order trial.", "#F59E0B"))
    
    if metrics['total_po'] > 0 and metrics['total_sales'] > 0:
        sell_through = (metrics['total_sales'] / metrics['total_po'] * 100)
        if sell_through < 40:
            diagnoses.append(("📦", f"{prefix}Low Sell-Through", f"Hanya {sell_through:.1f}% PO terjual. Risiko dead stock!", "#EF4444"))
        elif sell_through > 100:
            diagnoses.append(("🔥", f"{prefix}High Demand", f"Sales {sell_through:.0f}% > PO. Potensi lost sales!", "#F59E0B"))
    
    return diagnoses

# --- Load Data ---
client = init_gsheet_connection()
if client is None:
    st.stop()

with st.spinner('🔄 Loading data...'):
    all_data = load_data(client)
    
    df_product = all_data.get('product', pd.DataFrame())
    df_sales = all_data.get('sales', pd.DataFrame())
    df_po = all_data.get('po', pd.DataFrame())
    df_po_delivered = all_data.get('po_delivered', pd.DataFrame())
    df_stock = all_data.get('stock', pd.DataFrame())

# --- Prepare sales analysis data ---
df_sales_analysis = prepare_sales_analysis_data(df_sales, df_product)

# --- Create Tabs ---
tab_sku, tab_sales_analytics, tab_stock = st.tabs([
    "🔍 SKU Evaluator",
    "📊 Sales Analytics Pro",
    "📦 Stock Analysis"
])

# =============================================================================
# TAB 1: SKU EVALUATOR
# =============================================================================
with tab_sku:
    all_skus = get_all_skus_from_data(df_sales, df_po, df_po_delivered, df_product)
    
    if not all_skus:
        st.error("❌ Tidak ada data SKU ditemukan.")
        st.stop()
    
    # --- Control Panel ---
    st.markdown('<div class="control-panel">', unsafe_allow_html=True)
    
    col_brand, col_sku1, col_sku2 = st.columns([1.5, 2.5, 2.5])
    
    with col_brand:
        st.markdown('<div class="control-label">🏢 PILIH BRAND</div>', unsafe_allow_html=True)
        all_brands_filter = sorted(df_product['Brand'].dropna().unique().tolist())
        selected_brand_filter = st.selectbox("Filter Brand", ["Semua Brand"] + all_brands_filter, label_visibility="collapsed")
    
    with col_sku1:
        st.markdown('<div class="control-label">📦 SKU UTAMA</div>', unsafe_allow_html=True)
        sku_display_map = {}
        filtered_skus = all_skus
        if selected_brand_filter != "Semua Brand":
            brand_skus = df_product[df_product['Brand'] == selected_brand_filter]['SKU_ID'].tolist()
            filtered_skus = [s for s in all_skus if s in brand_skus]
        
        for sku in filtered_skus:
            product_row = df_product[df_product['SKU_ID'] == sku]
            if not product_row.empty:
                product_name = product_row.iloc[0].get('Product_Name', '')
                display = f"{sku} - {product_name}" if product_name else sku
            else:
                display = f"{sku} (No Product Data)"
            sku_display_map[display] = sku
        
        sku_display_list = list(sku_display_map.keys())
        selected_main_display = st.selectbox("SKU Utama", sku_display_list, key="main_sku", label_visibility="collapsed")
    
    with col_sku2:
        st.markdown('<div class="control-label">🔄 SKU PEMBANDING</div>', unsafe_allow_html=True)
        compare_options = ["[Tidak ada perbandingan]"] + sku_display_list
        selected_compare_display = st.selectbox("SKU Pembanding", compare_options, key="compare_sku", label_visibility="collapsed")
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- Get SKU IDs ---
    main_sku = sku_display_map[selected_main_display]
    compare_sku = sku_display_map[selected_compare_display] if selected_compare_display != "[Tidak ada perbandingan]" else None
    
    # --- Calculate Metrics ---
    main_metrics = calculate_sku_metrics(df_sales, df_po, df_po_delivered, main_sku, df_product)
    stock_metrics = calculate_stock_metrics(df_stock, df_product, main_sku)
    
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
    
    # --- Metric Cards ---
    avg_sales_3m = 0
    if not main_metrics['sales_data'].empty:
        sales_df = main_metrics['sales_data'].sort_values('Month')
        last_3_months = sales_df.tail(3)
        if not last_3_months.empty:
            avg_sales_3m = last_3_months['Sales_Qty'].mean()
    
    stock_cover = 0
    stock_cover_color = "#10B981"
    if stock_metrics['has_stock'] and avg_sales_3m > 0:
        stock_cover = stock_metrics['total_stock'] / avg_sales_3m
        if stock_cover < 1:
            stock_cover_color = "#EF4444"
        elif stock_cover > 6:
            stock_cover_color = "#F59E0B"
    
    col_m1, col_m2, col_m3 = st.columns(3)
    
    with col_m1:
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: #3B82F6;">
            <div class="metric-value">{stock_metrics['total_stock']:,.0f}</div>
            <div class="metric-label">📦 STOCK ONHAND</div>
            <div class="metric-sub">{stock_metrics['batch_count']} batch</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_m2:
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: #6366F1;">
            <div class="metric-value">{avg_sales_3m:.0f}</div>
            <div class="metric-label">📊 AVG SALES (Last 3M)</div>
            <div class="metric-sub">unit/bulan</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_m3:
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: {stock_cover_color};">
            <div class="metric-value">{stock_cover:.1f} <span style="font-size:0.8rem;">bln</span></div>
            <div class="metric-label">📦 STOCK COVER</div>
            <div class="metric-sub">Coverage</div>
        </div>
        """, unsafe_allow_html=True)
    
    # --- Detail Batch ---
    if stock_metrics['has_stock'] and not stock_metrics['batch_details'].empty:
        with st.expander("📋 Lihat Detail Batch", expanded=False):
            batch_df = stock_metrics['batch_details'].copy()
            batch_df['Expiry_Date'] = batch_df['Expiry_Date'].dt.strftime('%d %b %Y') if not batch_df['Expiry_Date'].isna().all() else 'N/A'
            batch_df = batch_df.rename(columns={
                'Batch_Number': 'Batch Number',
                'Physical_Stock': 'Stock Qty',
                'Expiry_Date': 'Expiry Date'
            })
            st.dataframe(batch_df, use_container_width=True, hide_index=True)
    
    # --- ECHARTS TREND CHART ---
    st.markdown("---")
    st.subheader("📈 Tren Sales vs PO vs Inbound")
    
    main_df = prepare_chart_data(main_sku, main_metrics)
    compare_df = prepare_chart_data(compare_sku, calculate_sku_metrics(df_sales, df_po, df_po_delivered, compare_sku, df_product)) if compare_sku else None
    
    if not main_df.empty:
        # Date range filter
        all_months = sorted(main_df['Month'].unique())
        month_options = [m.strftime('%b %Y') for m in all_months]
        
        if len(month_options) > 1:
            start_idx, end_idx = st.select_slider(
                "📅 Pilih Range Periode",
                options=month_options,
                value=(month_options[0], month_options[-1])
            )
            start_date = datetime.strptime(start_idx, '%b %Y')
            end_date = datetime.strptime(end_idx, '%b %Y')
            
            main_df_filtered = main_df[(main_df['Month'] >= start_date) & (main_df['Month'] <= end_date)].copy()
            compare_df_filtered = compare_df[(compare_df['Month'] >= start_date) & (compare_df['Month'] <= end_date)].copy() if compare_df is not None else None
            
            st.caption(f"📅 Menampilkan data dari **{start_idx}** hingga **{end_idx}**")
        else:
            main_df_filtered = main_df
            compare_df_filtered = compare_df
        
        # Render ECharts
        chart_options = create_trend_chart_echarts(
            main_df_filtered, 
            compare_df_filtered, 
            main_sku, 
            compare_sku if compare_sku else ""
        )
        
        if chart_options:
            st_echarts(options=chart_options, height="450px")
    else:
        st.info("📊 Tidak ada data Sales, PO, atau Inbound untuk SKU ini")
    
    # --- Diagnostics ---
    st.markdown("---")
    st.subheader("🩺 Smart Diagnostics")
    
    diagnoses = get_diagnostics(main_metrics, "")
    for icon, title, desc, color in diagnoses:
        bg_color = "#F0FDF4" if "🟢" in icon else "#FEF2F2" if "🔴" in icon or "📉" in icon else "#FFFBEB"
        st.markdown(f"""
        <div class="diagnostic-box" style="background:{bg_color}; border-left-color:{color};">
            <div class="diagnostic-title"><span style="font-size:1rem;">{icon}</span> {title}</div>
            <div class="diagnostic-desc">{desc}</div>
        </div>
        """, unsafe_allow_html=True)

# =============================================================================
# TAB 2: SALES ANALYTICS PRO
# =============================================================================
with tab_sales_analytics:
    st.subheader("📊 Sales Analytics Pro")
    st.caption("Deep Dive Analysis with Interactive ECharts")
    
    if df_sales_analysis.empty:
        st.warning("⚠️ Tidak ada data sales untuk dianalisis.")
    else:
        # --- Filter Panel ---
        st.markdown('<div class="control-panel">', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1.5, 2, 1])
        
        with col1:
            all_dates = sorted(df_sales_analysis['Month'].unique())
            date_options = [d.strftime('%b %Y') for d in all_dates]
            if len(date_options) > 1:
                start_idx, end_idx = st.select_slider(
                    "📅 Range Periode",
                    options=date_options,
                    value=(date_options[0], date_options[-1])
                )
                start_date = datetime.strptime(start_idx, '%b %Y')
                end_date = datetime.strptime(end_idx, '%b %Y')
            else:
                start_date = all_dates[0]
                end_date = all_dates[-1]
        
        with col2:
            all_brands = sorted(df_sales_analysis['Brand'].unique())
            selected_brands = st.multiselect(
                "🏷️ Pilih Brand",
                options=all_brands,
                default=all_brands[:5] if len(all_brands) > 5 else all_brands
            )
        
        with col3:
            display_metric = st.selectbox("📊 Metric", ["Quantity", "Revenue"])
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Filter data
        df_filtered = df_sales_analysis[
            (df_sales_analysis['Month'] >= start_date) & 
            (df_sales_analysis['Month'] <= end_date)
        ].copy()
        
        if selected_brands:
            df_filtered = df_filtered[df_filtered['Brand'].isin(selected_brands)]
        
        if not df_filtered.empty:
            # Summary Metrics
            total_qty = df_filtered['Sales_Qty'].sum()
            total_value = df_filtered['Sales_Value'].sum()
            unique_skus = df_filtered['SKU_ID'].nunique()
            unique_brands = df_filtered['Brand'].nunique()
            
            col_s1, col_s2, col_s3, col_s4 = st.columns(4)
            with col_s1:
                st.metric("Total Sales Qty", f"{total_qty:,.0f}")
            with col_s2:
                st.metric("Total Sales Value", format_rupiah(total_value))
            with col_s3:
                st.metric("Active SKU", f"{unique_skus:,}")
            with col_s4:
                st.metric("Active Brand", f"{unique_brands:,}")
            
            # Multi-Brand Chart
            st.markdown("---")
            brand_chart = create_multi_brand_chart_echarts(df_filtered, display_metric)
            if brand_chart:
                st_echarts(options=brand_chart, height="450px")
            
            # Tier & Status Analysis
            st.markdown("---")
            col_t1, col_t2 = st.columns(2)
            
            with col_t1:
                tier_perf = get_tier_performance(df_filtered)
                if not tier_perf.empty:
                    tier_chart = create_tier_chart_echarts(tier_perf)
                    if tier_chart:
                        st_echarts(options=tier_chart, height="350px")
            
            with col_t2:
                status_perf = get_status_performance(df_filtered)
                if not status_perf.empty:
                    status_chart = create_status_pie_echarts(status_perf)
                    if status_chart:
                        st_echarts(options=status_chart, height="350px")
            
            # Seasonality & Pareto
            st.markdown("---")
            col_szn1, col_szn2 = st.columns(2)
            
            with col_szn1:
                seasonality = get_seasonality_pattern(df_filtered)
                if not seasonality.empty:
                    szn_chart = create_seasonality_chart_echarts(seasonality)
                    if szn_chart:
                        st_echarts(options=szn_chart, height="350px")
            
            with col_szn2:
                pareto_80, all_skus = calculate_pareto(df_filtered, 80)
                if not all_skus.empty:
                    pareto_chart = create_pareto_chart_echarts(all_skus)
                    if pareto_chart:
                        st_echarts(options=pareto_chart, height="350px")
                    
                    st.markdown(f"""
                    <div class="insight-card">
                        <div class="insight-title">📊 PARETO INSIGHT</div>
                        <div class="insight-value">{len(pareto_80)} SKU</div>
                        <div class="insight-desc">menyumbang {pareto_80['Cumulative_Percent'].iloc[-1]:.1f}% dari total sales</div>
                    </div>
                    """, unsafe_allow_html=True)
            
            # MoM Growth
            st.markdown("---")
            growth_metrics = calculate_growth_metrics(df_filtered)
            if not growth_metrics.empty:
                growth_chart = create_mom_growth_chart_echarts(growth_metrics)
                if growth_chart:
                    st_echarts(options=growth_chart, height="400px")
        else:
            st.warning("⚠️ Tidak ada data untuk filter yang dipilih.")

# =============================================================================
# TAB 3: STOCK ANALYSIS
# =============================================================================
with tab_stock:
    st.subheader("📦 Stock Onhand Analysis")
    
    if df_stock.empty:
        st.warning("⚠️ Tidak ada data stok yang tersedia.")
    else:
        all_stock_skus = sorted(df_stock['SKU_ID'].dropna().unique())
        
        selected_stock_sku = st.selectbox("🔍 Pilih SKU", all_stock_skus)
        
        product_info = df_product[df_product['SKU_ID'] == selected_stock_sku]
        product_name = product_info.iloc[0].get('Product_Name', selected_stock_sku) if not product_info.empty else selected_stock_sku
        brand = product_info.iloc[0].get('Brand', '-') if not product_info.empty else '-'
        
        st.markdown(f"""
        <div class="sku-header" style="background: linear-gradient(135deg, #3B82F6 0%, #1D4ED8 100%);">
            <div>
                <div class="sku-title">{product_name} <span style="font-size:0.8rem;">({selected_stock_sku})</span></div>
                <div class="sku-badges"><span class="badge">🏷️ {brand}</span></div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        stock_m = calculate_stock_metrics(df_stock, df_product, selected_stock_sku)
        
        if stock_m['has_stock']:
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Stock", f"{stock_m['total_stock']:,.0f}")
            with col2:
                st.metric("Fresh Stock", f"{stock_m['fresh_stock']:,.0f}")
            with col3:
                st.metric("Expiring ≤6m", f"{stock_m['expiring_6months'] + stock_m['expiring_3months'] + stock_m['expiring_soon']:,.0f}")
            with col4:
                st.metric("Expired", f"{stock_m['expired_stock']:,.0f}")
            
            # Expiry Pie Chart
            expiry_data = []
            if stock_m['expired_stock'] > 0:
                expiry_data.append({'Category': 'Expired (< 0 days)', 'Qty': stock_m['expired_stock']})
            if stock_m['expiring_soon'] > 0:
                expiry_data.append({'Category': 'Expiring Soon (0-30 days)', 'Qty': stock_m['expiring_soon']})
            if stock_m['expiring_3months'] > 0:
                expiry_data.append({'Category': 'Expiring in 1-3 months', 'Qty': stock_m['expiring_3months']})
            if stock_m['expiring_6months'] > 0:
                expiry_data.append({'Category': 'Expiring in 3-6 months', 'Qty': stock_m['expiring_6months']})
            if stock_m['fresh_stock'] > 0:
                expiry_data.append({'Category': 'Fresh (> 6 months)', 'Qty': stock_m['fresh_stock']})
            
            if expiry_data:
                expiry_df = pd.DataFrame(expiry_data)
                expiry_chart = create_expiry_pie_echarts(expiry_df)
                if expiry_chart:
                    st_echarts(options=expiry_chart, height="400px")
            
            # Batch Details
            st.markdown("---")
            st.subheader("📋 Batch Details")
            batch_df = stock_m['batch_details'].copy()
            batch_df['Expiry_Date'] = batch_df['Expiry_Date'].dt.strftime('%d %b %Y')
            st.dataframe(batch_df, use_container_width=True, hide_index=True)
        else:
            st.info(f"📦 Tidak ada data stok untuk SKU {selected_stock_sku}")

# --- Footer ---
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #888; font-size: 0.7rem; padding: 0.5rem;">
    <p>📊 SKU 360° Evaluator Pro | Powered by ECharts | Data Sales 2025-2026</p>
</div>
""", unsafe_allow_html=True)
