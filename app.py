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
        
        # 2.5. Data PO Delivered (Inbound) - NEW
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
                
        except Exception as e:
            st.warning(f"⚠️ Gagal load Data_PO_Delivered(Inbound): {str(e)}")
            data['po_delivered'] = pd.DataFrame()
        
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
        
        # 5. Stock Onhand
        data['stock'] = pd.DataFrame()
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
                
        except Exception as e:
            st.warning(f"⚠️ Gagal load Stock_Onhand: {str(e)}")
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

def calculate_sku_metrics(df_sales, df_po, df_po_delivered, sku_id, df_product):
    """Hitung metrik untuk SKU tertentu termasuk PO Delivered"""
    metrics = {
        'sales_data': pd.DataFrame(),
        'po_data': pd.DataFrame(),
        'po_delivered_data': pd.DataFrame(),
        'total_sales': 0,
        'total_po': 0,
        'total_po_delivered': 0,
        'avg_monthly_sales': 0,
        'avg_monthly_po': 0,
        'avg_monthly_po_delivered': 0,
        'months_with_sales': 0,
        'months_with_po': 0,
        'months_with_po_delivered': 0,
        'floor_price': 0,
        'purchase_price': 0,
        'status': 'NOT_FOUND',
        'product_name': sku_id,
        'brand': '-',
        'tier': '-',
        'moq': 0,
        'first_sales_month': None,
        'last_sales_month': None,
        'first_po_month': None,
        'last_po_month': None,
        'first_po_delivered_month': None,
        'last_po_delivered_month': None
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
        sales_monthly = sales_sku.groupby('Month')['Sales_Qty'].sum().reset_index()
        sales_monthly = sales_monthly.sort_values('Month')
        metrics['sales_data'] = sales_monthly
        metrics['total_sales'] = sales_monthly['Sales_Qty'].sum()
        metrics['months_with_sales'] = len(sales_monthly)
        if not sales_monthly.empty:
            metrics['avg_monthly_sales'] = sales_monthly['Sales_Qty'].mean()
            metrics['first_sales_month'] = sales_monthly['Month'].min()
            metrics['last_sales_month'] = sales_monthly['Month'].max()
    
    po_sku = df_po[df_po['SKU_ID'] == sku_id].copy() if not df_po.empty else pd.DataFrame()
    if not po_sku.empty:
        po_monthly = po_sku.groupby('Month')['PO_Qty'].sum().reset_index()
        po_monthly = po_monthly.sort_values('Month')
        metrics['po_data'] = po_monthly
        metrics['total_po'] = po_monthly['PO_Qty'].sum()
        metrics['months_with_po'] = len(po_monthly)
        if not po_monthly.empty:
            metrics['avg_monthly_po'] = po_monthly['PO_Qty'].mean()
            metrics['first_po_month'] = po_monthly['Month'].min()
            metrics['last_po_month'] = po_monthly['Month'].max()
    
    po_delivered_sku = df_po_delivered[df_po_delivered['SKU_ID'] == sku_id].copy() if not df_po_delivered.empty else pd.DataFrame()
    if not po_delivered_sku.empty:
        po_delivered_monthly = po_delivered_sku.groupby('Month')['PO_Delivered_Qty'].sum().reset_index()
        po_delivered_monthly = po_delivered_monthly.sort_values('Month')
        metrics['po_delivered_data'] = po_delivered_monthly
        metrics['total_po_delivered'] = po_delivered_monthly['PO_Delivered_Qty'].sum()
        metrics['months_with_po_delivered'] = len(po_delivered_monthly)
        if not po_delivered_monthly.empty:
            metrics['avg_monthly_po_delivered'] = po_delivered_monthly['PO_Delivered_Qty'].mean()
            metrics['first_po_delivered_month'] = po_delivered_monthly['Month'].min()
            metrics['last_po_delivered_month'] = po_delivered_monthly['Month'].max()
    
    return metrics

def calculate_stock_metrics(df_stock, df_product, sku_id, avg_monthly_sales=0):
    """Hitung metrik stock untuk SKU tertentu (dengan multi-batch)"""
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
    stock_metrics['batch_details'] = stock_metrics['batch_details'].sort_values('Expiry_Date')
    
    return stock_metrics

def get_stock_health_status(stock_metrics, avg_monthly_sales):
    """Tentukan status kesehatan stock"""
    if not stock_metrics['has_stock']:
        return "⚪ No Stock", "#9CA3AF", "Tidak ada stok tercatat"
    
    total_stock = stock_metrics['total_stock']
    
    if stock_metrics['expired_stock'] > 0:
        return "🔴 Expired Stock Detected", "#EF4444", f"{stock_metrics['expired_stock']:,.0f} unit sudah expired"
    
    if stock_metrics['expiring_soon'] > 0:
        return "🟠 Expiring Soon (<30 days)", "#F59E0B", f"{stock_metrics['expiring_soon']:,.0f} unit akan expired dalam 30 hari"
    
    if stock_metrics['expiring_3months'] > 0:
        return "🟡 Expiring in 1-3 Months", "#FBBF24", f"{stock_metrics['expiring_3months']:,.0f} unit akan expired dalam 1-3 bulan"
    
    if avg_monthly_sales > 0:
        cover_months = total_stock / avg_monthly_sales
        if cover_months < 1:
            return "🔴 Low Stock", "#EF4444", f"Stok hanya cukup untuk {cover_months:.1f} bulan"
        elif cover_months > 6:
            return "🟠 Overstock", "#F59E0B", f"Stok cukup untuk {cover_months:.1f} bulan"
        else:
            return "🟢 Healthy Stock", "#10B981", f"Stok cukup untuk {cover_months:.1f} bulan"
    
    return "🟢 Stock Available", "#10B981", f"Total stok {total_stock:,.0f} unit"

def get_all_skus_from_data(df_sales, df_po, df_po_delivered, df_product):
    """Dapatkan semua SKU unik dari Sales, PO, PO Delivered, dan Product Master"""
    sku_set = set()
    
    if not df_sales.empty:
        sku_set.update(df_sales['SKU_ID'].dropna().unique())
    
    if not df_po.empty:
        sku_set.update(df_po['SKU_ID'].dropna().unique())
    
    if not df_po_delivered.empty:
        sku_set.update(df_po_delivered['SKU_ID'].dropna().unique())
    
    if not df_product.empty:
        sku_set.update(df_product['SKU_ID'].dropna().unique())
    
    sku_set = {s for s in sku_set if pd.notna(s) and str(s).strip() != ''}
    return sorted(list(sku_set))

def prepare_chart_data(sku_id, metrics):
    """Menyiapkan data chart untuk periode yang memiliki data (Sales, PO, PO Delivered)"""
    combined = []
    all_months = set()
    
    if not metrics['sales_data'].empty:
        for month in metrics['sales_data']['Month']:
            all_months.add(month)
    
    if not metrics['po_data'].empty:
        for month in metrics['po_data']['Month']:
            all_months.add(month)
    
    if not metrics['po_delivered_data'].empty:
        for month in metrics['po_delivered_data']['Month']:
            all_months.add(month)
    
    if not all_months:
        return pd.DataFrame()
    
    sales_dict = {}
    if not metrics['sales_data'].empty:
        for _, row in metrics['sales_data'].iterrows():
            sales_dict[row['Month']] = row['Sales_Qty']
    
    po_dict = {}
    if not metrics['po_data'].empty:
        for _, row in metrics['po_data'].iterrows():
            po_dict[row['Month']] = row['PO_Qty']
    
    po_delivered_dict = {}
    if not metrics['po_delivered_data'].empty:
        for _, row in metrics['po_delivered_data'].iterrows():
            po_delivered_dict[row['Month']] = row['PO_Delivered_Qty']
    
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
# FUNGSI ANALISIS LANJUTAN UNTUK TAB SALES ANALYTICS PRO
# =============================================================================

def prepare_sales_analysis_data(df_sales, df_product):
    """Siapkan data untuk analisis sales dengan menggabungkan info produk"""
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
            df = df.drop(columns=['Brand_prod'])
        if 'SKU_Tier_prod' in df.columns:
            df['SKU_Tier'] = df['SKU_Tier_prod'].fillna('Unknown')
        if 'Status_prod' in df.columns:
            df['Status'] = df['Status_prod'].fillna('UNKNOWN')
        if 'Floor_Price_prod' in df.columns:
            df['Floor_Price'] = df['Floor_Price_prod'].fillna(0)
    
    if 'Brand' not in df.columns:
        df['Brand'] = 'Unknown'
    if 'SKU_Tier' not in df.columns:
        df['SKU_Tier'] = 'Unknown'
    if 'Status' not in df.columns:
        df['Status'] = 'UNKNOWN'
    if 'Floor_Price' not in df.columns:
        df['Floor_Price'] = 0
    
    df['Sales_Value'] = df['Sales_Qty'] * df['Floor_Price']
    df['Year'] = df['Month'].dt.year
    df['Month_Num'] = df['Month'].dt.month
    df['Quarter'] = df['Month'].dt.quarter
    df['Month_Name'] = df['Month'].dt.strftime('%b')
    
    return df

def get_top_brands(df, metric='Sales_Qty', n=10):
    if df.empty:
        return pd.DataFrame()
    
    if metric == 'Sales_Qty':
        result = df.groupby('Brand')['Sales_Qty'].sum().reset_index()
        result.columns = ['Brand', 'Total_Sales_Qty']
    else:
        result = df.groupby('Brand')['Sales_Value'].sum().reset_index()
        result.columns = ['Brand', 'Total_Sales_Value']
    
    result['Share_Percent'] = (result.iloc[:, 1] / result.iloc[:, 1].sum() * 100)
    return result.sort_values(result.columns[1], ascending=False).head(n)

def calculate_growth_metrics(df):
    if df.empty:
        return pd.DataFrame()
    
    monthly = df.groupby(df['Month'].dt.to_period('M')).agg({
        'Sales_Qty': 'sum',
        'Sales_Value': 'sum'
    }).reset_index()
    monthly['Month'] = monthly['Month'].dt.to_timestamp()
    monthly = monthly.sort_values('Month')
    monthly['MoM_Growth_Qty'] = monthly['Sales_Qty'].pct_change() * 100
    monthly['MoM_Growth_Value'] = monthly['Sales_Value'].pct_change() * 100
    
    return monthly

def get_tier_performance(df):
    if df.empty:
        return pd.DataFrame()
    
    if 'SKU_Tier' not in df.columns:
        df = df.copy()
        df['SKU_Tier'] = 'Unknown'
    
    tier_stats = df.groupby('SKU_Tier').agg({
        'SKU_ID': 'nunique',
        'Sales_Qty': 'sum',
        'Sales_Value': 'sum'
    }).reset_index()
    tier_stats.columns = ['SKU_Tier', 'SKU_Count', 'Total_Sales_Qty', 'Total_Sales_Value']
    tier_stats['Avg_Qty_per_SKU'] = tier_stats['Total_Sales_Qty'] / tier_stats['SKU_Count']
    tier_stats['Share_Qty'] = (tier_stats['Total_Sales_Qty'] / tier_stats['Total_Sales_Qty'].sum() * 100)
    tier_stats['Share_Value'] = (tier_stats['Total_Sales_Value'] / tier_stats['Total_Sales_Value'].sum() * 100)
    
    return tier_stats.sort_values('Total_Sales_Qty', ascending=False)

def get_status_performance(df):
    if df.empty:
        return pd.DataFrame()
    
    if 'Status' not in df.columns:
        df = df.copy()
        df['Status'] = 'UNKNOWN'
    
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
    if 'SKU_Tier' in df.columns:
        group_cols.append('SKU_Tier')
    
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
        'Sales_Qty': 'mean',
        'Sales_Value': 'mean'
    }).reset_index()
    
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    monthly_pattern['Month_Name'] = monthly_pattern['Month_Num'].apply(lambda x: month_names[x-1] if 1 <= x <= 12 else 'Unknown')
    
    avg_qty = monthly_pattern['Sales_Qty'].mean()
    monthly_pattern['Seasonal_Index'] = monthly_pattern['Sales_Qty'] / avg_qty
    
    return monthly_pattern

def get_brand_growth_matrix(df):
    if df.empty:
        return pd.DataFrame()
    
    df_2025 = df[df['Year'] == 2025].copy() if 2025 in df['Year'].values else pd.DataFrame()
    df_2026 = df[df['Year'] == 2026].copy() if 2026 in df['Year'].values else pd.DataFrame()
    
    brand_stats = []
    
    if 'Brand' not in df.columns:
        return pd.DataFrame()
    
    brands = df['Brand'].unique()
    
    for brand in brands:
        sales_2025 = df_2025[df_2025['Brand'] == brand]['Sales_Qty'].sum() if not df_2025.empty else 0
        sales_2026 = df_2026[df_2026['Brand'] == brand]['Sales_Qty'].sum() if not df_2026.empty else 0
        
        growth = ((sales_2026 - sales_2025) / sales_2025 * 100) if sales_2025 > 0 else 0
        total_sales = sales_2025 + sales_2026
        
        brand_stats.append({
            'Brand': brand,
            'Sales_2025': sales_2025,
            'Sales_2026': sales_2026,
            'Growth_2026': growth,
            'Total_Sales': total_sales
        })
    
    df_brand = pd.DataFrame(brand_stats)
    if df_brand.empty:
        return pd.DataFrame()
    
    total_market = df_brand['Total_Sales'].sum()
    df_brand['Market_Share'] = (df_brand['Total_Sales'] / total_market * 100) if total_market > 0 else 0
    
    def categorize(growth, share):
        if growth > 10 and share > 10:
            return "🌟 Star (High Growth, High Share)"
        elif growth > 10:
            return "🚀 Question Mark (High Growth, Low Share)"
        elif share > 10:
            return "💰 Cash Cow (Low Growth, High Share)"
        else:
            return "🐕 Dog (Low Growth, Low Share)"
    
    df_brand['Category'] = df_brand.apply(lambda x: categorize(x['Growth_2026'], x['Market_Share']), axis=1)
    
    return df_brand.sort_values('Total_Sales', ascending=False)

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
    
    st.markdown('<div class="control-panel">', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    main_sku = sku_display_map[selected_main_display]
    compare_sku = sku_display_map[selected_compare_display] if selected_compare_display != "[Tidak ada perbandingan]" else None
    
    main_metrics = calculate_sku_metrics(df_sales, df_po, df_po_delivered, main_sku, df_product)
    
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
    
    # METRIC CARDS
    with st.expander("📊 Lihat Detail Metrik SKU", expanded=False):
        col1, col2, col3, col4, col5 = st.columns(5)
        
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
            <div class="metric-card" style="border-top-color: #3B82F6;">
                <div class="metric-value">{main_metrics['total_po_delivered']:,.0f}</div>
                <div class="metric-label">📥 TOTAL INBOUND</div>
                <div class="metric-sub">{main_metrics['months_with_po_delivered']} bulan inbound</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class="metric-card" style="border-top-color: #6366F1;">
                <div class="metric-value">{main_metrics['avg_monthly_sales']:.0f}</div>
                <div class="metric-label">📊 AVG MONTHLY SALES</div>
                <div class="metric-sub">Rata-rata per bulan</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col5:
            st.markdown(f"""
            <div class="metric-card" style="border-top-color: #8B5CF6;">
                <div class="metric-value">{main_metrics['avg_monthly_po']:.0f}</div>
                <div class="metric-label">🎯 AVG MONTHLY PO</div>
                <div class="metric-sub">Rata-rata order per bulan</div>
            </div>
            """, unsafe_allow_html=True)
    
    # STOCK METRIC CARDS
    st.markdown("---")
    st.subheader("📦 Stock Onhand Analysis")
    
    stock_metrics = calculate_stock_metrics(df_stock, df_product, main_sku, main_metrics['avg_monthly_sales'])
    
    col_s1, col_s2, col_s3, col_s4 = st.columns(4)
    
    with col_s1:
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: #3B82F6;">
            <div class="metric-value">{stock_metrics['total_stock']:,.0f}</div>
            <div class="metric-label">📦 TOTAL STOCK</div>
            <div class="metric-sub">{stock_metrics['batch_count']} batch</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_s2:
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: #10B981;">
            <div class="metric-value">{stock_metrics['fresh_stock']:,.0f}</div>
            <div class="metric-label">✅ FRESH STOCK</div>
            <div class="metric-sub">> 6 bulan</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_s3:
        warning_stock = stock_metrics['expiring_soon'] + stock_metrics['expiring_3months'] + stock_metrics['expiring_6months']
        color = "#EF4444" if stock_metrics['expiring_soon'] > 0 else "#F59E0B" if stock_metrics['expiring_3months'] > 0 else "#6B7280"
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: {color};">
            <div class="metric-value">{warning_stock:,.0f}</div>
            <div class="metric-label">⚠️ EXPIRING SOON</div>
            <div class="metric-sub">≤ 6 bulan</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_s4:
        health_status, health_color, health_desc = get_stock_health_status(stock_metrics, main_metrics['avg_monthly_sales'])
        st.markdown(f"""
        <div class="metric-card" style="border-top-color: {health_color};">
            <div class="metric-value" style="font-size: 1rem;">{health_status}</div>
            <div class="metric-label">🏥 STOCK HEALTH</div>
            <div class="metric-sub">{health_desc}</div>
        </div>
        """, unsafe_allow_html=True)
    
    if stock_metrics['has_stock'] and not stock_metrics['batch_details'].empty:
        with st.expander("📋 Lihat Detail Batch per SKU", expanded=False):
            batch_df = stock_metrics['batch_details'].copy()
            batch_df['Expiry_Date'] = batch_df['Expiry_Date'].dt.strftime('%d %b %Y') if not batch_df['Expiry_Date'].isna().all() else 'N/A'
            batch_df = batch_df.rename(columns={
                'Batch_Number': 'Batch Number',
                'Physical_Stock': 'Stock Qty',
                'Expiry_Date': 'Expiry Date'
            })
            st.dataframe(batch_df, use_container_width=True, hide_index=True)
            
            if stock_metrics['expired_stock'] > 0:
                st.error(f"⚠️ Terdapat {stock_metrics['expired_stock']:,.0f} unit stok yang sudah EXPIRED! Segera lakukan disposisi.")
            elif stock_metrics['expiring_soon'] > 0:
                st.warning(f"⚠️ Terdapat {stock_metrics['expiring_soon']:,.0f} unit stok yang akan EXPIRED dalam 30 hari.")
            elif stock_metrics['expiring_3months'] > 0:
                st.info(f"ℹ️ Terdapat {stock_metrics['expiring_3months']:,.0f} unit stok yang akan EXPIRED dalam 1-3 bulan.")
    else:
        st.info("📦 Tidak ada data stok untuk SKU ini")
    
    # TREND CHART - COMBO CHART (Sales & Inbound = Bar, PO = Line)
    st.markdown("---")
    st.subheader("📈 Tren Sales vs PO vs Inbound")
    
    main_df = prepare_chart_data(main_sku, main_metrics)
    
    if compare_sku:
        compare_metrics = calculate_sku_metrics(df_sales, df_po, df_po_delivered, compare_sku, df_product)
        compare_df = prepare_chart_data(compare_sku, compare_metrics)
    else:
        compare_metrics = None
        compare_df = None
    
    if not main_df.empty:
        fig = go.Figure()
        
        # === SKU UTAMA ===
        # Sales - Bar Chart (Hijau)
        fig.add_trace(go.Bar(
            x=main_df['Month_Label'], y=main_df['Sales'],
            name=f'{main_sku} - Sales',
            marker_color='#10B981',
            text=main_df['Sales'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='outside',
            opacity=0.8
        ))
        
        # Inbound (PO Delivered) - Bar Chart (Biru)
        fig.add_trace(go.Bar(
            x=main_df['Month_Label'], y=main_df['PO_Delivered'],
            name=f'{main_sku} - Inbound',
            marker_color='#3B82F6',
            text=main_df['PO_Delivered'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='outside',
            opacity=0.8
        ))
        
        # PO - Line Chart (Orange, dashed)
        fig.add_trace(go.Scatter(
            x=main_df['Month_Label'], y=main_df['PO'],
            name=f'{main_sku} - PO',
            mode='lines+markers',
            line=dict(color='#F59E0B', width=3, dash='dash'),
            marker=dict(size=8, color='#F59E0B', symbol='diamond'),
            text=main_df['PO'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
            textposition='top center'
        ))
        
        # === SKU PEMBANDING (jika ada) ===
        if compare_df is not None and not compare_df.empty:
            # Sales - Bar Chart (Hijau Muda)
            fig.add_trace(go.Bar(
                x=compare_df['Month_Label'], y=compare_df['Sales'],
                name=f'{compare_sku} - Sales',
                marker_color='#A7F3D0',
                text=compare_df['Sales'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
                textposition='outside',
                opacity=0.7
            ))
            
            # Inbound - Bar Chart (Biru Muda)
            fig.add_trace(go.Bar(
                x=compare_df['Month_Label'], y=compare_df['PO_Delivered'],
                name=f'{compare_sku} - Inbound',
                marker_color='#BFDBFE',
                text=compare_df['PO_Delivered'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
                textposition='outside',
                opacity=0.7
            ))
            
            # PO - Line Chart (Orange Muda, dashed)
            fig.add_trace(go.Scatter(
                x=compare_df['Month_Label'], y=compare_df['PO'],
                name=f'{compare_sku} - PO',
                mode='lines+markers',
                line=dict(color='#FDE68A', width=3, dash='dash'),
                marker=dict(size=8, color='#FDE68A', symbol='diamond'),
                text=compare_df['PO'].apply(lambda x: f"{x:,.0f}" if x > 0 else ""),
                textposition='top center'
            ))
        
        # Update layout untuk barmode='group' agar bar berdampingan
        fig.update_layout(
            height=450,
            xaxis_title='Periode',
            yaxis_title='Quantity (Units)',
            barmode='group',  # Bar berdampingan
            hovermode='x unified',
            plot_bgcolor='white',
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
            margin=dict(t=40, b=40, l=20, r=20)
        )
        
        st.plotly_chart(fig, use_container_width=True)
        st.caption(f"📅 Periode data: {main_df['Month_Label'].iloc[0]} - {main_df['Month_Label'].iloc[-1]} | Total {len(main_df)} bulan | 📊 Bar = Sales & Inbound | 📈 Line = PO")
    else:
        st.info("📊 Tidak ada data Sales, PO, atau Inbound untuk SKU ini")
    
    # SMART DIAGNOSTICS
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
        
        # Inbound vs PO comparison
        if metrics['total_po'] > 0 and metrics['total_po_delivered'] > 0:
            inbound_rate = (metrics['total_po_delivered'] / metrics['total_po'] * 100)
            if inbound_rate < 80:
                diagnoses.append(("⚠️", f"{prefix}Low Inbound Rate", f"Hanya {inbound_rate:.1f}% PO yang sudah masuk. Cek status pengiriman!", "#F59E0B"))
        
        return diagnoses
    
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
    
    if compare_sku and compare_metrics:
        st.markdown(f'<div class="small-text" style="margin-top:8px;">🔍 <strong>Perbandingan dengan {compare_sku}</strong></div>', unsafe_allow_html=True)
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
    
    # FINANCIAL SUMMARY
    st.markdown("---")
    st.subheader("💰 Financial Summary")
    
    po_value_main = main_metrics['total_po'] * main_metrics['purchase_price']
    sales_value_main = main_metrics['total_sales'] * main_metrics['floor_price']
    inbound_value_main = main_metrics['total_po_delivered'] * main_metrics['purchase_price']
    gap_main = po_value_main - sales_value_main
    
    col_fin1, col_fin2, col_fin3, col_fin4 = st.columns(4)
    
    with col_fin1:
        st.metric(f"📦 {main_sku} - Total PO Value", format_rupiah(po_value_main),
                  help=f"PO Qty × Purchase Price ({format_rupiah(main_metrics['purchase_price'])}/unit)")
    
    with col_fin2:
        st.metric(f"📥 {main_sku} - Inbound Value", format_rupiah(inbound_value_main),
                  help=f"Inbound Qty × Purchase Price")
    
    with col_fin3:
        st.metric(f"💰 {main_sku} - Total Sales Value", format_rupiah(sales_value_main),
                  help=f"Sales Qty × Floor Price ({format_rupiah(main_metrics['floor_price'])}/unit)")
    
    with col_fin4:
        delta_color = "normal" if gap_main >= 0 else "inverse"
        st.metric(f"⚖️ {main_sku} - Gap", format_rupiah(gap_main),
                  delta=f"{gap_main/po_value_main*100:.1f}% dari PO" if po_value_main > 0 else None,
                  delta_color=delta_color)
    
    if compare_sku and compare_metrics:
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
    
    if main_metrics['floor_price'] == 0 or main_metrics['purchase_price'] == 0:
        st.warning("⚠️ Harga (Floor_Price atau Purchase_Order_Price) = 0. Periksa data Product Master.")
    
    # DETAIL DATA PER BULAN
    with st.expander("📋 Lihat Detail Data per Bulan", expanded=False):
        if not main_df.empty:
            detail_df = main_df.copy()
            detail_df['Bulan'] = detail_df['Month'].dt.strftime('%b %Y')
            detail_df['Sales Qty'] = detail_df['Sales'].apply(lambda x: f"{x:,.0f}")
            detail_df['PO Qty'] = detail_df['PO'].apply(lambda x: f"{x:,.0f}")
            detail_df['Inbound Qty'] = detail_df['PO_Delivered'].apply(lambda x: f"{x:,.0f}")
            detail_df['Sales Value'] = (detail_df['Sales'] * main_metrics['floor_price']).apply(format_rupiah)
            detail_df['PO Value'] = (detail_df['PO'] * main_metrics['purchase_price']).apply(format_rupiah)
            detail_df['Inbound Value'] = (detail_df['PO_Delivered'] * main_metrics['purchase_price']).apply(format_rupiah)
            
            display_cols = ['Bulan', 'Sales Qty', 'Sales Value', 'PO Qty', 'PO Value', 'Inbound Qty', 'Inbound Value']
            st.dataframe(detail_df[display_cols], use_container_width=True, hide_index=True)
            
            if compare_sku and compare_df is not None and not compare_df.empty:
                st.markdown(f'<div class="small-text" style="margin-top:12px;">📊 <strong>Data {compare_sku}</strong></div>', unsafe_allow_html=True)
                compare_detail = compare_df.copy()
                compare_detail['Bulan'] = compare_detail['Month'].dt.strftime('%b %Y')
                compare_detail['Sales Qty'] = compare_detail['Sales'].apply(lambda x: f"{x:,.0f}")
                compare_detail['PO Qty'] = compare_detail['PO'].apply(lambda x: f"{x:,.0f}")
                compare_detail['Inbound Qty'] = compare_detail['PO_Delivered'].apply(lambda x: f"{x:,.0f}")
                st.dataframe(compare_detail[['Bulan', 'Sales Qty', 'PO Qty', 'Inbound Qty']], use_container_width=True, hide_index=True)
        else:
            st.info("Tidak ada data")

# =============================================================================
# TAB 2: SALES ANALYTICS PRO
# =============================================================================
with tab_sales_analytics:
    st.subheader("📊 Sales Analytics Pro")
    st.caption("Deep Dive Analysis: Brand, Tier, Status, Growth Matrix, Seasonality & Pareto")
    
    if df_sales_analysis.empty:
        st.warning("⚠️ Tidak ada data sales untuk dianalisis.")
    else:
        available_years = sorted(df_sales_analysis['Year'].unique())
        selected_years = st.multiselect("📅 Filter Tahun", available_years, default=available_years, key="sales_analytics_year")
        
        df_filtered = df_sales_analysis[df_sales_analysis['Year'].isin(selected_years)] if selected_years else df_sales_analysis
        
        if df_filtered.empty:
            st.warning("Tidak ada data untuk periode yang dipilih.")
        else:
            # Executive Summary
            st.markdown("### 🎯 Executive Summary")
            
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
            
            # Top Brands
            st.markdown("---")
            st.markdown("### 🏆 Top Brands Performance")
            
            col_b1, col_b2 = st.columns(2)
            
            with col_b1:
                top_brands_qty = get_top_brands(df_filtered, 'Sales_Qty', 10)
                fig_top_brands_qty = px.bar(
                    top_brands_qty, x='Brand', y='Total_Sales_Qty',
                    title='Top 10 Brands by Sales Quantity',
                    text=top_brands_qty['Total_Sales_Qty'].apply(lambda x: f"{x:,.0f}"),
                    color='Total_Sales_Qty',
                    color_continuous_scale='Greens'
                )
                fig_top_brands_qty.update_traces(textposition='outside')
                fig_top_brands_qty.update_layout(height=400, xaxis_tickangle=-45)
                st.plotly_chart(fig_top_brands_qty, use_container_width=True)
            
            with col_b2:
                top_brands_value = get_top_brands(df_filtered, 'Sales_Value', 10)
                fig_top_brands_value = px.bar(
                    top_brands_value, x='Brand', y='Total_Sales_Value',
                    title='Top 10 Brands by Sales Value',
                    text=top_brands_value['Total_Sales_Value'].apply(lambda x: format_rupiah(x)),
                    color='Total_Sales_Value',
                    color_continuous_scale='Blues'
                )
                fig_top_brands_value.update_traces(textposition='outside')
                fig_top_brands_value.update_layout(height=400, xaxis_tickangle=-45)
                st.plotly_chart(fig_top_brands_value, use_container_width=True)
            
            # Market Share
            st.markdown("#### 📊 Market Share Distribution")
            col_ms1, col_ms2 = st.columns(2)
            
            with col_ms1:
                market_share_qty = get_top_brands(df_filtered, 'Sales_Qty', 8)
                fig_donut_qty = px.pie(
                    market_share_qty, values='Total_Sales_Qty', names='Brand',
                    title='Market Share by Quantity',
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                fig_donut_qty.update_traces(textposition='inside', textinfo='percent+label')
                fig_donut_qty.update_layout(height=400)
                st.plotly_chart(fig_donut_qty, use_container_width=True)
            
            with col_ms2:
                market_share_value = get_top_brands(df_filtered, 'Sales_Value', 8)
                fig_donut_value = px.pie(
                    market_share_value, values='Total_Sales_Value', names='Brand',
                    title='Market Share by Value',
                    hole=0.4,
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig_donut_value.update_traces(textposition='inside', textinfo='percent+label')
                fig_donut_value.update_layout(height=400)
                st.plotly_chart(fig_donut_value, use_container_width=True)
            
            # Tier & Status
            st.markdown("---")
            st.markdown("### 💎 SKU Tier & Status Analysis")
            
            col_t1, col_t2 = st.columns(2)
            
            with col_t1:
                tier_performance = get_tier_performance(df_filtered)
                if not tier_performance.empty:
                    fig_tier = px.bar(
                        tier_performance, x='SKU_Tier', y='Total_Sales_Qty',
                        title='Sales Quantity by SKU Tier',
                        text=tier_performance['Total_Sales_Qty'].apply(lambda x: f"{x:,.0f}"),
                        color='Total_Sales_Qty',
                        color_continuous_scale='Oranges'
                    )
                    fig_tier.update_traces(textposition='outside')
                    fig_tier.update_layout(height=350)
                    st.plotly_chart(fig_tier, use_container_width=True)
                    
                    st.dataframe(
                        tier_performance[['SKU_Tier', 'SKU_Count', 'Total_Sales_Qty', 'Avg_Qty_per_SKU', 'Share_Qty']],
                        column_config={
                            'Share_Qty': st.column_config.NumberColumn('Share %', format='%.1f%%'),
                            'Avg_Qty_per_SKU': st.column_config.NumberColumn('Avg Qty/SKU', format='%.0f')
                        },
                        use_container_width=True,
                        hide_index=True
                    )
            
            with col_t2:
                status_performance = get_status_performance(df_filtered)
                if not status_performance.empty:
                    fig_status = px.pie(
                        status_performance, values='Total_Sales_Qty', names='Status',
                        title='Sales Contribution by Status',
                        hole=0.4,
                        color_discrete_sequence=['#10B981', '#EF4444', '#9CA3AF']
                    )
                    fig_status.update_traces(textposition='inside', textinfo='percent+label')
                    fig_status.update_layout(height=350)
                    st.plotly_chart(fig_status, use_container_width=True)
                    
                    st.dataframe(status_performance, use_container_width=True, hide_index=True)
            
            # Growth Matrix
            st.markdown("---")
            st.markdown("### 📈 Growth Matrix (Brand Performance)")
            
            brand_growth = get_brand_growth_matrix(df_filtered)
            if not brand_growth.empty:
                fig_growth = px.scatter(
                    brand_growth, x='Market_Share', y='Growth_2026',
                    size='Total_Sales', color='Category',
                    text='Brand',
                    title='Brand Portfolio Matrix: Growth vs Market Share',
                    labels={'Market_Share': 'Market Share (%)', 'Growth_2026': 'YoY Growth 2026 (%)'},
                    color_discrete_map={
                        '🌟 Star (High Growth, High Share)': '#10B981',
                        '🚀 Question Mark (High Growth, Low Share)': '#F59E0B',
                        '💰 Cash Cow (Low Growth, High Share)': '#3B82F6',
                        '🐕 Dog (Low Growth, Low Share)': '#9CA3AF'
                    }
                )
                fig_growth.update_traces(textposition='top center')
                fig_growth.add_hline(y=10, line_dash="dash", line_color="gray", annotation_text="Growth Threshold")
                fig_growth.add_vline(x=10, line_dash="dash", line_color="gray", annotation_text="Share Threshold")
                fig_growth.update_layout(height=500, xaxis_range=[0, brand_growth['Market_Share'].max() * 1.1])
                st.plotly_chart(fig_growth, use_container_width=True)
                
                st.info("💡 **Insight:** Star brands are your growth engines. Question Marks need investment. Cash Cows fund operations. Dogs need evaluation.")
            else:
                st.info("Data tidak cukup untuk analisis growth (butuh data 2025 & 2026).")
            
            # Seasonality
            st.markdown("---")
            st.markdown("### 🌙 Seasonality & Monthly Pattern")
            
            col_szn1, col_szn2 = st.columns(2)
            
            with col_szn1:
                monthly_trend = df_filtered.groupby(df_filtered['Month'].dt.to_period('M'))['Sales_Qty'].sum().reset_index()
                monthly_trend['Month'] = monthly_trend['Month'].dt.to_timestamp()
                monthly_trend['Month_Label'] = monthly_trend['Month'].dt.strftime('%b %Y')
                
                fig_monthly = px.line(
                    monthly_trend, x='Month_Label', y='Sales_Qty',
                    title='Monthly Sales Trend',
                    markers=True,
                    line_shape='spline'
                )
                fig_monthly.update_traces(line=dict(color='#6366F1', width=2.5), marker=dict(size=6))
                fig_monthly.update_layout(height=350, xaxis_tickangle=-45)
                st.plotly_chart(fig_monthly, use_container_width=True)
            
            with col_szn2:
                seasonality = get_seasonality_pattern(df_filtered)
                if not seasonality.empty:
                    fig_season = px.bar(
                        seasonality, x='Month_Name', y='Seasonal_Index',
                        title='Seasonal Index (Average Sales by Month)',
                        text=seasonality['Seasonal_Index'].apply(lambda x: f"{x:.2f}x"),
                        color='Seasonal_Index',
                        color_continuous_scale='RdBu',
                        range_color=[0.5, 1.5]
                    )
                    fig_season.add_hline(y=1, line_dash="dash", line_color="gray", annotation_text="Average")
                    fig_season.update_traces(textposition='outside')
                    fig_season.update_layout(height=350)
                    st.plotly_chart(fig_season, use_container_width=True)
                    
                    peak_month = seasonality.loc[seasonality['Seasonal_Index'].idxmax(), 'Month_Name']
                    low_month = seasonality.loc[seasonality['Seasonal_Index'].idxmin(), 'Month_Name']
                    st.caption(f"📊 **Peak Season:** {peak_month} | **Low Season:** {low_month}")
            
            # Pareto Analysis
            st.markdown("---")
            st.markdown("### 🎯 Pareto Analysis (80/20 Rule)")
            
            pareto_80, all_skus_pareto = calculate_pareto(df_filtered, 80)
            
            col_p1, col_p2 = st.columns(2)
            
            with col_p1:
                fig_pareto = go.Figure()
                
                top_20 = all_skus_pareto.head(20)
                fig_pareto.add_trace(go.Bar(
                    x=top_20['SKU_ID'].astype(str),
                    y=top_20['Sales_Qty'],
                    name='Sales Qty',
                    marker_color='#10B981',
                    text=top_20['Sales_Qty'].apply(lambda x: f"{x:,.0f}"),
                    textposition='outside'
                ))
                
                fig_pareto.add_trace(go.Scatter(
                    x=top_20['SKU_ID'].astype(str),
                    y=top_20['Cumulative_Percent'],
                    name='Cumulative %',
                    mode='lines+markers',
                    line=dict(color='#F59E0B', width=2),
                    yaxis='y2',
                    text=top_20['Cumulative_Percent'].apply(lambda x: f"{x:.1f}%"),
                    textposition='top center'
                ))
                
                fig_pareto.update_layout(
                    title='Top 20 SKUs by Sales (Pareto)',
                    xaxis_title='SKU ID',
                    yaxis_title='Sales Quantity',
                    yaxis2=dict(title='Cumulative %', overlaying='y', side='right', range=[0, 100]),
                    height=450,
                    xaxis_tickangle=-45
                )
                st.plotly_chart(fig_pareto, use_container_width=True)
            
            with col_p2:
                st.markdown(f"""
                <div class="insight-card">
                    <div class="insight-title">📊 PARETO INSIGHT</div>
                    <div class="insight-value">{len(pareto_80)} SKU</div>
                    <div class="insight-desc">meng贡献 {pareto_80['Cumulative_Percent'].iloc[-1]:.1f}% dari total sales</div>
                    <hr style="margin: 10px 0; opacity:0.3;">
                    <div class="insight-title">🎯 RECOMMENDATION</div>
                    <div class="insight-desc">
                        • Fokus pada {len(pareto_80)} SKU ini untuk optimalisasi stok<br>
                        • Evaluasi SKU di luar pareto untuk potensi discontinuasi<br>
                        • Alokasikan budget marketing ke SKU high-performer
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                with st.expander("📋 Lihat Daftar SKU Pareto (80%)"):
                    display_cols = []
                    if 'SKU_ID' in pareto_80.columns:
                        display_cols.append('SKU_ID')
                    if 'Product_Name' in pareto_80.columns:
                        display_cols.append('Product_Name')
                    if 'Brand' in pareto_80.columns:
                        display_cols.append('Brand')
                    if 'SKU_Tier' in pareto_80.columns:
                        display_cols.append('SKU_Tier')
                    display_cols.append('Sales_Qty')
                    display_cols.append('Cumulative_Percent')
                    
                    st.dataframe(
                        pareto_80[display_cols],
                        column_config={
                            'Cumulative_Percent': st.column_config.NumberColumn('Cumulative %', format='%.1f%%')
                        },
                        use_container_width=True,
                        hide_index=True
                    )
            
            # MoM Growth
            st.markdown("---")
            st.markdown("### 📉 Month-over-Month Growth Analysis")
            
            growth_metrics = calculate_growth_metrics(df_filtered)
            if not growth_metrics.empty:
                fig_mom = go.Figure()
                
                fig_mom.add_trace(go.Bar(
                    x=growth_metrics['Month'].dt.strftime('%b %Y'),
                    y=growth_metrics['MoM_Growth_Qty'],
                    name='MoM Growth %',
                    marker_color=np.where(growth_metrics['MoM_Growth_Qty'] >= 0, '#10B981', '#EF4444'),
                    text=growth_metrics['MoM_Growth_Qty'].apply(lambda x: f"{x:+.1f}%"),
                    textposition='outside'
                ))
                
                fig_mom.add_hline(y=0, line_dash="solid", line_color="gray")
                fig_mom.update_layout(
                    title='Month-over-Month Sales Growth',
                    xaxis_title='Month',
                    yaxis_title='Growth (%)',
                    height=400,
                    xaxis_tickangle=-45
                )
                st.plotly_chart(fig_mom, use_container_width=True)
                
                avg_growth = growth_metrics['MoM_Growth_Qty'].mean()
                st.caption(f"📈 **Average MoM Growth:** {avg_growth:+.1f}%")
            
            # Data Explorer
            st.markdown("---")
            with st.expander("📋 Data Explorer - Raw Sales Data", expanded=False):
                st.dataframe(df_filtered, use_container_width=True, height=400)

# =============================================================================
# TAB 3: STOCK ANALYSIS
# =============================================================================
with tab_stock:
    st.subheader("📦 Stock Onhand Analysis")
    st.caption("Analisis stok multi-batch: Expiry Date monitoring & Stock Health")
    
    if df_stock.empty:
        st.warning("⚠️ Tidak ada data stok yang tersedia.")
    else:
        all_stock_skus = sorted(df_stock['SKU_ID'].dropna().unique())
        
        col_filter1, col_filter2 = st.columns([2, 1])
        
        with col_filter1:
            selected_stock_sku = st.selectbox("🔍 Pilih SKU untuk Analisis Stok", all_stock_skus, key="stock_sku")
        
        with col_filter2:
            if st.button("🔄 Refresh Stock Data", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
        
        product_info = df_product[df_product['SKU_ID'] == selected_stock_sku]
        product_name = product_info.iloc[0].get('Product_Name', selected_stock_sku) if not product_info.empty else selected_stock_sku
        brand = product_info.iloc[0].get('Brand', '-') if not product_info.empty else '-'
        
        st.markdown(f"""
        <div class="sku-header" style="background: linear-gradient(135deg, #3B82F6 0%, #1D4ED8 100%);">
            <div>
                <div class="sku-title">{product_name} <span style="font-size:0.8rem;">({selected_stock_sku})</span></div>
                <div class="sku-badges">
                    <span class="badge">🏷️ {brand}</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        stock_metrics = calculate_stock_metrics(df_stock, df_product, selected_stock_sku, 0)
        
        if stock_metrics['has_stock']:
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Stock", f"{stock_metrics['total_stock']:,.0f}", 
                          delta=f"{stock_metrics['batch_count']} batch")
            
            with col2:
                st.metric("Fresh Stock (>6 bln)", f"{stock_metrics['fresh_stock']:,.0f}")
            
            with col3:
                st.metric("Expiring ≤6 bln", f"{stock_metrics['expiring_6months'] + stock_metrics['expiring_3months'] + stock_metrics['expiring_soon']:,.0f}")
            
            with col4:
                expired_total = stock_metrics['expired_stock']
                st.metric("Expired Stock", f"{expired_total:,.0f}", 
                          delta="⚠️ PERLU DISPOSISI" if expired_total > 0 else None,
                          delta_color="inverse" if expired_total > 0 else "normal")
            
            st.markdown("---")
            st.subheader("📊 Expiry Date Distribution")
            
            expiry_data = []
            if stock_metrics['expired_stock'] > 0:
                expiry_data.append({'Category': 'Expired (< 0 days)', 'Qty': stock_metrics['expired_stock']})
            if stock_metrics['expiring_soon'] > 0:
                expiry_data.append({'Category': 'Expiring Soon (0-30 days)', 'Qty': stock_metrics['expiring_soon']})
            if stock_metrics['expiring_3months'] > 0:
                expiry_data.append({'Category': 'Expiring in 1-3 months', 'Qty': stock_metrics['expiring_3months']})
            if stock_metrics['expiring_6months'] > 0:
                expiry_data.append({'Category': 'Expiring in 3-6 months', 'Qty': stock_metrics['expiring_6months']})
            if stock_metrics['fresh_stock'] > 0:
                expiry_data.append({'Category': 'Fresh (> 6 months)', 'Qty': stock_metrics['fresh_stock']})
            
            if expiry_data:
                expiry_df = pd.DataFrame(expiry_data)
                color_map = {
                    'Expired (< 0 days)': '#EF4444',
                    'Expiring Soon (0-30 days)': '#F59E0B',
                    'Expiring in 1-3 months': '#FBBF24',
                    'Expiring in 3-6 months': '#60A5FA',
                    'Fresh (> 6 months)': '#10B981'
                }
                fig_expiry = px.pie(
                    expiry_df, values='Qty', names='Category',
                    title='Stock Distribution by Expiry Status',
                    hole=0.4,
                    color='Category',
                    color_discrete_map=color_map
                )
                fig_expiry.update_traces(textposition='inside', textinfo='percent+label')
                st.plotly_chart(fig_expiry, use_container_width=True)
            
            st.markdown("---")
            st.subheader("📋 Batch Details")
            
            batch_df = stock_metrics['batch_details'].copy()
            batch_df['Expiry_Date'] = batch_df['Expiry_Date'].dt.strftime('%d %b %Y')
            batch_df = batch_df.rename(columns={
                'Batch_Number': 'Batch Number',
                'Physical_Stock': 'Stock Qty',
                'Expiry_Date': 'Expiry Date'
            })
            
            today = datetime.now().date()
            def get_expiry_warning(expiry_str):
                if expiry_str == 'N/A' or pd.isna(expiry_str):
                    return "⚪"
                try:
                    expiry_date = datetime.strptime(expiry_str, '%d %b %Y').date()
                    days = (expiry_date - today).days
                    if days < 0:
                        return "🔴 EXPIRED"
                    elif days <= 30:
                        return "🟠 CRITICAL"
                    elif days <= 90:
                        return "🟡 WARNING"
                    else:
                        return "🟢 OK"
                except:
                    return "⚪"
            
            batch_df['Status'] = batch_df['Expiry Date'].apply(get_expiry_warning)
            batch_df = batch_df[['Batch Number', 'Stock Qty', 'Expiry Date', 'Status']]
            
            st.dataframe(batch_df, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            if stock_metrics['expired_stock'] > 0:
                st.error(f"""
                🚨 **URGENT - EXPIRED STOCK DETECTED!**
                - Total expired: {stock_metrics['expired_stock']:,.0f} unit
                - **Action Required:** Segera lakukan disposisi dan pisahkan dari stok baik.
                """)
            
            if stock_metrics['expiring_soon'] > 0:
                st.warning(f"""
                ⚠️ **EXPIRING SOON (< 30 DAYS)**
                - Total akan expired: {stock_metrics['expiring_soon']:,.0f} unit
                - **Recommendation:** Prioritaskan penjualan atau diskon untuk batch ini.
                """)
            
            if stock_metrics['expiring_3months'] > 0:
                st.info(f"""
                ℹ️ **EXPIRING IN 1-3 MONTHS**
                - Total akan expired: {stock_metrics['expiring_3months']:,.0f} unit
                - **Recommendation:** Mulai rencanakan promo atau bundling untuk mempercepat pergerakan stok.
                """)
                
        else:
            st.info(f"📦 Tidak ada data stok untuk SKU {selected_stock_sku}")

# =============================================================================
# FOOTER
# =============================================================================
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #888; font-size: 0.7rem; padding: 0.5rem;">
    <p>📊 SKU 360° Evaluator Pro | Data Sales 2025-2026 | Data PO & Inbound hingga Mar 2026 | Multi-batch Stock Management | Sales Analytics Pro</p>
</div>
""", unsafe_allow_html=True)
