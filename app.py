import io
import re
import pandas as pd
import pypdf
import streamlit as st
from supabase import create_client, Client

st.set_page_config(page_title="Trimize Operations Portal", layout="wide")

# --- SUPABASE SETUP ---
@st.cache_resource
def init_supabase():
    """Safely initialize Supabase client using secrets."""
    try:
        url = st.secrets["SUPABASE_URL"].strip().rstrip("/")
        key = st.secrets["SUPABASE_KEY"].strip()
        return create_client(url, key)
    except Exception as e:
        return None

supabase = init_supabase()

# --- HELPER FUNCTIONS ---
def get_bundle_recipes():
    """Fetch bundle recipes from Supabase, or fall back to default hardcoded recipes if DB fails."""
    if supabase:
        try:
            response = supabase.table("bundle_recipes").select("*").execute()
            if response.data:
                recipes = {}
                for row in response.data:
                    bundle = row["bundle_name"]
                    if bundle not in recipes:
                        recipes[bundle] = []
                    recipes[bundle].append({
                        "component": row["component_name"],
                        "qty": row["quantity"]
                    })
                return recipes
        except Exception as e:
            st.warning("⚠️ Could not load recipes from database. Using default fallback recipes.")

    # Fallback default dictionary
    return {
        'The Ultimate Grooming Bundle - أسود': [
            {'component': 'ماكينة حلاقة تريمايز للرجال - أسود', 'qty': 1},
            {'component': 'تريمايز جل الاستحمام', 'qty': 1},
            {'component': 'مزيل رائحة العرق من تريمايز', 'qty': 1},
            {'component': 'تريمايز غسول المناطق الحساسة', 'qty': 1}
        ],
        'The Ultimate Grooming Bundle - أزرق': [
            {'component': 'ماكينة حلاقة تريمايز للرجال - أزرق', 'qty': 1},
            {'component': 'تريمايز جل الاستحمام', 'qty': 1},
            {'component': 'مزيل رائحة العرق من تريمايز', 'qty': 1},
            {'component': 'تريمايز غسول المناطق الحساسة', 'qty': 1}
        ],
        'The Ultimate Grooming Bundle - أخضر': [
            {'component': 'ماكينة حلاقة تريمايز للرجال - أخضر', 'qty': 1},
            {'component': 'تريمايز جل الاستحمام', 'qty': 1},
            {'component': 'مزيل رائحة العرق من تريمايز', 'qty': 1},
            {'component': 'تريمايز غسول المناطق الحساسة', 'qty': 1}
        ],
        'The Full Routine Bundle': [
            {'component': 'تريمايز جل الاستحمام', 'qty': 1},
            {'component': 'مزيل رائحة العرق من تريمايز', 'qty': 1},
            {'component': 'تريمايز غسول المناطق الحساسة', 'qty': 1}
        ],
        'Intimate Bundle': [
            {'component': 'مزيل رائحة العرق من تريمايز', 'qty': 1},
            {'component': 'تريمايز غسول المناطق الحساسة', 'qty': 1}
        ]
    }

def explode_shopify_orders(df, recipes):
    """Explode bundle items into individual SKU rows while keeping original metadata."""
    exploded_rows = []
    
    for idx, row in df.iterrows():
        item_name = str(row.get("Lineitem name", "")).strip()
        quantity = int(row.get("Lineitem quantity", 1))
        
        # Check if item is a registered bundle
        if item_name in recipes:
            components = recipes[item_name]
            first = True
            for comp in components:
                new_row = row.to_dict()
                new_row["Lineitem name"] = comp["component"]
                new_row["Lineitem quantity"] = comp["qty"] * quantity
                
                # Blank out financial totals on secondary component rows to avoid double counting
                if not first:
                    if "Subtotal" in new_row:
                        new_row["Subtotal"] = 0
                    if "Total" in new_row:
                        new_row["Total"] = 0
                first = False
                exploded_rows.append(new_row)
        else:
            exploded_rows.append(row.to_dict())
            
    return pd.DataFrame(exploded_rows)

def parse_bosta_pdf(pdf_file):
    """Exact parsing logic from Tag revise.txt."""
    reader = pypdf.PdfReader(pdf_file)
    records = []
    
    for idx, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        
        # 1. Order Reference
        order_ref = None
        ref_match = re.search(r'Order Reference:\s*(\S+)', text)
        if ref_match:
            order_ref = ref_match.group(1)

        # 2. Tracking Number
        tracking_num = None
        track_match = re.search(r'Tracking Number\s*(\d+)', text)
        if track_match:
            tracking_num = track_match.group(1)

        # 3. COD / Collection Amount
        cod_amount = "0"
        cod_match = re.search(r'مبلغ التحصيل:\s*([\d,]+(?:\.\d+)?)', text)
        if not cod_match:
            cod_match = re.search(r'ج\.م\s*([\d,]+(?:\.\d+)?)', text)
        if cod_match:
            cod_amount = cod_match.group(1).replace(',', '')

        records.append({
            "Page": idx,
            "Order Reference": order_ref,
            "Tracking Number": tracking_num,
            "COD Amount": cod_amount
        })

    return pd.DataFrame(records)

def extract_order_num(val):
    """Extract digits to match '#120166', '120166', or '#120166.1' seamlessly."""
    if pd.isna(val):
        return ""
    match = re.search(r'\d+', str(val))
    return match.group(0) if match else ""

# --- NAVIGATION & INTERFACE ---
st.title("📦 Trimize Operations Portal")

tab1, tab2, tab3 = st.tabs(["1. Process Orders & Tags", "2. Warehouse Log", "3. Inventory Overview"])

# TAB 1: PROCESS SHOPIFY & TAGS
with tab1:
    st.subheader("Process Shopify Export & Carrier PDF")
    
    col1, col2 = st.columns(2)
    
    with col1:
        shopify_file = st.file_uploader("Upload Shopify Export (XLSX)", type=["xlsx"])
    with col2:
        tag_file = st.file_uploader("Upload Bosta Airway Bills (PDF)", type=["pdf"])
        
    if shopify_file and tag_file:
        if st.button("🚀 Explode Bundles & Merge Airway Bills"):
            try:
                # 1. Read and explode Shopify Export
                raw_shopify = pd.read_excel(shopify_file)
                recipes = get_bundle_recipes()
                exploded_df = explode_shopify_orders(raw_shopify, recipes)
                
                # 2. Extract PDF Tag data
                pdf_df = parse_bosta_pdf(tag_file)
                
                # 3. Check for missing columns
                if "Name" not in exploded_df.columns:
                    st.error("Error: Could not find 'Name' column in Shopify Excel file.")
                elif "Order Reference" not in pdf_df.columns:
                    st.error("Error: Could not extract 'Order Reference' from PDF airway bills.")
                else:
                    # 4. Normalize join keys by extracting digits (e.g. #120166.1 -> 120166)
                    exploded_df["join_key"] = exploded_df["Name"].apply(extract_order_num)
                    pdf_df["join_key"] = pdf_df["Order Reference"].apply(extract_order_num)
                    
                    # 5. Merge Shopify data with PDF Tracking Number
                    merged_df = pd.merge(
                        exploded_df, 
                        pdf_df, 
                        on="join_key", 
                        how="left"
                    ).drop(columns=["join_key"])
                    
                    st.success("Successfully processed and merged order records!")
                    st.dataframe(merged_df, use_container_width=True)
                    
                    # Download merged CSV
                    csv_buffer = io.BytesIO()
                    merged_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="📥 Download Merged CSV",
                        data=csv_buffer.getvalue(),
                        file_name="processed_shipments.csv",
                        mime="text/csv"
                    )
            except Exception as err:
                st.error(f"Processing Error: {str(err)}")

# TAB 2: WAREHOUSE LOG
with tab2:
    st.subheader("Daily Warehouse Log")
    st.info("Warehouse staff can view required shipments and log dispatched stock directly.")

# TAB 3: INVENTORY LEDGER
with tab3:
    st.subheader("Reconciled Stock Levels")
    st.info("Automated calculation: Beginning Inventory + Restocked Returns - Dispatches")
