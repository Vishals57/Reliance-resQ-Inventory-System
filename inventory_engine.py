import pandas as pd
import os
from datetime import datetime

DB_FILE = "resQ_Enterprise_Inventory.xlsx"
ENGINEERS = ["Rahul Sharma", "Amit Patil", "Vinay Deshmukh", "Suresh Kumar"]

def initialize_db():
    """Creates the file with both sheets if it doesn't exist."""
    if not os.path.exists(DB_FILE):
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            pd.DataFrame(columns=["Article_No", "Part_Name", "CP", "SP", "Category", "Stock_Level"]).to_excel(writer, sheet_name='Master', index=False)
            pd.DataFrame(columns=["Article_No", "Purchase_Type", "Status", "Engineer", "Charges", "Date"]).to_excel(writer, sheet_name='Transactions', index=False)
        print("✅ resQ Enterprise Database Initialized.")

def register_new_article(art_no, name, cp, sp, cat="OG"):
    # Load all existing data first to prevent losing sheets
    df_master = pd.read_excel(DB_FILE, sheet_name='Master', dtype={'Article_No': str})
    try:
        df_trans = pd.read_excel(DB_FILE, sheet_name='Transactions', dtype={'Article_No': str})
    except:
        df_trans = pd.DataFrame(columns=["Article_No", "Purchase_Type", "Status", "Engineer", "Charges", "Date"])

    art_no = str(art_no).strip()
    if art_no in df_master['Article_No'].astype(str).str.strip().values:
        return False, f"❌ Article '{art_no}' already exists!"
    
    new_entry = pd.DataFrame([{
        "Article_No": art_no, "Part_Name": name, "CP": cp, "SP": sp, "Category": cat, "Stock_Level": 0
    }])
    
    df_master = pd.concat([df_master, new_entry], ignore_index=True)
    
    # SAVE BOTH SHEETS to ensure 'Transactions' doesn't disappear
    with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
        df_master.to_excel(writer, sheet_name='Master', index=False)
        df_trans.to_excel(writer, sheet_name='Transactions', index=False)
        
    return True, f"✅ Article {art_no} registered."

def process_movement(art_no, movement_type, purchase_type="Company", engineer=None):
    try:
        df_master = pd.read_excel(DB_FILE, sheet_name='Master', dtype={'Article_No': str})
        df_trans = pd.read_excel(DB_FILE, sheet_name='Transactions', dtype={'Article_No': str})
    except:
        return False, "❌ Error reading database."

    search_id = str(art_no).strip()
    df_master['Article_No'] = df_master['Article_No'].astype(str).str.strip()
    df_trans['Article_No'] = df_trans['Article_No'].astype(str).str.strip()

    if search_id not in df_master['Article_No'].values:
        return False, f"❌ Article '{search_id}' not found in Master!"

    master_idx = df_master.index[df_master['Article_No'] == search_id][0]
    product_name = df_master.at[master_idx, 'Part_Name']
    
    # 1. SEARCH: Does this specific ID already have a record?
    existing_row_mask = (df_trans['Article_No'] == search_id)

    if movement_type.upper() == "IN":
        # --- THE DUPLICATE CHECK ---
        # If the row exists AND the status is already 'IN', stop the user!
        if existing_row_mask.any():
            current_status = df_trans.loc[existing_row_mask, 'Status'].values[0]
            if current_status == "IN":
                return False, f"⚠️ BLOCK: Article {search_id} is ALREADY IN store. You cannot add it twice!"
            
            # If it exists but was 'OUT', we REFILL/RESET it
            target_idx = df_trans[existing_row_mask].index[0]
            df_trans.at[target_idx, 'Status'] = "IN"
            df_trans.at[target_idx, 'Purchase_Type'] = purchase_type # Allow switching type on refill
            df_trans.at[target_idx, 'Engineer'] = "Store"
            df_trans.at[target_idx, 'Charges'] = 0
            df_trans.at[target_idx, 'Date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            msg = f"✅ RE-ENTRY: {product_name} ({search_id}) returned to store."
        else:
            # BRAND NEW ID: Create the first and only row for this ID
            new_log = pd.DataFrame([{
                "Article_No": search_id,
                "Purchase_Type": purchase_type,
                "Status": "IN",
                "Engineer": "Store",
                "Charges": 0,
                "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }])
            df_trans = pd.concat([df_trans, new_log], ignore_index=True)
            msg = f"🆕 REGISTERED: Unique ID {search_id} added to shelf."
        
        # Update Master stock (Max will always be 1 for this specific ID)
        df_master.at[master_idx, 'Stock_Level'] = 1

    elif movement_type.upper() == "OUT":
        # Only allow OUT if the current status is IN
        active_stock_mask = existing_row_mask & (df_trans['Status'] == "IN")
        
        if active_stock_mask.any():
            target_idx = df_trans[active_stock_mask].index[0]
            df_trans.at[target_idx, 'Status'] = "OUT"
            df_trans.at[target_idx, 'Engineer'] = engineer
            df_trans.at[target_idx, 'Charges'] = df_master.at[master_idx, 'SP']
            df_trans.at[target_idx, 'Date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            df_master.at[master_idx, 'Stock_Level'] = 0
            msg = f"📤 ISSUED: Unique ID {search_id} given to {engineer}."
        else:
            return False, f"❌ ERROR: ID {search_id} is not in store (It might be already OUT)."

    # 4. Save both sheets
    with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
        df_master.to_excel(writer, sheet_name='Master', index=False)
        df_trans.to_excel(writer, sheet_name='Transactions', index=False)
    
    return True, msg