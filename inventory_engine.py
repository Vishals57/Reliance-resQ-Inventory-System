import pandas as pd
import os
from datetime import datetime

DB_FILE = "resQ_Enterprise_Inventory.xlsx"
# Backward compatible default list (used only to seed the Engineers sheet once).
ENGINEERS = ["Yogesh Bhosale", "Sidram Battul", "Sham Kachi", "sameer shaikh", "Rajesh chavan", "Sahil Dighe", "Gajanan gawande", "Kiran misal", "kishor panchal", "Raju Yadav", "Lalemashak kunnure", "Faruk Shaikh", "Ashraf Momin", "Abhitabh Gupta", "Prahalad Yadav", "Abdul Rehman", "Ishwar Panchal", "Nagendra Yadav", "Pramod Valekar", "Rahim Dindore"]

ENGINEER_SHEET = "Engineers"
ENGINEER_COLUMNS = ["Engineer_Name", "BP_ID", "PPRR_ID", "Active", "Created_At"]

MASTER_COLUMNS = ["Article_No", "Part_Name", "CP", "SP", "Category", "Stock_Level"]
# `Date` is kept as a generic "last update" timestamp for backward compatibility,
# while `In_Date` and `Out_Date` store movement timestamps explicitly.
TRANSACTION_COLUMNS = [
    "Article_No",
    "Sr_No",
    "Purchase_Type",
    "Tax_Invoice_No",
    "Status",
    "Engineer",
    "Charges",
    "In_Date",
    "Out_Date",
    "Date",
]

def initialize_db():
    """Creates the file with both sheets if it doesn't exist."""
    if not os.path.exists(DB_FILE):
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            pd.DataFrame(columns=MASTER_COLUMNS).to_excel(writer, sheet_name='Master', index=False)
            pd.DataFrame(columns=TRANSACTION_COLUMNS).to_excel(writer, sheet_name='Transactions', index=False)
        print("✅ resQ Enterprise Database Initialized.")

def migrate_db():
    """
    Ensures existing Excel file has the latest schema/columns.
    Safe to call every startup.
    """
    if not os.path.exists(DB_FILE):
        initialize_db()
        return

    try:
        df_master = pd.read_excel(DB_FILE, sheet_name="Master", dtype={"Article_No": str})
    except Exception:
        df_master = pd.DataFrame(columns=MASTER_COLUMNS)

    try:
        df_trans = pd.read_excel(DB_FILE, sheet_name="Transactions", dtype={"Article_No": str})
    except Exception:
        df_trans = pd.DataFrame(columns=TRANSACTION_COLUMNS)

    # Engineers sheet (create if missing)
    try:
        df_eng = pd.read_excel(DB_FILE, sheet_name=ENGINEER_SHEET, dtype=str)
    except Exception:
        seeded = []
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for n in ENGINEERS:
            seeded.append({"Engineer_Name": n, "BP_ID": "", "PPRR_ID": "", "Active": True, "Created_At": now})
        df_eng = pd.DataFrame(seeded, columns=ENGINEER_COLUMNS)

    # Normalize columns and upgrade schema
    if not df_master.empty:
        for col in MASTER_COLUMNS:
            if col not in df_master.columns:
                df_master[col] = None
        df_master = df_master[MASTER_COLUMNS]
    else:
        df_master = pd.DataFrame(columns=MASTER_COLUMNS)

    df_trans = _ensure_schema(df_trans)

    # Normalize engineers schema
    if df_eng is None or df_eng.empty:
        df_eng = pd.DataFrame(columns=ENGINEER_COLUMNS)
    else:
        for col in ENGINEER_COLUMNS:
            if col not in df_eng.columns:
                df_eng[col] = None
        df_eng = df_eng[ENGINEER_COLUMNS]
        df_eng["Engineer_Name"] = df_eng["Engineer_Name"].astype(str).str.strip()
        # Ensure unique by name
        df_eng = df_eng.dropna(subset=["Engineer_Name"])
        df_eng = df_eng[df_eng["Engineer_Name"].astype(str).str.strip() != ""]
        df_eng = df_eng.drop_duplicates(subset=["Engineer_Name"], keep="last")

    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master", index=False)
        df_trans.to_excel(writer, sheet_name="Transactions", index=False)
        df_eng.to_excel(writer, sheet_name=ENGINEER_SHEET, index=False)

def get_engineers(active_only: bool = True):
    """Return engineers as list of dicts: name + ids."""
    migrate_db()
    try:
        df = pd.read_excel(DB_FILE, sheet_name=ENGINEER_SHEET, dtype=str)
    except Exception:
        return []
    if df is None or df.empty:
        return []
    for col in ENGINEER_COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df[ENGINEER_COLUMNS].copy()
    df["Engineer_Name"] = df["Engineer_Name"].astype(str).str.strip()
    df = df[df["Engineer_Name"] != ""]
    if active_only and "Active" in df.columns:
        # Accept True/False or "True"/"False"
        df["Active"] = df["Active"].astype(str).str.lower().isin(["true", "1", "yes", "y"])
        df = df[df["Active"] == True]
    out = []
    for _, r in df.iterrows():
        out.append({
            "Engineer_Name": str(r.get("Engineer_Name", "")).strip(),
            "BP_ID": "" if r.get("BP_ID") is None else str(r.get("BP_ID")).strip(),
            "PPRR_ID": "" if r.get("PPRR_ID") is None else str(r.get("PPRR_ID")).strip(),
        })
    return out

def get_engineer_names(active_only: bool = True):
    return [e["Engineer_Name"] for e in get_engineers(active_only=active_only)]

def add_engineer(name: str, bp_id: str = "", pprr_id: str = ""):
    """Add or update an engineer in Engineers sheet."""
    migrate_db()
    name = "" if name is None else str(name).strip()
    if not name:
        return False, "Engineer name is required."

    bp_id = "" if bp_id is None else str(bp_id).strip()
    pprr_id = "" if pprr_id is None else str(pprr_id).strip()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        df = pd.read_excel(DB_FILE, sheet_name=ENGINEER_SHEET, dtype=str)
    except Exception:
        df = pd.DataFrame(columns=ENGINEER_COLUMNS)

    for col in ENGINEER_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["Engineer_Name"] = df["Engineer_Name"].astype(str).str.strip()
    mask = df["Engineer_Name"].str.lower() == name.lower()
    if mask.any():
        idx = df[mask].index[0]
        df.at[idx, "Engineer_Name"] = name
        df.at[idx, "BP_ID"] = bp_id
        df.at[idx, "PPRR_ID"] = pprr_id
        df.at[idx, "Active"] = True
        msg = f"Updated engineer: {name}"
    else:
        df = pd.concat([df, pd.DataFrame([{
            "Engineer_Name": name,
            "BP_ID": bp_id,
            "PPRR_ID": pprr_id,
            "Active": True,
            "Created_At": now,
        }])], ignore_index=True)
        msg = f"Added engineer: {name}"

    # Keep unique
    df = df.drop_duplicates(subset=["Engineer_Name"], keep="last")

    # Write back all sheets safely
    try:
        df_master = pd.read_excel(DB_FILE, sheet_name="Master", dtype={"Article_No": str})
    except Exception:
        df_master = pd.DataFrame(columns=MASTER_COLUMNS)
    try:
        df_trans = pd.read_excel(DB_FILE, sheet_name="Transactions", dtype={"Article_No": str})
    except Exception:
        df_trans = pd.DataFrame(columns=TRANSACTION_COLUMNS)
    df_trans = _ensure_schema(df_trans)

    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master", index=False)
        df_trans.to_excel(writer, sheet_name="Transactions", index=False)
        df[ENGINEER_COLUMNS].to_excel(writer, sheet_name=ENGINEER_SHEET, index=False)

    return True, msg

def remove_engineer(name: str):
    """Soft-remove engineer by setting Active=False."""
    migrate_db()
    name = "" if name is None else str(name).strip()
    if not name:
        return False, "Engineer name is required."

    try:
        df = pd.read_excel(DB_FILE, sheet_name=ENGINEER_SHEET, dtype=str)
    except Exception:
        return False, "Engineers sheet not found."

    for col in ENGINEER_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df["Engineer_Name"] = df["Engineer_Name"].astype(str).str.strip()
    mask = df["Engineer_Name"].str.lower() == name.lower()
    if not mask.any():
        return False, f"Engineer not found: {name}"

    idx = df[mask].index[0]
    df.at[idx, "Active"] = False

    # Write back all sheets safely
    try:
        df_master = pd.read_excel(DB_FILE, sheet_name="Master", dtype={"Article_No": str})
    except Exception:
        df_master = pd.DataFrame(columns=MASTER_COLUMNS)
    try:
        df_trans = pd.read_excel(DB_FILE, sheet_name="Transactions", dtype={"Article_No": str})
    except Exception:
        df_trans = pd.DataFrame(columns=TRANSACTION_COLUMNS)
    df_trans = _ensure_schema(df_trans)

    with pd.ExcelWriter(DB_FILE, engine="openpyxl") as writer:
        df_master.to_excel(writer, sheet_name="Master", index=False)
        df_trans.to_excel(writer, sheet_name="Transactions", index=False)
        df[ENGINEER_COLUMNS].to_excel(writer, sheet_name=ENGINEER_SHEET, index=False)

    return True, f"Removed engineer: {name}"

def _ensure_schema(df_trans: pd.DataFrame) -> pd.DataFrame:
    """
    Backward-compatible schema upgrade:
    - Old files may not have `Sr_No`. Default existing rows to Sr_No=1.
    - Ensure all expected columns exist.
    """
    if df_trans is None or df_trans.empty:
        return pd.DataFrame(columns=TRANSACTION_COLUMNS)

    df = df_trans.copy()
    if "Sr_No" not in df.columns:
        df.insert(1, "Sr_No", 1)

    for col in TRANSACTION_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # If we are upgrading from older schema that had only `Date`,
    # populate missing In_Date / Out_Date based on current Status.
    if "Date" in df.columns:
        if "In_Date" in df.columns:
            in_missing = df["In_Date"].isna() | (df["In_Date"].astype(str).str.strip() == "") | (df["In_Date"].astype(str) == "nan")
            df.loc[in_missing & (df["Status"] == "IN"), "In_Date"] = df.loc[in_missing & (df["Status"] == "IN"), "Date"]
        if "Out_Date" in df.columns:
            out_missing = df["Out_Date"].isna() | (df["Out_Date"].astype(str).str.strip() == "") | (df["Out_Date"].astype(str) == "nan")
            df.loc[out_missing & (df["Status"] == "OUT"), "Out_Date"] = df.loc[out_missing & (df["Status"] == "OUT"), "Date"]

    return df[TRANSACTION_COLUMNS]

def _normalize_ids(df_master: pd.DataFrame, df_trans: pd.DataFrame):
    df_master["Article_No"] = df_master["Article_No"].astype(str).str.strip()
    df_trans["Article_No"] = df_trans["Article_No"].astype(str).str.strip()
    df_trans["Sr_No"] = pd.to_numeric(df_trans["Sr_No"], errors="coerce").fillna(1).astype(int)
    return df_master, df_trans

def get_available_sr_nos(article_no: str):
    """Returns Sr_No list that are currently IN for an Article_No."""
    try:
        df_trans = pd.read_excel(DB_FILE, sheet_name="Transactions", dtype={"Article_No": str})
    except Exception:
        return []

    df_trans = _ensure_schema(df_trans)
    df_trans["Article_No"] = df_trans["Article_No"].astype(str).str.strip()
    df_trans["Sr_No"] = pd.to_numeric(df_trans["Sr_No"], errors="coerce").fillna(1).astype(int)

    search_id = str(article_no).strip()
    active_mask = (df_trans["Article_No"] == search_id) & (df_trans["Status"] == "IN")
    srs = sorted(df_trans.loc[active_mask, "Sr_No"].astype(int).unique().tolist())
    return srs

def get_scan_details(article_no: str):
    """
    Returns details to show after scanning:
    - Article_No, Part_Name, SP
    - In-units list: [{Sr_No, Tax_Invoice_No}]
    """
    search_id = str(article_no).strip()
    try:
        df_master = pd.read_excel(DB_FILE, sheet_name="Master", dtype={"Article_No": str})
    except Exception:
        return {
            "Article_No": search_id,
            "Part_Name": "",
            "SP": None,
            "in_units": [],
        }

    try:
        df_trans = pd.read_excel(DB_FILE, sheet_name="Transactions", dtype={"Article_No": str})
    except Exception:
        df_trans = pd.DataFrame(columns=TRANSACTION_COLUMNS)

    df_trans = _ensure_schema(df_trans)
    df_master, df_trans = _normalize_ids(df_master, df_trans)

    part_name = ""
    sp = None
    if search_id in df_master["Article_No"].values:
        idx = df_master.index[df_master["Article_No"] == search_id][0]
        part_name = df_master.at[idx, "Part_Name"]
        sp = df_master.at[idx, "SP"]

    active_mask = (df_trans["Article_No"] == search_id) & (df_trans["Status"] == "IN")
    active = df_trans.loc[active_mask, ["Sr_No", "Tax_Invoice_No"]].copy()
    if not active.empty:
        active["Sr_No"] = pd.to_numeric(active["Sr_No"], errors="coerce").fillna(1).astype(int)
        active = active.sort_values("Sr_No")

    in_units = []
    for _, r in active.iterrows():
        inv = r.get("Tax_Invoice_No", "")
        inv = "" if inv is None else str(inv).strip()
        in_units.append({"Sr_No": int(r["Sr_No"]), "Tax_Invoice_No": inv})

    return {
        "Article_No": search_id,
        "Part_Name": part_name,
        "SP": sp,
        "in_units": in_units,
    }

def get_next_inward_sr_list(article_no: str, qty: int):
    """Returns the Sr_No list that will be allocated by auto-inward logic."""
    search_id = str(article_no).strip()
    try:
        qty_int = int(qty)
    except Exception:
        qty_int = 1
    if qty_int < 1:
        qty_int = 1

    try:
        df_trans = pd.read_excel(DB_FILE, sheet_name="Transactions", dtype={"Article_No": str})
    except Exception:
        df_trans = pd.DataFrame(columns=TRANSACTION_COLUMNS)

    df_trans = _ensure_schema(df_trans)
    df_trans["Article_No"] = df_trans["Article_No"].astype(str).str.strip()
    df_trans["Sr_No"] = pd.to_numeric(df_trans["Sr_No"], errors="coerce").fillna(0).astype(int)

    existing_for_article = df_trans[df_trans["Article_No"] == search_id]
    max_sr = int(existing_for_article["Sr_No"].max()) if not existing_for_article.empty else 0
    return list(range(max_sr + 1, max_sr + qty_int + 1))

def register_new_article(art_no, name, cp, sp, cat="OG"):
    # Load all existing data first to prevent losing sheets
    df_master = pd.read_excel(DB_FILE, sheet_name='Master', dtype={'Article_No': str})
    try:
        df_trans = pd.read_excel(DB_FILE, sheet_name='Transactions', dtype={'Article_No': str})
    except:
        df_trans = pd.DataFrame(columns=TRANSACTION_COLUMNS)

    df_trans = _ensure_schema(df_trans)

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

def process_movement(art_no, movement_type, purchase_type="Company", engineer=None, sr_no=None, qty: int = 1, tax_invoice_no: str = ""):
    try:
        df_master = pd.read_excel(DB_FILE, sheet_name='Master', dtype={'Article_No': str})
        df_trans = pd.read_excel(DB_FILE, sheet_name='Transactions', dtype={'Article_No': str})
    except:
        return False, "❌ Error reading database."

    df_trans = _ensure_schema(df_trans)

    search_id = str(art_no).strip()
    df_master, df_trans = _normalize_ids(df_master, df_trans)

    if search_id not in df_master['Article_No'].values:
        return False, f"❌ Article '{search_id}' not found in Master!"

    master_idx = df_master.index[df_master['Article_No'] == search_id][0]
    product_name = df_master.at[master_idx, 'Part_Name']

    if movement_type.upper() == "IN":
        try:
            qty_int = int(qty)
        except Exception:
            qty_int = 1
        if qty_int < 1:
            qty_int = 1

        # Auto-assign Sr_No sequence if not provided.
        if sr_no is not None:
            sr_list = [int(sr_no)]
        else:
            existing_for_article = df_trans[df_trans["Article_No"] == search_id]
            max_sr = int(existing_for_article["Sr_No"].max()) if not existing_for_article.empty else 0
            sr_list = list(range(max_sr + 1, max_sr + qty_int + 1))

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        added_or_reset = 0
        tax_invoice_no = "" if tax_invoice_no is None else str(tax_invoice_no).strip()

        for sr in sr_list:
            unit_mask = (df_trans["Article_No"] == search_id) & (df_trans["Sr_No"] == int(sr))

            if unit_mask.any():
                # If this unit is already IN, don't duplicate it.
                current_status = df_trans.loc[unit_mask, "Status"].values[0]
                if current_status == "IN":
                    continue
                # If OUT, allow re-entry/reset to IN
                target_idx = df_trans[unit_mask].index[0]
                df_trans.at[target_idx, "Status"] = "IN"
                df_trans.at[target_idx, "Purchase_Type"] = purchase_type
                df_trans.at[target_idx, "Tax_Invoice_No"] = tax_invoice_no
                df_trans.at[target_idx, "Engineer"] = "Store"
                df_trans.at[target_idx, "Charges"] = 0
                df_trans.at[target_idx, "In_Date"] = now
                df_trans.at[target_idx, "Out_Date"] = None
                df_trans.at[target_idx, "Date"] = now
                added_or_reset += 1
            else:
                new_log = pd.DataFrame([{
                    "Article_No": search_id,
                    "Sr_No": int(sr),
                    "Purchase_Type": purchase_type,
                    "Tax_Invoice_No": tax_invoice_no,
                    "Status": "IN",
                    "Engineer": "Store",
                    "Charges": 0,
                    "In_Date": now,
                    "Out_Date": None,
                    "Date": now
                }])
                df_trans = pd.concat([df_trans, new_log], ignore_index=True)
                added_or_reset += 1

        # Update Master stock = count of IN units for this article
        in_count = int(((df_trans["Article_No"] == search_id) & (df_trans["Status"] == "IN")).sum())
        df_master.at[master_idx, "Stock_Level"] = in_count

        if len(sr_list) == 1:
            msg = f"✅ INWARD: {product_name} ({search_id}) Sr#{sr_list[0]} added to store."
        else:
            msg = f"✅ INWARD: {product_name} ({search_id}) added {added_or_reset}/{len(sr_list)} unit(s) to store."

    elif movement_type.upper() == "OUT":
        if engineer is None or str(engineer).strip() == "":
            return False, "❌ Select Engineer first."

        # If Sr_No not given, auto-pick if only one unit is IN, otherwise block.
        active_mask = (df_trans["Article_No"] == search_id) & (df_trans["Status"] == "IN")
        active_srs = sorted(df_trans.loc[active_mask, "Sr_No"].astype(int).unique().tolist())

        if not active_srs:
            return False, f"❌ ERROR: {search_id} is not in store (all units OUT)."

        if sr_no is None:
            if len(active_srs) == 1:
                sr_no = active_srs[0]
            else:
                return False, f"⚠️ Multiple units IN for {search_id}. Select Sr No: {active_srs}"

        unit_mask = (df_trans["Article_No"] == search_id) & (df_trans["Sr_No"] == int(sr_no)) & (df_trans["Status"] == "IN")
        if unit_mask.any():
            target_idx = df_trans[unit_mask].index[0]
            df_trans.at[target_idx, "Status"] = "OUT"
            df_trans.at[target_idx, "Engineer"] = engineer
            df_trans.at[target_idx, "Charges"] = df_master.at[master_idx, "SP"]
            out_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df_trans.at[target_idx, "Out_Date"] = out_now
            df_trans.at[target_idx, "Date"] = out_now

            in_count = int(((df_trans["Article_No"] == search_id) & (df_trans["Status"] == "IN")).sum())
            df_master.at[master_idx, "Stock_Level"] = in_count
            msg = f"📤 ISSUED: {product_name} ({search_id}) Sr#{int(sr_no)} given to {engineer}."
        else:
            return False, f"❌ ERROR: {search_id} Sr#{int(sr_no)} is not IN store (might be already OUT)."

    # 4. Save both sheets
    with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
        df_master.to_excel(writer, sheet_name='Master', index=False)
        df_trans.to_excel(writer, sheet_name='Transactions', index=False)
    
    return True, msg