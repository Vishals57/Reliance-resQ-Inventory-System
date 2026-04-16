import pandas as pd
import os
from datetime import datetime

DB_FILE = "resQ_Inventory.xlsx"

def initialize_db():
    if not os.path.exists(DB_FILE):
        df = pd.DataFrame(columns=["Part_ID", "Part_Name", "Stock_Level", "Last_Updated"])
        df.to_excel(DB_FILE, index=False)
        print("✅ System: Database Initialized.")

def update_stock(part_id, part_name, movement_type):
    # Load the current database
    df = pd.read_excel(DB_FILE)
    
    # Clean the data: Ensure Part_ID is treated as a string and stripped of spaces
    df['Part_ID'] = df['Part_ID'].astype(str).str.strip()
    search_id = str(part_id).strip()
    
    movement_type = movement_type.upper()
    
    # CHECK: Does this Part_ID already exist in our Excel?
    if search_id in df['Part_ID'].values:
        # Get the index of the row where the ID matches
        idx = df.index[df['Part_ID'] == search_id][0]
        
        if movement_type == "IN":
            # UPDATE EXISTING ROW: Add 1 to stock
            df.at[idx, 'Stock_Level'] = int(df.at[idx, 'Stock_Level']) + 1
            msg = f"📈 Stock Increased: {part_name} (New Level: {df.at[idx, 'Stock_Level']})"
        
        elif movement_type == "OUT":
            current_stock = int(df.at[idx, 'Stock_Level'])
            if current_stock > 0:
                # UPDATE EXISTING ROW: Subtract 1
                df.at[idx, 'Stock_Level'] = current_stock - 1
                msg = f"📉 Stock Decreased: {part_name} (Remaining: {current_stock - 1})"
            else:
                return False, f"❌ ERROR: {part_name} is OUT OF STOCK!"
        
        # Update the name and timestamp for that specific row
        df.at[idx, 'Part_Name'] = part_name
        df.at[idx, 'Last_Updated'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # IF PART DOES NOT EXIST: Create it for the first time
    else:
        if movement_type == "IN":
            new_row = {
                "Part_ID": search_id, 
                "Part_Name": part_name, 
                "Stock_Level": 1,
                "Last_Updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            msg = f"🆕 New Entry: {part_name} Registered with 1 unit."
        else:
            return False, "❌ ERROR: Cannot scan OUT a part that was never registered."

    # Save the cleaned and updated dataframe back to Excel
    df.to_excel(DB_FILE, index=False)
    return True, msg