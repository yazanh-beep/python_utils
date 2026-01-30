import pandas as pd
import json
import openpyxl
import glob
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ============================================================================
# CONFIGURATION
# ============================================================================

DIRECTORY_REPORT_EXCEL = 'directory_report.xlsx'
TRACKER_EXCEL = 'camera-switch-tracker.xlsx'

DIRECTORY_SHEET_NAME = 'Sheet1'
TRACKER_CAMERA_SHEET = 'camera'

OUTPUT_EXCEL = 'camera-switch-tracker.xlsx'

# Default Column names (Script will auto-detect Exporter if this isn't exact)
DIR_COL_CAMERA_NAME = 'Camera stream name'
DIR_COL_CAMERA_STATE = 'Camera state/status'
DIR_COL_IP_ADDRESS = 'IP address'
DIR_COL_MAC_ADDRESS = 'MAC address'
DIR_COL_LOCATION = 'Location'
DIR_COL_EXPORTER = 'Exporter name'

# ============================================================================

def normalize_mac(mac_address):
    if pd.isna(mac_address) or mac_address == '':
        return None
    # Remove any separators and spaces, convert to uppercase
    mac = str(mac_address).upper().replace(':', '').replace('-', '').replace('.', '').replace(' ', '')
    try:
        # Reformat to XX:XX:XX:XX:XX:XX
        return ':'.join([mac[i:i+2] for i in range(0, len(mac), 2)])
    except:
        return str(mac_address)

# Define Highlight Colors
YELLOW_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # Inventory only
ORANGE_FILL = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')  # No switch info
LIGHTBLUE_FILL = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid') # Duplicate MAC
RED_FILL = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')       # IP Conflict

def get_inventory_file():
    files = glob.glob('camera_inventory*.json')
    if not files:
        return None
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

def main():
    print("Searching for inventory file...")
    camera_inventory_json = get_inventory_file()
    
    if not camera_inventory_json:
        print("CRITICAL ERROR: No JSON inventory file found.")
        return

    print(f"✓ Found and using: {camera_inventory_json}")

    with open(camera_inventory_json, 'r') as f:
        camera_inventory_data = json.load(f)
    
    if isinstance(camera_inventory_data, dict):
        camera_inventory = camera_inventory_data.get('cameras', [])
    else:
        camera_inventory = camera_inventory_data
    
    inventory_dict = {}
    for entry in camera_inventory:
        mac = normalize_mac(entry['mac_address'])
        if mac:
            inventory_dict[mac] = {
                'switch_name': entry['switch_name'],
                'switch_type': entry.get('switch_type', 'UNKNOWN'),
                'port': entry['port']
            }
    
    print(f"Loaded {len(inventory_dict)} unique MACs from inventory.")

    print(f"\nReading {DIRECTORY_REPORT_EXCEL}...")
    # Force string type for MAC column to avoid scientific notation issues
    directory_df = pd.read_excel(DIRECTORY_REPORT_EXCEL, sheet_name=DIRECTORY_SHEET_NAME, dtype={DIR_COL_MAC_ADDRESS: str})
    
    # --- FIX 1: Clean column names ---
    directory_df.columns = directory_df.columns.str.strip()
    
    # --- FIX 2: Dynamically find the Exporter column ---
    # We use a local variable 'actual_exporter_col' instead of modifying the global one
    actual_exporter_col = DIR_COL_EXPORTER
    
    if DIR_COL_EXPORTER not in directory_df.columns:
        print(f"⚠ Standard column '{DIR_COL_EXPORTER}' not found. Searching...")
        # Look for any column containing "Exporter" (case-insensitive)
        found_col = next((c for c in directory_df.columns if 'exporter' in c.lower()), None)
        if found_col:
            print(f"✓ Found and using column: '{found_col}'")
            actual_exporter_col = found_col
        else:
            print(f"❌ CRITICAL: Could not find any column looking like 'Exporter'. Exporter data will be empty.")
            actual_exporter_col = None

    # Filter out empty/formatting rows
    if DIR_COL_CAMERA_NAME in directory_df.columns:
        directory_df = directory_df[directory_df[DIR_COL_CAMERA_NAME].notna()]
        directory_df = directory_df[directory_df[DIR_COL_CAMERA_NAME].astype(str).str.strip() != '']
    
    directory_df['MAC_normalized'] = directory_df[DIR_COL_MAC_ADDRESS].apply(normalize_mac)
    
    print(f"\nLoading {TRACKER_EXCEL}...")
    wb = load_workbook(TRACKER_EXCEL)
    ws = wb[TRACKER_CAMERA_SHEET]
    
    # Clear existing data
    ws.delete_rows(2, ws.max_row)
    
    used_macs = set()
    written_macs = {} 
    
    # Buffer to store all rows before writing (allows sorting)
    # Structure: {'sort_key': (SwitchName, Port), 'data': [CellValues], 'color': FillObj}
    rows_buffer = []

    # --- PASS 1: Process Cameras from Directory Report ---
    for idx, cam_row in directory_df.iterrows():
        mac = cam_row['MAC_normalized']
        camera_name = cam_row[DIR_COL_CAMERA_NAME]
        ip_address = cam_row[DIR_COL_IP_ADDRESS]
        original_mac = cam_row[DIR_COL_MAC_ADDRESS]
        
        # Safely get columns using the dynamically found column name
        status = cam_row.get(DIR_COL_CAMERA_STATE, '')
        location = cam_row.get(DIR_COL_LOCATION, '')
        
        exporter = ''
        if actual_exporter_col:
            exporter = cam_row.get(actual_exporter_col, '')
        
        # Conflict detection
        is_duplicate_mac = mac in written_macs if mac else False
        is_same_mac_diff_ip = False
        if is_duplicate_mac and ip_address:
            for r, n, ip in written_macs[mac]:
                if ip and ip != ip_address:
                    is_same_mac_diff_ip = True
                    break

        switch_info = inventory_dict.get(mac)
        
        row_data = [
            camera_name,        # 1
            original_mac,       # 2
            ip_address,         # 3
            "",                 # 4 (Switch)
            "",                 # 5 (Port)
            status,             # 6
            location,           # 7
            exporter            # 8
        ]
        
        row_color = None
        sort_key = ("ZZZ_NOT_FOUND", "ZZZ") # Default to bottom

        if switch_info:
            switch_display = f"{switch_info['switch_name']} [{switch_info['switch_type']}]"
            row_data[3] = switch_display
            row_data[4] = switch_info['port']
            sort_key = (switch_info['switch_name'], switch_info['port'])
            used_macs.add(mac)
        else:
            row_data[3] = 'NOT FOUND'
            row_data[4] = 'NOT FOUND'
            row_color = ORANGE_FILL

        if is_same_mac_diff_ip:
            row_color = RED_FILL
        elif is_duplicate_mac:
            row_color = LIGHTBLUE_FILL

        rows_buffer.append({'sort': sort_key, 'data': row_data, 'color': row_color})

        if mac:
            if mac not in written_macs: written_macs[mac] = []
            written_macs[mac].append((0, camera_name, ip_address))

    # --- PASS 2: Add Inventory Items Not In Directory (Yellow) ---
    print("Adding unmapped inventory items...")
    count_yellow = 0
    for mac, info in inventory_dict.items():
        if mac not in used_macs:
            switch_display = f"{info['switch_name']} [{info['switch_type']}]"
            
            row_data = [
                'NAME NOT FOUND',   # 1
                mac,                # 2
                '',                 # 3
                switch_display,     # 4
                info['port'],       # 5
                'UNKNOWN',          # 6
                '',                 # 7
                ''                  # 8
            ]
            
            sort_key = (info['switch_name'], info['port'])
            rows_buffer.append({'sort': sort_key, 'data': row_data, 'color': YELLOW_FILL})
            count_yellow += 1

    print(f"Added {count_yellow} extra devices found on switches.")

    # --- PASS 3: Sort and Write ---
    print("Sorting data by Switch and Port to group devices...")
    
    # Sort key: (Switch Name, Port). Converting to str ensures stability.
    rows_buffer.sort(key=lambda x: (str(x['sort'][0]), str(x['sort'][1])))

    print("Writing to Excel...")
    row_num = 2
    for item in rows_buffer:
        data = item['data']
        color = item['color']
        
        for i, value in enumerate(data, 1):
            cell = ws.cell(row=row_num, column=i, value=value)
            if color:
                cell.fill = color
        
        row_num += 1

    wb.save(OUTPUT_EXCEL)
    print(f"Successfully saved {row_num-2} sorted rows to {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()
