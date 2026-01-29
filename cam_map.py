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

# Column names in directory report
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
    mac = str(mac_address).upper().replace(':', '').replace('-', '').replace('.', '')
    return ':'.join([mac[i:i+2] for i in range(0, len(mac), 2)])

YELLOW_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
ORANGE_FILL = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
LIGHTBLUE_FILL = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
RED_FILL = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

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
    
    print(f"\nReading {DIRECTORY_REPORT_EXCEL}...")
    directory_df = pd.read_excel(DIRECTORY_REPORT_EXCEL, sheet_name=DIRECTORY_SHEET_NAME)
    
    # Filter formatting rows
    directory_df = directory_df[directory_df[DIR_COL_CAMERA_NAME].notna()]
    directory_df = directory_df[directory_df[DIR_COL_CAMERA_NAME].astype(str).str.strip() != '']
    
    directory_df['MAC_normalized'] = directory_df[DIR_COL_MAC_ADDRESS].apply(normalize_mac)
    
    print(f"\nLoading {TRACKER_EXCEL}...")
    wb = load_workbook(TRACKER_EXCEL)
    ws = wb[TRACKER_CAMERA_SHEET]
    
    # Clear existing data
    ws.delete_rows(2, ws.max_row)
    
    row_num = 2
    used_macs = set()
    written_macs = {} 
    written_ips = {}

    for idx, cam_row in directory_df.iterrows():
        mac = cam_row['MAC_normalized']
        camera_name = cam_row[DIR_COL_CAMERA_NAME]
        ip_address = cam_row[DIR_COL_IP_ADDRESS]
        original_mac = cam_row[DIR_COL_MAC_ADDRESS]
        
        # New data points
        status = cam_row[DIR_COL_CAMERA_STATE]
        location = cam_row[DIR_COL_LOCATION]
        exporter = cam_row[DIR_COL_EXPORTER]
        
        # Duplicate/Conflict detection logic
        is_duplicate_mac = mac in written_macs if mac else False
        is_same_mac_diff_ip = False
        if is_duplicate_mac and ip_address:
            for r, n, ip in written_macs[mac]:
                if ip and ip != ip_address:
                    is_same_mac_diff_ip = True
                    break

        switch_info = inventory_dict.get(mac)
        
        # Write basic columns
        ws.cell(row=row_num, column=1, value=camera_name)
        ws.cell(row=row_num, column=2, value=original_mac)
        ws.cell(row=row_num, column=3, value=ip_address)
        
        # Write switch info if found
        if switch_info:
            switch_display = f"{switch_info['switch_name']} [{switch_info['switch_type']}]"
            ws.cell(row=row_num, column=4, value=switch_display)
            ws.cell(row=row_num, column=5, value=switch_info['port'])
            used_macs.add(mac)
        else:
            ws.cell(row=row_num, column=4, value='NOT FOUND')
            ws.cell(row=row_num, column=5, value='NOT FOUND')
            for col in range(1, 9): # Highlight orange if no switch info
                ws.cell(row=row_num, column=col).fill = ORANGE_FILL

        # Write requested extra columns (6, 7, 8)
        ws.cell(row=row_num, column=6, value=status)
        ws.cell(row=row_num, column=7, value=location)
        ws.cell(row=row_num, column=8, value=exporter)

        # Apply special highlighting
        if is_same_mac_diff_ip:
            for col in range(1, 9):
                ws.cell(row=row_num, column=col).fill = RED_FILL
        elif is_duplicate_mac:
            for col in range(1, 9):
                ws.cell(row=row_num, column=col).fill = LIGHTBLUE_FILL

        # Track written data
        if mac:
            if mac not in written_macs: written_macs[mac] = []
            written_macs[mac].append((row_num, camera_name, ip_address))
        
        row_num += 1

    # Add inventory items not in directory (Yellow)
    for mac, info in inventory_dict.items():
        if mac not in used_macs:
            ws.cell(row=row_num, column=1, value='NAME NOT FOUND')
            ws.cell(row=row_num, column=2, value=mac)
            ws.cell(row=row_num, column=4, value=f"{info['switch_name']} [{info['switch_type']}]")
            ws.cell(row=row_num, column=5, value=info['port'])
            for col in range(1, 9):
                ws.cell(row=row_num, column=col).fill = YELLOW_FILL
            row_num += 1

    wb.save(OUTPUT_EXCEL)
    print(f"Successfully saved {row_num-2} rows to {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()
