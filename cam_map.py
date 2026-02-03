import pandas as pd
import json
import glob
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ============================================================================
# CONFIGURATION
# ============================================================================
DIRECTORY_REPORT_EXCEL = 'directory_report.xlsx'
TRACKER_EXCEL = 'camera-switch-tracker.xlsx'
OUTPUT_EXCEL = 'camera-switch-tracker.xlsx'

# Sheet Names
SHEET_CAMERA = 'camera'
SHEET_SERVER = 'Server'

# Exact Column Names from Directory Report
DIR_COL_CAMERA_NAME = 'Camera stream name'
DIR_COL_CAMERA_STATE = 'Camera state/status'
DIR_COL_IP_ADDRESS = 'IP address'
DIR_COL_MAC_ADDRESS = 'MAC address'
DIR_COL_LOCATION = 'Location'
DIR_COL_EXPORTER = 'Exporter name'

# Colors
YELLOW_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
ORANGE_FILL = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')

def normalize_mac(mac_address):
    if pd.isna(mac_address) or str(mac_address).strip() == '':
        return None
    mac = str(mac_address).upper().replace(':', '').replace('-', '').replace('.', '')
    return ':'.join([mac[i:i+2] for i in range(0, len(mac), 2)])

def get_inventory_file():
    files = glob.glob('camera_inventory*.json')
    return max(files, key=os.path.getmtime) if files else None

def main():
    print("🚀 Starting Update (Preserving Template Headers)...")
    
    # ---------------------------------------------------------
    # 1. LOAD INVENTORY
    # ---------------------------------------------------------
    inventory_file = get_inventory_file()
    if not inventory_file:
        print("❌ ERROR: No inventory JSON found.")
        return

    with open(inventory_file, 'r') as f:
        data = json.load(f)
        inventory_list = data.get('cameras', data) if isinstance(data, dict) else data

    inventory_dict = {}
    for entry in inventory_list:
        if 'mac_address' in entry:
            mac = normalize_mac(entry['mac_address'])
            if mac:
                inventory_dict[mac] = entry

    # ---------------------------------------------------------
    # 2. READ & CONSOLIDATE DIRECTORY REPORT
    # ---------------------------------------------------------
    print(f"Reading {DIRECTORY_REPORT_EXCEL}...")
    df = pd.read_excel(DIRECTORY_REPORT_EXCEL, sheet_name='Sheet1')
    df.columns = df.columns.astype(str).str.strip()

    # CRITICAL: Forward Fill MAC addresses to handle split blocks
    if DIR_COL_MAC_ADDRESS in df.columns:
        df[DIR_COL_MAC_ADDRESS] = df[DIR_COL_MAC_ADDRESS].ffill()

    # Consolidate Data:
    # We want static info (Name, IP) from the TOP row (first)
    # We want Status from the BOTTOM row (last)
    consolidated_data = {}
    
    for _, row in df.iterrows():
        raw_mac = row.get(DIR_COL_MAC_ADDRESS)
        mac = normalize_mac(raw_mac)
        if not mac: continue

        # If new MAC, initialize with current row (Top Row)
        if mac not in consolidated_data:
            consolidated_data[mac] = {
                'raw_mac': raw_mac,
                'name': row.get(DIR_COL_CAMERA_NAME, ''),
                'ip': row.get(DIR_COL_IP_ADDRESS, ''),
                'loc': row.get(DIR_COL_LOCATION, ''),
                'exp': row.get(DIR_COL_EXPORTER, ''),
                'status': row.get(DIR_COL_CAMERA_STATE, '')
            }
        else:
            # If MAC exists, we are on the Bottom Row.
            # Update STATUS to grab the latest/bottom value
            # Update others only if they were empty previously
            consolidated_data[mac]['status'] = row.get(DIR_COL_CAMERA_STATE, '')
            
            if not consolidated_data[mac]['name']: 
                consolidated_data[mac]['name'] = row.get(DIR_COL_CAMERA_NAME, '')
            if not consolidated_data[mac]['ip']: 
                consolidated_data[mac]['ip'] = row.get(DIR_COL_IP_ADDRESS, '')

    # ---------------------------------------------------------
    # 3. UPDATE EXCEL (PRESERVE HEADERS)
    # ---------------------------------------------------------
    wb = load_workbook(TRACKER_EXCEL)
    
    # Helper to clear data but keep headers
    def prepare_sheet(sheet_name):
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            # If creating new, we must add headers manually, 
            # otherwise we assume headers exist in row 1
            headers = ["Camera Name", "MAC Address", "IP Address", "Switch Name", "Port", "Location", "Cloud Cam Exporter", "Status"]
            for c, h in enumerate(headers, 1):
                ws.cell(row=1, column=c, value=h)
        
        ws = wb[sheet_name]
        # Clear data from row 2 downwards
        if ws.max_row >= 2:
            ws.delete_rows(2, ws.max_row)
        return ws

    ws_cam = prepare_sheet(SHEET_CAMERA)
    ws_srv = prepare_sheet(SHEET_SERVER)
    
    cam_row = 2
    srv_row = 2
    used_macs = set()

    print("Writing consolidated data...")

    # Process Consolidated Directory Data
    for mac, data in consolidated_data.items():
        inv = inventory_dict.get(mac)
        
        # Determine if Server
        is_server = False
        switch_str = "NOT FOUND"
        port_str = "NOT FOUND"
        fill_color = ORANGE_FILL

        if inv:
            s_name = str(inv.get('switch_name', '')).upper()
            s_type = str(inv.get('switch_type', '')).upper()
            
            # CHECK BOTH NAME AND TYPE FOR 'SERVER'
            if s_type == 'SERVER' or 'SERVER' in s_name:
                is_server = True
            
            switch_str = f"{inv.get('switch_name')} [{inv.get('switch_type')}]"
            port_str = inv.get('port')
            used_macs.add(mac)
            fill_color = None

        # Select Sheet
        if is_server:
            ws = ws_srv
            r = srv_row
            srv_row += 1
        else:
            ws = ws_cam
            r = cam_row
            cam_row += 1

        # Write Row
        ws.cell(row=r, column=1, value=data['name'])
        ws.cell(row=r, column=2, value=data['raw_mac'])
        ws.cell(row=r, column=3, value=data['ip'])
        ws.cell(row=r, column=4, value=switch_str)
        ws.cell(row=r, column=5, value=port_str)
        ws.cell(row=r, column=6, value=data['loc'])
        ws.cell(row=r, column=7, value=data['exp'])
        ws.cell(row=r, column=8, value=data['status']) # Bottom status

        if fill_color:
            for c in range(1, 9): ws.cell(row=r, column=c).fill = fill_color

    # Process Remaining Inventory (Inventory Only)
    for mac, inv in inventory_dict.items():
        if mac not in used_macs:
            s_name = str(inv.get('switch_name', '')).upper()
            s_type = str(inv.get('switch_type', '')).upper()
            is_server = (s_type == 'SERVER' or 'SERVER' in s_name)

            if is_server:
                ws = ws_srv
                r = srv_row
                srv_row += 1
            else:
                ws = ws_cam
                r = cam_row
                cam_row += 1

            ws.cell(row=r, column=1, value="INVENTORY ONLY")
            ws.cell(row=r, column=2, value=mac)
            ws.cell(row=r, column=4, value=inv.get('switch_name'))
            ws.cell(row=r, column=5, value=inv.get('port'))
            
            for c in range(1, 9): ws.cell(row=r, column=c).fill = YELLOW_FILL

    wb.save(OUTPUT_EXCEL)
    print(f"✅ Update Complete. \n   Camera Sheet Rows: {cam_row-2}\n   Server Sheet Rows: {srv_row-2}")

if __name__ == "__main__":
    main()
