import pandas as pd
import json
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ============================================================================
# CONFIGURATION - Edit these variables as needed
# ============================================================================

# Input files
CAMERA_INVENTORY_JSON = 'camera_inventory.json'
DIRECTORY_REPORT_EXCEL = 'INZ1A directory report.xlsx'
TRACKER_EXCEL = 'JP-NRT-INZ1A_switch_camera_tracker.xlsx'

# Sheet names
DIRECTORY_SHEET_NAME = 'Sheet1'
TRACKER_CAMERA_SHEET = 'camera'

# Output file
OUTPUT_EXCEL = 'JP-NRT-INZ1A_switch_camera_tracker_populated.xlsx'

# Column names in directory report (adjust if your Excel has different column names)
DIR_COL_CAMERA_NAME = 'Camera stream name'
DIR_COL_CAMERA_STATE = 'Camera state/status'
DIR_COL_IP_ADDRESS = 'IP address'
DIR_COL_MAC_ADDRESS = 'MAC address'

# ============================================================================

def normalize_mac(mac_address):
    """Normalize MAC address format to uppercase with colons"""
    if pd.isna(mac_address) or mac_address == '':
        return None
    # Remove any separators and convert to uppercase
    mac = str(mac_address).upper().replace(':', '').replace('-', '').replace('.', '')
    # Add colons back in standard format
    return ':'.join([mac[i:i+2] for i in range(0, len(mac), 2)])

# Define highlight colors
YELLOW_FILL = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # No name found
ORANGE_FILL = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')  # No switch info found
LIGHTBLUE_FILL = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')  # Duplicate MAC (same camera, different names)

def main():
    # Read the JSON file with camera inventory (switch info)
    print(f"Reading {CAMERA_INVENTORY_JSON}...")
    with open(CAMERA_INVENTORY_JSON, 'r') as f:
        camera_inventory = json.load(f)
    
    # Create a dictionary with MAC address as key
    inventory_dict = {}
    for entry in camera_inventory:
        mac = normalize_mac(entry['mac_address'])
        if mac:
            inventory_dict[mac] = {
                'switch_name': entry['switch_name'],
                'port': entry['port'],
                'vlan': entry.get('vlan', '')
            }
    
    print(f"Loaded {len(inventory_dict)} camera records from inventory")
    
    # Read the directory report Excel file
    print(f"\nReading {DIRECTORY_REPORT_EXCEL}...")
    directory_df = pd.read_excel(DIRECTORY_REPORT_EXCEL, sheet_name=DIRECTORY_SHEET_NAME)
    
    print(f"Total rows in directory report: {len(directory_df)}")
    
    # IMPORTANT: We only filter out formatting/status rows, NOT duplicate cameras
    # Duplicate cameras (same MAC/IP, different names) are kept intentionally
    original_count = len(directory_df)
    
    # Remove rows where camera name is NaN or empty (these are formatting rows)
    directory_df = directory_df[directory_df[DIR_COL_CAMERA_NAME].notna()]
    directory_df = directory_df[directory_df[DIR_COL_CAMERA_NAME].astype(str).str.strip() != '']
    
    # Remove rows that are ONLY status updates (empty camera name + "Streaming" status)
    # These are formatting rows, not actual duplicate cameras
    # We check if camera name is empty to ensure we don't remove real cameras
    directory_df = directory_df[
        ~((directory_df[DIR_COL_CAMERA_STATE] == 'Streaming') & 
          (directory_df[DIR_COL_CAMERA_NAME].astype(str).str.strip() == ''))
    ]
    
    filtered_count = len(directory_df)
    removed_count = original_count - filtered_count
    
    print(f"Filtered out {removed_count} empty/status formatting rows")
    print(f"Processing {filtered_count} camera records (including duplicates)")
    
    # Check for duplicate MACs to inform user
    directory_df['MAC_normalized'] = directory_df[DIR_COL_MAC_ADDRESS].apply(normalize_mac)
    duplicate_macs = directory_df[directory_df['MAC_normalized'].notna()].duplicated(subset=['MAC_normalized'], keep=False).sum()
    if duplicate_macs > 0:
        print(f"Note: Found {duplicate_macs} rows with duplicate MAC addresses (will be highlighted in light blue)")

    
    # Load the tracker workbook
    print(f"\nLoading {TRACKER_EXCEL}...")
    wb = load_workbook(TRACKER_EXCEL)
    ws = wb[TRACKER_CAMERA_SHEET]
    
    # Clear existing data (except header)
    print("Clearing existing data in camera sheet...")
    ws.delete_rows(2, ws.max_row)
    
    # Prepare data to write
    print("\nMatching cameras and preparing data...")
    row_num = 2
    matched_count = 0
    no_switch_info_count = 0
    no_name_info_count = 0
    duplicate_mac_count = 0
    
    # Track which MACs from inventory we've used
    used_macs = set()
    
    # Track MACs we've already written to detect duplicates
    written_macs = {}  # mac -> list of (row_num, camera_name)
    
    # NOTE: This script handles cases where the same camera (MAC/IP) appears 
    # with different names in the directory. When duplicates are found,
    # all name variations will be written to separate rows and highlighted
    # in light blue so you can identify which entries need reconciliation.
    
    # First pass: Process all cameras from directory report
    for idx, cam_row in directory_df.iterrows():
        mac = cam_row['MAC_normalized']
        camera_name = cam_row[DIR_COL_CAMERA_NAME]
        ip_address = cam_row[DIR_COL_IP_ADDRESS]
        original_mac = cam_row[DIR_COL_MAC_ADDRESS]
        
        # Check if this MAC was already written (duplicate detection)
        is_duplicate = mac in written_macs if mac else False
        
        # Look up switch info from inventory
        switch_info = inventory_dict.get(mac)
        
        if switch_info:
            # Write to Excel - fully matched
            ws.cell(row=row_num, column=1, value=camera_name)
            ws.cell(row=row_num, column=2, value=original_mac)
            ws.cell(row=row_num, column=3, value=ip_address)
            ws.cell(row=row_num, column=4, value=switch_info['switch_name'])
            ws.cell(row=row_num, column=5, value=switch_info['port'])
            
            # If this is a duplicate MAC, highlight in light blue
            if is_duplicate:
                for col in range(1, 6):
                    ws.cell(row=row_num, column=col).fill = LIGHTBLUE_FILL
                duplicate_mac_count += 1
                prev_entries = written_macs[mac]
                print(f"  LIGHT BLUE: Duplicate MAC {original_mac} found:")
                print(f"    - Previous: {prev_entries[-1][1]}")
                print(f"    - Current:  {camera_name}")
                # Also highlight the previous entries
                for prev_row, prev_name in prev_entries:
                    for col in range(1, 6):
                        ws.cell(row=prev_row, column=col).fill = LIGHTBLUE_FILL
            
            matched_count += 1
            if mac:
                used_macs.add(mac)
                if mac not in written_macs:
                    written_macs[mac] = []
                written_macs[mac].append((row_num, camera_name))
        else:
            # Write partial data - have name but no switch info (highlight ORANGE)
            ws.cell(row=row_num, column=1, value=camera_name)
            ws.cell(row=row_num, column=2, value=original_mac)
            ws.cell(row=row_num, column=3, value=ip_address)
            ws.cell(row=row_num, column=4, value='NOT FOUND')
            ws.cell(row=row_num, column=5, value='NOT FOUND')
            
            # Highlight the entire row in ORANGE
            for col in range(1, 6):
                ws.cell(row=row_num, column=col).fill = ORANGE_FILL
            
            no_switch_info_count += 1
            print(f"  ORANGE: No switch info found for {camera_name} (MAC: {original_mac})")
            
            # Track written MACs even if no switch info found
            if mac:
                if mac not in written_macs:
                    written_macs[mac] = []
                written_macs[mac].append((row_num, camera_name))
        
        row_num += 1
    
    # Second pass: Add cameras from inventory that don't have names (not in directory report)
    print("\nChecking for cameras in inventory without names...")
    for mac, switch_info in inventory_dict.items():
        if mac not in used_macs:
            # This MAC has switch info but no camera name (highlight YELLOW)
            ws.cell(row=row_num, column=1, value='NAME NOT FOUND')  # camera name
            ws.cell(row=row_num, column=2, value=mac)  # mac address (normalized format)
            ws.cell(row=row_num, column=3, value='')  # ip address
            ws.cell(row=row_num, column=4, value=switch_info['switch_name'])  # access switch
            ws.cell(row=row_num, column=5, value=switch_info['port'])  # uplink port
            
            # Highlight the entire row in YELLOW
            for col in range(1, 6):
                ws.cell(row=row_num, column=col).fill = YELLOW_FILL
            
            no_name_info_count += 1
            print(f"  YELLOW: No camera name found for MAC: {mac} on {switch_info['switch_name']} port {switch_info['port']}")
            
            row_num += 1
    
    # Save the workbook
    print(f"\nSaving updated tracker to {OUTPUT_EXCEL}...")
    wb.save(OUTPUT_EXCEL)
    
    # Print summary
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"Total cameras from directory: {len(directory_df)}")
    print(f"Total MACs from inventory: {len(inventory_dict)}")
    print(f"\n✓ Successfully matched: {matched_count}")
    print(f"⚠  ORANGE - Have name but no switch info: {no_switch_info_count}")
    print(f"⚠  YELLOW - Have switch info but no name: {no_name_info_count}")
    print(f"⚠  LIGHT BLUE - Duplicate MAC (same camera, different names): {duplicate_mac_count}")
    print(f"\nTotal rows written: {row_num - 2}")
    print(f"\nOutput saved to: {OUTPUT_EXCEL}")
    print("\nColor coding:")
    print("  - No color: Fully matched (name + switch info)")
    print("  - ORANGE: Camera name found but no switch info")
    print("  - YELLOW: Switch info found but no camera name")
    print("  - LIGHT BLUE: Same MAC/IP with different names (needs reconciliation)")
    print("\nNOTE: Light blue entries indicate the same physical camera")
    print("      appears with multiple names in the directory report.")
    print("      Review these entries to determine the correct name.")
    print("="*60)

if __name__ == "__main__":
    main()
