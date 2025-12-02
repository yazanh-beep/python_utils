import ipaddress
import sys
from typing import List, Dict, Any, Tuple
import pandas as pd 

# --- Configuration for Subnetting Scheme ---
NUM_BLDGS_BASE = 10 
NUM_SPECIAL = 1 
NUM_RESERVED = 2
BUILDING_IDS = [f"BLDG {i}" for i in range(1, NUM_BLDGS_BASE + 1)] 
NUM_TOTAL_BUILDING_BLOCKS = len(BUILDING_IDS) # 10 locations total

# Standard Column Headers for insertion
COLUMN_HEADERS = {
    "Building": "Building", "Assigned Equipment": "Assigned Equipment", 
    "Network Address": "Network Address", "Assigned IP Range": "Assigned IP Range", 
    "CIDR": "CIDR", "Subnet Mask": "Subnet Mask", 
    "SVI GATEWAY": "SVI GATEWAY", "Hosts": "Hosts"
}

# --- The MASTER_PLAN ---
MASTER_PLAN = [
    # --- PHASE 1: Linear Allocation ---
    
    # 1. Cameras (10 + 2 Reserved)
    # Header placeholder (range calculated later)
    {"phase": 1, "type": "header_with_range", "carving_name": "Cameras", "purpose": "Cameras", "cidr": 21, "count": 12, "carving_label": "Cameras"},
    {"phase": 1, "type": "allocation", "category": "Cameras", "purpose": "Cameras", "cidr": 21, "count": NUM_TOTAL_BUILDING_BLOCKS, "building_specific": True, "carving_label": "Cameras"}, 
    {"phase": 1, "type": "allocation", "category": "Cameras", "purpose": "Reserved", "cidr": 21, "count": NUM_RESERVED, "building_specific": False, "carving_label": "Cameras"}, 
    
    {"phase": 1, "type": "blank_row"}, 
    
    # 2. Switch MGMT Aggregate Block (/20)
    {"phase": 1, "type": "aggregate", "category": "Switch MGMT", "purpose": "Switch MGMT Aggregate Block", "cidr": 20, "name": "SWITCH_MGMT_AGGREGATE"},
    
    {"phase": 1, "type": "blank_row"}, 
    
    # 3. OSPF PEERING Aggregate Block (/21)
    {"phase": 1, "type": "aggregate", "category": "OSPF PEERING", "purpose": "OSPF Peering Aggregate Block", "cidr": 21, "name": "OSPF_PEERING_AGGREGATE"},

    {"phase": 1, "type": "blank_row"}, 
    
    # 4. PIDS and Intercoms Aggregate Block (/20)
    {"phase": 1, "type": "aggregate", "category": "PIDS and Intercoms", "purpose": "PIDS and Intercoms Aggregate Block", "cidr": 20, "name": "PIDS_INTERCOM_BLOCK"},
    
    {"phase": 1, "type": "blank_row"}, 
    
    # 5. UNUSED REMAINDER (Master /17 Remainder)
    {"phase": 1, "type": "remainder", "category": "UNUSED (Master /17 Remainder)", "purpose": "UNUSED", "cidr": 24}, 


    # --- PHASE 2: Detailed Carvings ---
    
    # Carve 1: Switch MGMT (from SWITCH_MGMT_AGGREGATE)
    {"phase": 2, "type": "header_with_range", "base_block": "SWITCH_MGMT_AGGREGATE", "carving_name": "switch MGMT", "purpose": "management", "cidr": 24, "carving_label": "management"},
    {"phase": 2, "type": "allocation", "base_block": "SWITCH_MGMT_AGGREGATE", "purpose": "management", "cidr": 24, "count": NUM_BLDGS_BASE + NUM_RESERVED, "building_specific": True, "carving_label": "management"},
    
    {"phase": 2, "type": "blank_row"}, 

    # Carve 2: Servers (from SWITCH_MGMT_AGGREGATE)
    {"phase": 2, "type": "header_with_range", "base_block": "SWITCH_MGMT_AGGREGATE", "carving_name": "Servers", "purpose": "Servers", "cidr": 27, "carving_label": "Servers"},
    {"phase": 2, "type": "allocation", "base_block": "SWITCH_MGMT_AGGREGATE", "purpose": "Servers", "cidr": 27, "count": NUM_BLDGS_BASE + NUM_RESERVED, "building_specific": True, "carving_label": "Servers"},
    
    {"phase": 2, "type": "blank_row"}, 
    
    # Carve 3: OSPF PEERING (from OSPF_PEERING_AGGREGATE)
    {"phase": 2, "type": "header_with_range", "base_block": "OSPF_PEERING_AGGREGATE", "carving_name": "IP OSPF PEERING", "purpose": "IP OSPF PEERING", "cidr": 31, "carving_label": "IP OSPF PEERING"},
    {"phase": 2, "type": "allocation", "base_block": "OSPF_PEERING_AGGREGATE", "purpose": "IP OSPF PEERING", "cidr": 31, "count": NUM_BLDGS_BASE + NUM_RESERVED, "building_specific": True, "carving_label": "IP OSPF PEERING"},

    # --- Unused Remainder of Switch MGMT Block ---
    {"phase": 2, "type": "blank_row"}, 
    {"phase": 2, "type": "header", "category": "UNUSED (Switch MGMT Block Remainder)", "purpose": "UNUSED"},
    {"phase": 2, "type": "carve_remainder", "base_block": "SWITCH_MGMT_AGGREGATE", "purpose": "UNUSED", "cidr": 24, "carving_label": "UNUSED"},
    
    {"phase": 2, "type": "blank_row"}, 

    # Carve 4A: PIDS (from PIDS_INTERCOM_BLOCK)
    {"phase": 2, "type": "header_with_range", "base_block": "PIDS_INTERCOM_BLOCK", "carving_name": "PIDS", "purpose": "PIDS", "cidr": 24, "carving_label": "PIDS"},
    {"phase": 2, "type": "allocation", "base_block": "PIDS_INTERCOM_BLOCK", "purpose": "PIDS", "cidr": 24, "count": NUM_BLDGS_BASE + NUM_RESERVED, "building_specific": True, "carving_label": "PIDS"},
    
    {"phase": 2, "type": "blank_row"},
    
    # Carve 4B: Intercoms (from PIDS_INTERCOM_BLOCK)
    {"phase": 2, "type": "header_with_range", "base_block": "PIDS_INTERCOM_BLOCK", "carving_name": "Intercoms", "purpose": "Intercoms", "cidr": 27, "carving_label": "Intercoms"},
    {"phase": 2, "type": "allocation", "base_block": "PIDS_INTERCOM_BLOCK", "purpose": "Intercoms", "cidr": 27, "count": NUM_BLDGS_BASE + NUM_RESERVED, "building_specific": True, "carving_label": "Intercoms"},

    # --- Unused Remainder of PIDS/Intercoms Block ---
    {"phase": 2, "type": "blank_row"}, 
    {"phase": 2, "type": "header", "category": "UNUSED (PIDS/Intercoms Block Remainder)", "purpose": "UNUSED"},
    {"phase": 2, "type": "carve_remainder", "base_block": "PIDS_INTERCOM_BLOCK", "purpose": "UNUSED", "cidr": 24, "carving_label": "UNUSED"},
]

EMPTY_ROW_TEMPLATE = {
    "Building": "", "Assigned Equipment": "", "Network Address": "", 
    "Assigned IP Range": "", "CIDR": "", "Usable Hosts": "",
    "Subnet Mask": "", "SVI GATEWAY": ""
}

def format_allocation_row(network: ipaddress.IPv4Network, building: str, purpose: str) -> Dict[str, Any]:
    """Formats a single allocated network row."""
    if network.prefixlen >= 31:
        usable_hosts = network.num_addresses
        svi_gateway = 'N/A'
        ip_range = f"{network.network_address} - {network.broadcast_address}" 
    else:
        usable_hosts = network.num_addresses - 2
        svi_gateway = str(network.network_address + 1)
        first_usable = str(network.network_address + 1)
        last_usable = str(network.broadcast_address - 1)
        ip_range = f"{first_usable} - {last_usable}"

    # FORCE SVI GATEWAY EMPTY FOR ALL ROWS (User Request)
    svi_gateway = ""

    return {
        "Building": building,
        "Assigned Equipment": purpose,
        "Network Address": str(network.network_address),
        "Assigned IP Range": ip_range,
        "CIDR": f"/{network.prefixlen}",
        "Hosts": usable_hosts, 
        "Subnet Mask": str(network.netmask),
        "SVI GATEWAY": svi_gateway
    }

def generate_ip_allocation(master_network_str: str, plan: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], ipaddress.IPv4Address]:
    try:
        master_network = ipaddress.IPv4Network(master_network_str)
    except ipaddress.AddressValueError:
        print(f"Error: Invalid IP network format.")
        return [], ipaddress.IPv4Address(master_network_str.split('/')[0])

    current_network_address = master_network.network_address
    aggregate_block_pointers = {}
    carving_pointers = {}
    allocation_results = []
    
    print(f"\n--- Starting VLSM Allocation from {master_network_str} ---")
    
    # --- PHASE 1 Loop ---
    for item in [p for p in plan if p.get('phase') == 1]:
        item_type = item.get("type")
        if item_type == "blank_row":
            allocation_results.append(EMPTY_ROW_TEMPLATE.copy())
            continue

        category = item.get('category')
        purpose = item.get('purpose')
        count = item.get('count', 1)
        target_cidr = item.get('cidr')

        if item_type == "header":
            header_row = EMPTY_ROW_TEMPLATE.copy()
            header_row["Building"] = f"{category}"
            allocation_results.append(header_row)
            continue
        elif item_type == "header_with_range":
            header_row = EMPTY_ROW_TEMPLATE.copy()
            header_row['is_range_placeholder'] = True 
            header_row['carving_name'] = item['carving_name'] if 'carving_name' in item else category
            header_row['allocation_start_index'] = len(allocation_results) + 1 # +1 because we insert column headers next
            header_row['carving_label'] = item.get('carving_label')
            allocation_results.append(header_row)
            allocation_results.append(COLUMN_HEADERS.copy())
            continue
        elif item_type == "remainder":
            pass # Handled implicitly if loop breaks
        
        elif item_type == "allocation" or item_type == "aggregate":
            building_counter = 0
            if item.get('building_specific'):
                pass

            for i in range(count):
                try:
                    network_to_allocate = ipaddress.ip_network(f"{current_network_address}/{target_cidr}", strict=False)
                    if current_network_address != network_to_allocate.network_address:
                         current_network_address = network_to_allocate.network_address
                         network_to_allocate = ipaddress.ip_network(f"{current_network_address}/{target_cidr}", strict=False)
                except ValueError:
                    print(f"FATAL: Allocation failed for {purpose}.")
                    return allocation_results, current_network_address
                
                if not master_network.overlaps(network_to_allocate):
                    print(f"FATAL: Allocation outside master.")
                    return allocation_results, current_network_address

                building_name = ""
                is_reserved = False
                if item.get('building_specific'):
                    is_reserved = (i >= NUM_BLDGS_BASE)
                    if not is_reserved:
                        building_name = f"BLDG {i + 1}"
                    else:
                        building_name = "Reserved"
                elif purpose == "Reserved":
                    building_name = "Reserved"
                
                if item_type == "aggregate":
                    block_name = item['name']
                    aggregate_block_pointers[block_name] = network_to_allocate
                    carving_pointers[block_name] = network_to_allocate.network_address
                
                row = format_allocation_row(network_to_allocate, building_name, purpose)
                if 'carving_label' in item:
                    row['carving_label'] = item['carving_label']
                
                if item_type != "aggregate":
                    allocation_results.append(row)
                
                current_network_address = network_to_allocate.broadcast_address + 1

    # --- PHASE 2 Loop ---
    for item in [p for p in plan if p.get('phase') == 2]:
        item_type = item.get("type")
        if item_type == "blank_row":
            allocation_results.append(EMPTY_ROW_TEMPLATE.copy())
            continue
        
        base_block_name = item.get('base_block')
        # If carving from aggregate, ensure it exists
        if base_block_name and base_block_name not in aggregate_block_pointers:
             continue

        if base_block_name:
            carving_address = carving_pointers.get(base_block_name)
            base_network = aggregate_block_pointers.get(base_block_name)

        if item_type == "header":
            header_row = EMPTY_ROW_TEMPLATE.copy()
            header_row["Building"] = f"{item['category']}"
            allocation_results.append(header_row)
            continue
        
        if item_type == "header_with_range":
            header_row = EMPTY_ROW_TEMPLATE.copy()
            header_row['is_range_placeholder'] = True 
            header_row['carving_name'] = item['carving_name']
            header_row['allocation_start_index'] = len(allocation_results) + 1
            header_row['carving_label'] = item['carving_label']
            header_row['aggregate_network'] = base_network # Fallback
            allocation_results.append(header_row)
            allocation_results.append(COLUMN_HEADERS.copy())
            continue
        
        elif item_type == "allocation":
            target_cidr = item['cidr']
            count = item['count']
            purpose = item['purpose']
            
            for i in range(count):
                try:
                    network_to_allocate = ipaddress.ip_network(f"{carving_address}/{target_cidr}", strict=False)
                except ValueError:
                    break
                
                if network_to_allocate.broadcast_address > base_network.broadcast_address:
                    print(f"Error: Aggregate overflow in {base_block_name}")
                    break
                
                building_name = ""
                is_reserved = False
                if item.get('building_specific'):
                    is_reserved = (i >= NUM_BLDGS_BASE)
                    if not is_reserved:
                        building_name = f"BLDG {i + 1}"
                    else:
                        building_name = "Reserved"
                
                # Determine final purpose name (override if reserved)
                p_final = "Reserved" if is_reserved else purpose

                row = format_allocation_row(network_to_allocate, building_name, p_final)
                row['carving_label'] = item['carving_label']
                allocation_results.append(row)
                
                carving_pointers[base_block_name] = network_to_allocate.broadcast_address + 1
                carving_address = carving_pointers[base_block_name]

        elif item_type == "carve_remainder":
            purpose = item['purpose']
            target_cidr = item['cidr']
            
            remaining = ipaddress.ip_network(f"{carving_address}/{base_network.prefixlen}", strict=False)
            if remaining.network_address >= base_network.broadcast_address:
                continue
                
            for sub in remaining.subnets(new_prefix=target_cidr):
                if sub.broadcast_address > base_network.broadcast_address: break
                if base_network.overlaps(sub):
                    row = format_allocation_row(sub, "UNUSED", purpose)
                    row['carving_label'] = item['carving_label']
                    allocation_results.append(row)
                    carving_pointers[base_block_name] = sub.broadcast_address + 1
                else:
                    break
    
    # --- POST-PROCESSING: Fill Headers ---
    for i, row in enumerate(allocation_results):
        if row.get('is_range_placeholder'):
            carving_name = row['carving_name']
            target_label = row['carving_label']
            start_idx = row['allocation_start_index']
            
            # Find all rows that belong to this section (matching carving_label)
            # Scan forward from start_idx until we hit a blank row or another header
            found_subnets = []
            for j in range(start_idx, len(allocation_results)):
                candidate = allocation_results[j]
                if candidate.get("Building") == "": # Blank row check
                    break
                if candidate.get("is_range_placeholder"): # Another header
                    break
                
                # Check if this is a data row (has Network Address)
                if candidate.get("Network Address"):
                    # If label matches, or if label is UNUSED (special case)
                    if candidate.get("carving_label") == target_label:
                        found_subnets.append(candidate)
            
            if found_subnets:
                # Use the exact start and end of the subnets we found
                start_ip = found_subnets[0]['Assigned IP Range'].split(' - ')[0]
                end_ip = found_subnets[-1]['Assigned IP Range'].split(' - ')[-1]
                row['Building'] = f"{carving_name} - {start_ip} - {end_ip}"
            else:
                # If no subnets were carved, use the calculated aggregate block range for debugging insight
                aggregate_network = row.get('aggregate_network')
                if aggregate_network:
                    first_ip = aggregate_network.network_address + 1
                    last_ip = aggregate_network.broadcast_address - 1
                    row['Building'] = f"{carving_name} - {first_ip} - {last_ip}"
                else:
                     row['Building'] = f"{carving_name} - Range Not Determined"


        # Clean up temporary keys and add to final results list
        cleaned_row = {k: v for k, v in row.items() if k not in ['is_range_placeholder', 'carving_name', 'allocation_start_index', 'base_block', 'aggregate_network', 'carving_label', 'target_count', 'target_cidr']}
        # final_results is undefined here, should use allocation_results directly but clean up in place
        pass # Doing in-place modification of 'row' dict, keys removed above.

    # Final Cleanup of carving_label in all rows
    final_results = []
    for row in allocation_results:
        cleaned = {k:v for k,v in row.items() if k != 'carving_label'}
        final_results.append(cleaned)
        
    return final_results, current_network_address

def create_summary_rows(master_network, results, next_ip):
    summary = []
    # Row 1
    m_range = f"{master_network.network_address + 1} - {master_network.broadcast_address - 1}"
    row1 = EMPTY_ROW_TEMPLATE.copy()
    row1["Building"] = f"{master_network} - {m_range}"
    summary.append(row1)
    summary.append(EMPTY_ROW_TEMPLATE.copy())
    return summary

def main():
    if len(sys.argv) < 2:
        print("Usage: python ip_planner.py <subnet>")
        sys.exit(1)
        
    start_net = sys.argv[1]
    res, _ = generate_ip_allocation(start_net, MASTER_PLAN)
    
    if res:
        summ = create_summary_rows(ipaddress.IPv4Network(start_net), res, _)
        final = summ + res
        df = pd.DataFrame(final)
        cols = ["Building", "Assigned Equipment", "Network Address", "Assigned IP Range", "CIDR", "Subnet Mask", "SVI GATEWAY", "Hosts"]
        df = df.reindex(columns=cols)
        
        # Fix for older pandas/xlsxwriter engines if needed
        try:
            with pd.ExcelWriter('ip_allocation_plan.xlsx', engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='IP Plan', index=False, header=False)
                ws = writer.sheets['IP Plan']
                
                # Define formats
                # Orange Bold with Border for Headers
                orange_bold = writer.book.add_format({'bold': True, 'bg_color': '#FFC000', 'font_color': 'black', 'text_wrap': True, 'border': 1})
                
                # Regular Text Wrap with Border for Data
                text_wrap = writer.book.add_format({'text_wrap': True, 'border': 1})
                
                # Auto-fit column widths
                for i, col in enumerate(df.columns):
                    column_data = df[col].astype(str)
                    max_len = max(column_data.map(len).max(), len(col)) + 2
                    ws.set_column(i, i, max_len)

                # Apply row formatting
                for idx, row_data in enumerate(final):
                    building_val = str(row_data.get("Building", ""))
                    
                    is_header = False
                    # Check for Range Headers or Column Headers
                    if " - " in building_val and not row_data.get("Network Address"):
                         is_header = True
                    elif building_val == "Building":
                        is_header = True
                    
                    # Loop through the first 8 columns (0-7)
                    for c in range(8):
                        col_name = cols[c]
                        val = row_data.get(col_name, "")
                        if pd.isna(val): val = ""
                        
                        if is_header:
                             ws.write(idx, c, val, orange_bold)
                        else:
                            # For Data Rows:
                            # Only apply border (text_wrap format) if the cell has content (val is not empty)
                            if val:
                                ws.write(idx, c, val, text_wrap)
                            else:
                                # Write empty string with default format (no border)
                                ws.write(idx, c, "", writer.book.add_format({'text_wrap': True})) 

        except Exception as e:
            print(e)
            df.to_csv('ip_allocation_plan.csv', index=False)
            
        print("Done. Saved to ip_allocation_plan.xlsx")

if __name__ == "__main__":
    main()
