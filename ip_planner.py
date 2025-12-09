import ipaddress
import sys
from typing import List, Dict, Any, Tuple
import pandas as pd

# --- Configuration for Subnetting Scheme ---
NUM_BLDGS_BASE = 10 
NUM_SPECIAL = 1 
NUM_RESERVED = 2
NUM_ALLOCATIONS = NUM_BLDGS_BASE + NUM_RESERVED 

# --- FINAL COLUMNS TO DISPLAY ---
COLUMN_HEADERS = {
    "Assigned Equipment": "Assigned Equipment", 
    "Assigned IP Range": "Assigned IP Range"
}

# --- The MASTER_PLAN ---
MASTER_PLAN = [
    # =========================================================================
    # TABLE 1: Switch Management (Block A)
    # =========================================================================
    {"phase": 2, "type": "header_block", "block_name": "Block A", "purpose": "Switch Management /24s"},
    {"phase": 2, "type": "allocation", "target_block_indices": [0], "purpose": "Switch Management", "cidr": 24, "count": NUM_ALLOCATIONS},
    
    {"phase": 2, "type": "blank_row"}, 

    # =========================================================================
    # TABLE 2: Cameras (Block A Remainder + B + C + D Overflow)
    # =========================================================================
    {"phase": 2, "type": "header_block", "block_name": "Block A/B/C & D (Overflow)", "purpose": "Cameras /21s"},
    {"phase": 2, "type": "allocation", "target_block_indices": [0, 1, 2, 3], "purpose": "Cameras", "cidr": 21, "count": NUM_ALLOCATIONS},
    # Only carve unused from A/B/C here to save D for others
    {"phase": 2, "type": "carve_remainder", "target_block_indices": [0, 1, 2], "purpose": "UNUSED (Camera Blocks Remainder)", "cidr": 21},

    {"phase": 2, "type": "blank_row"}, 

    # =========================================================================
    # TABLE 3: PIDS (Block D)
    # =========================================================================
    {"phase": 2, "type": "header_block", "block_name": "Block D", "purpose": "PIDS /24s"},
    {"phase": 2, "type": "allocation", "target_block_indices": [3], "purpose": "PIDS", "cidr": 24, "count": NUM_ALLOCATIONS},

    {"phase": 2, "type": "blank_row"}, 

    # =========================================================================
    # TABLE 4: Intercoms (Block D)
    # =========================================================================
    {"phase": 2, "type": "header_block", "block_name": "Block D", "purpose": "Intercoms /27s"},
    {"phase": 2, "type": "allocation", "target_block_indices": [3], "purpose": "Intercoms", "cidr": 27, "count": NUM_ALLOCATIONS},

    {"phase": 2, "type": "blank_row"}, 

    # =========================================================================
    # TABLE 5: Servers (Block D)
    # =========================================================================
    {"phase": 2, "type": "header_block", "block_name": "Block D", "purpose": "Servers /27s"},
    {"phase": 2, "type": "allocation", "target_block_indices": [3], "purpose": "Servers", "cidr": 27, "count": NUM_ALLOCATIONS},
    
    {"phase": 2, "type": "blank_row"}, 

    # =========================================================================
    # TABLE 6: Unused Block D Remainder
    # =========================================================================
    {"phase": 2, "type": "header_block", "block_name": "Block D", "purpose": "UNUSED / RESERVED SPACE"},
    {"phase": 2, "type": "carve_remainder", "target_block_indices": [3], "purpose": "UNUSED (Block D Remainder)", "cidr": 24},
]

EMPTY_ROW_TEMPLATE = {
    "Assigned Equipment": "", 
    "Assigned IP Range": ""
}

def format_allocation_row(network: ipaddress.IPv4Network, purpose: str) -> Dict[str, Any]:
    """Formats a single allocated network row."""
    if network.prefixlen >= 31:
        ip_range = f"{network.network_address} - {network.broadcast_address}" 
    else:
        first_usable = str(network.network_address + 1)
        last_usable = str(network.broadcast_address - 1)
        ip_range = f"{first_usable} - {last_usable}"

    return {
        "Assigned Equipment": purpose,
        "Assigned IP Range": ip_range
    }

def generate_ip_allocation(master_network_str: str, plan: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], ipaddress.IPv4Address]:
    try:
        master_network = ipaddress.IPv4Network(master_network_str)
    except ipaddress.AddressValueError:
        print(f"Error: Invalid IP network format.")
        return [], ipaddress.IPv4Address(master_network_str.split('/')[0])

    # --- BLOCK SEGMENTATION LOGIC ---
    try:
        blocks = list(master_network.subnets(prefixlen_diff=2))
    except ValueError:
        print("Error: Master network is too small to split into 4 blocks.")
        return [], master_network.network_address

    block_start_ips = [b.network_address for b in blocks]
    block_end_ips = [b.broadcast_address for b in blocks]
    current_pointers = list(block_start_ips) 
    
    allocation_results = []
    
    print(f"\n--- Starting 4-Block Segmentation from {master_network_str} ---")
    
    # --- PROCESS PLAN ---
    for item in plan:
        item_type = item.get("type")
        
        if item_type == "blank_row":
            allocation_results.append(EMPTY_ROW_TEMPLATE.copy())
            continue
        
        if item_type == "header_block":
            h1 = EMPTY_ROW_TEMPLATE.copy()
            h1["Assigned Equipment"] = f"--- {item['block_name']} ({item['purpose']}) ---"
            
            # Dynamic range calculation for display
            if "Block A" in item['block_name'] and "B/C" in item['block_name']:
                 range_str = f"Remainder of A + B + C + D"
            elif "Block A" in item['block_name']:
                 range_str = str(blocks[0])
            elif "Block D" in item['block_name']:
                 range_str = str(blocks[3])
            else:
                 range_str = "Range"
            
            h1["Assigned IP Range"] = range_str
            allocation_results.append(h1)
            allocation_results.append(COLUMN_HEADERS.copy())
            continue

        target_indices = item.get("target_block_indices")
        
        if item_type == "allocation":
            cidr = item['cidr']
            count = item['count']
            purpose = item['purpose']
            
            for i in range(count):
                allocated_net = None
                
                for b_idx in target_indices:
                    ptr = current_pointers[b_idx]
                    limit = block_end_ips[b_idx]
                    
                    try:
                        candidate = ipaddress.ip_network(f"{ptr}/{cidr}", strict=False)
                        if candidate.network_address < ptr:
                            while candidate.network_address < ptr:
                                candidate = ipaddress.ip_network(f"{candidate.broadcast_address + 1}/{cidr}", strict=False)

                        if candidate.broadcast_address <= limit:
                            allocated_net = candidate
                            current_pointers[b_idx] = candidate.broadcast_address + 1
                            break 
                    except ValueError:
                        continue 
                
                if allocated_net:
                    is_reserved = (i >= NUM_BLDGS_BASE)
                    p_final = f"{purpose} (Reserved)" if is_reserved else purpose
                    row = format_allocation_row(allocated_net, p_final)
                    allocation_results.append(row)
                else:
                    row = EMPTY_ROW_TEMPLATE.copy()
                    row["Assigned Equipment"] = f"ERROR: No space left for {purpose}"
                    allocation_results.append(row)
                    break

        elif item_type == "carve_remainder":
            cidr = item['cidr']
            purpose = item['purpose']
            
            for b_idx in target_indices:
                ptr = current_pointers[b_idx]
                limit = block_end_ips[b_idx]
                
                while ptr <= limit:
                    try:
                        sub = ipaddress.ip_network(f"{ptr}/{cidr}", strict=False)
                        if sub.network_address < ptr:
                             while sub.network_address < ptr:
                                sub = ipaddress.ip_network(f"{sub.broadcast_address + 1}/{cidr}", strict=False)

                        if sub.broadcast_address <= limit:
                            row = format_allocation_row(sub, purpose)
                            allocation_results.append(row)
                            ptr = sub.broadcast_address + 1
                        else:
                            break
                    except ValueError:
                        break
                current_pointers[b_idx] = ptr

    return allocation_results, master_network.network_address

def create_summary_rows(master_network, results, next_ip):
    summary = []
    m_range = f"{master_network.network_address + 1} - {master_network.broadcast_address - 1}"
    
    row1 = EMPTY_ROW_TEMPLATE.copy()
    row1["Assigned Equipment"] = "MASTER ALLOCATION PLAN"
    row1["Assigned IP Range"] = m_range
    summary.append(row1)
    summary.append(EMPTY_ROW_TEMPLATE.copy())
    return summary

def main():
    if len(sys.argv) < 2:
        print("Usage: python ip_planner.py <subnet>")
        print("Example: python ip_planner.py 10.15.0.0/17")
        sys.exit(1)
        
    start_net = sys.argv[1]
    res, _ = generate_ip_allocation(start_net, MASTER_PLAN)
    
    if res:
        summ = create_summary_rows(ipaddress.IPv4Network(start_net), res, _)
        final = summ + res
        df = pd.DataFrame(final)
        
        # --- 2 Columns Only ---
        cols = ["Assigned Equipment", "Assigned IP Range"]
        df = df.reindex(columns=cols)
        
        try:
            with pd.ExcelWriter('ip_allocation_plan.xlsx', engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='IP Plan', index=False, header=False)
                ws = writer.sheets['IP Plan']
                
                # Define formats
                orange_bold = writer.book.add_format({'bold': True, 'bg_color': '#FFC000', 'font_color': 'black', 'text_wrap': True, 'border': 1})
                text_wrap_border = writer.book.add_format({'text_wrap': True, 'border': 1})
                text_wrap_no_border = writer.book.add_format({'text_wrap': True}) # For blank rows
                
                for i, col in enumerate(df.columns):
                    column_data = df[col].astype(str)
                    max_len = max(column_data.map(len).max(), len(col)) + 2
                    ws.set_column(i, i, max_len)

                for idx, row_data in enumerate(final):
                    equip_val = str(row_data.get("Assigned Equipment", ""))
                    
                    is_header = False
                    if "--- Block" in equip_val or equip_val == "Assigned Equipment":
                         is_header = True
                    elif equip_val == "MASTER ALLOCATION PLAN":
                        is_header = True
                    
                    for c in range(len(cols)):
                        col_name = cols[c]
                        val = row_data.get(col_name, "")
                        if pd.isna(val): val = ""
                        
                        if is_header:
                             ws.write(idx, c, val, orange_bold)
                        else:
                            # If the cell has data, give it a border.
                            # If the cell is empty (blank row), NO border.
                            if val:
                                ws.write(idx, c, val, text_wrap_border)
                            else:
                                ws.write(idx, c, "", text_wrap_no_border) 

        except Exception as e:
            print(e)
            df.to_csv('ip_allocation_plan.csv', index=False)
            
        print("Done. Saved to ip_allocation_plan.xlsx")

if __name__ == "__main__":
    main()
