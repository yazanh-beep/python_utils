import json
import openpyxl
from collections import defaultdict

def load_json_data(json_file):
    """Load the network topology JSON file."""
    with open(json_file, 'r') as f:
        return json.load(f)

def is_aggregate_switch(hostname):
    """Check if a switch is an aggregate switch based on hostname."""
    return 'SMSAGG' in hostname.upper()

def clean_hostname(hostname):
    """
    Clean hostname by removing .CAM.INT or other domain suffixes if present.
    Returns the base hostname.
    """
    if not hostname:
        return hostname
    # Remove common domain suffixes
    for suffix in ['.CAM.INT', '.cam.int', '.local']:
        if hostname.endswith(suffix):
            return hostname[:-len(suffix)]
    return hostname

def format_aggregate_uplinks(neighbors):
    """
    Format aggregate switch uplinks intelligently.
    Handles multiple uplinks to aggregate switches.
    """
    agg_connections = defaultdict(list)
    
    # Group connections by aggregate switch
    for neighbor in neighbors:
        neighbor_hostname = neighbor.get('neighbor_hostname', '')
        if is_aggregate_switch(neighbor_hostname):
            # Use the cleaned full hostname as the key and display name
            clean_name = clean_hostname(neighbor_hostname)
            
            connection_info = {
                'local_port': neighbor.get('local_interface', 'Unknown'),
                'remote_port': neighbor.get('remote_interface', 'Unknown'),
                'agg_full_name': clean_name
            }
            agg_connections[clean_name].append(connection_info)
    
    if not agg_connections:
        return "", ""
    
    # Format the output
    agg_names = []
    uplink_details = []
    
    for agg_hostname, connections in sorted(agg_connections.items()):
        agg_names.append(agg_hostname)
        
        if len(connections) == 1:
            # Single connection
            conn = connections[0]
            uplink_details.append(f"On {agg_hostname}: Local {conn['local_port']} -> Remote {conn['remote_port']}")
        else:
            # Multiple connections to same aggregate
            ports_info = []
            for idx, conn in enumerate(connections, 1):
                ports_info.append(f"Link {idx}: Local {conn['local_port']} -> Remote {conn['remote_port']}")
            uplink_details.append(f"On {agg_hostname}: {'; '.join(ports_info)}")
    
    aggregate_switch_str = " and ".join(agg_names)
    uplink_port_str = " | ".join(uplink_details)
    
    return aggregate_switch_str, uplink_port_str

def format_aggregate_to_aggregate_uplinks(neighbors, current_hostname):
    """
    Special formatting for aggregate switches connecting to each other.
    Shows the peer aggregate switch and the ports on both sides.
    """
    agg_connections = []
    
    # Clean the current hostname for comparison
    clean_current = clean_hostname(current_hostname)
    
    for neighbor in neighbors:
        neighbor_hostname = neighbor.get('neighbor_hostname', '')
        clean_neighbor = clean_hostname(neighbor_hostname)
        
        if is_aggregate_switch(neighbor_hostname) and clean_neighbor != clean_current:
            local_port = neighbor.get('local_interface', 'Unknown')
            remote_port = neighbor.get('remote_interface', 'Unknown')
            
            agg_connections.append({
                'peer': clean_neighbor,
                'local_port': local_port,
                'remote_port': remote_port
            })
    
    if not agg_connections:
        return "", ""
    
    # Group by peer hostname
    peer_groups = defaultdict(list)
    for conn in agg_connections:
        peer_groups[conn['peer']].append(conn)
    
    agg_names = []
    uplink_details = []
    
    for peer_hostname, connections in sorted(peer_groups.items()):
        agg_names.append(peer_hostname)
        
        if len(connections) == 1:
            conn = connections[0]
            uplink_details.append(f"{peer_hostname}: Local {conn['local_port']} <-> Remote {conn['remote_port']}")
        else:
            ports_info = []
            for idx, conn in enumerate(connections, 1):
                ports_info.append(f"Link {idx}: Local {conn['local_port']} <-> Remote {conn['remote_port']}")
            uplink_details.append(f"{peer_hostname}: {'; '.join(ports_info)}")
    
    aggregate_switch_str = " and ".join(agg_names)
    uplink_port_str = " | ".join(uplink_details)
    
    return aggregate_switch_str, uplink_port_str

def populate_excel_tracker(json_file, excel_file, output_file):
    """
    Populate the Excel tracker with data from the JSON file.
    """
    # Load data
    topology_data = load_json_data(json_file)
    
    # Load Excel workbook
    wb = openpyxl.load_workbook(excel_file)
    
    # Access the 'switch' sheet
    if 'switch' not in wb.sheetnames:
        print("Error: 'switch' sheet not found in workbook")
        return
    
    ws = wb['switch']
    
    # Expected headers (as seen in the file)
    expected_headers = [
        "Switch name",
        "Serial Number",
        "Management IP address",
        "switch model",
        "firmware",
        "aggregate switch",
        "uplink port on aggregate switch"
    ]
    
    # Start writing from row 2 (row 1 has headers)
    current_row = 2
    
    for device in topology_data:
        hostname = device.get('hostname', '')
        serial_number = device.get('serial_number', '')
        mgmt_ip = device.get('management_ip', '')
        switch_model = device.get('switch_model', '')
        ios_version = device.get('ios_version', '')
        neighbors = device.get('neighbors', [])
        
        # Determine aggregate switch and uplink information
        if is_aggregate_switch(hostname):
            # This is an aggregate switch
            aggregate_switch, uplink_port = format_aggregate_to_aggregate_uplinks(neighbors, hostname)
        else:
            # This is a regular switch
            aggregate_switch, uplink_port = format_aggregate_uplinks(neighbors)
        
        # Write to Excel
        ws.cell(row=current_row, column=1, value=hostname)
        ws.cell(row=current_row, column=2, value=serial_number)
        ws.cell(row=current_row, column=3, value=mgmt_ip)
        ws.cell(row=current_row, column=4, value=switch_model)
        ws.cell(row=current_row, column=5, value=ios_version)
        ws.cell(row=current_row, column=6, value=aggregate_switch)
        ws.cell(row=current_row, column=7, value=uplink_port)
        
        current_row += 1
    
    # Save the workbook
    wb.save(output_file)
    print(f"Successfully populated {current_row - 2} switches in {output_file}")

if __name__ == "__main__":
    # File paths
    json_file = "network_topology.json"
    excel_file = "PHXMSA1Acamera_switch_tracker.xlsx"
    output_file = "PHXMSA1Acamera_switch_tracker_populated.xlsx"
    
    # Populate the tracker
    populate_excel_tracker(json_file, excel_file, output_file)
    
    print("\nDone! Check the output file for results.")
