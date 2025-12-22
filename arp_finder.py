import sys
import subprocess
import os
import re

def get_system_arp_table():
    """Retrieves the current ARP table from the Linux system."""
    arp_map = {}
    try:
        # Run 'ip neigh' to get the ARP table
        output = subprocess.check_output("ip neigh", shell=True).decode("utf-8")
        
        # Parse the output line by line
        for line in output.splitlines():
            parts = line.split()
            if len(parts) >= 5:
                # Format: 192.168.1.5 dev eth0 lladdr 00:11:22:33:44:55 REACHABLE
                ip_addr = parts[0]
                mac_addr = parts[4]
                # Store in dictionary (lowercase for consistent matching)
                arp_map[mac_addr.lower()] = ip_addr
    except Exception as e:
        print(f"Error reading ARP table: {e}")
        sys.exit(1)
    return arp_map

def main():
    # Check if filename is provided
    if len(sys.argv) < 2:
        print("Usage: python3 mac_lookup_file.py <filename>")
        print("Example: python3 mac_lookup_file.py macs.txt")
        sys.exit(1)

    filename = sys.argv[1]

    if not os.path.isfile(filename):
        print(f"Error: File '{filename}' not found.")
        sys.exit(1)

    # Read the file
    with open(filename, 'r') as f:
        content = f.read()

    # Extract MACs using Regex (Robust against extra spaces/newlines)
    # This finds any standard MAC address pattern in the file
    mac_list = re.findall(r'([0-9A-Fa-f]{2}(?::[0-9A-Fa-f]{2}){5})', content)

    if not mac_list:
        print("No valid MAC addresses found in the file.")
        sys.exit(1)

    # Get current ARP table
    arp_table = get_system_arp_table()

    # Print Header
    print(f"{'MAC ADDRESS':<20} | {'IP ADDRESS'}")
    print("-" * 35)

    found_count = 0
    
    # Match and Print
    for mac in mac_list:
        normalized_mac = mac.lower()
        ip = arp_table.get(normalized_mac, "Not Found")
        
        if ip != "Not Found":
            found_count += 1
            
        print(f"{mac:<20} | {ip}")

    print("-" * 35)
    print(f"Total MACs in file: {len(mac_list)}")
    print(f"IPs found: {found_count}")

    if found_count == 0:
        print("\n[!] No IPs found. Remember to populate your ARP cache first:")
        print("    sudo nmap -sn 192.168.1.0/24")

if __name__ == "__main__":
    main()
