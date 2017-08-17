# Script using Netmiko library to backup
# Cisco IOS devices.
# Created by:  Rob Rathbun
# Date Create:  August 17, 2017

from getpass import getpass
from netmiko import ConnectHandler
import re

# Function to make connection to each IOS device.
def make_connection(ip, username, password):
    return ConnectHandler(device_type='cisco_ios', ip=ip, username=username, password=password)

# Use Regex to compare data from input to select the IP address.
def get_ip(input):
    return(re.findall(r'(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?).){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)', input))

# Function to take input in the form of a text file and select the IP address(s) from each line.
def get_ips(file_name):
    for line in open(file_name, 'r').readlines():
        line = get_ip(line)
        for ip in line:
            ips.append(ip)

# Empty array of IP addresses that will be populated by the folowing line.
ips = []

# Call to get_ips function to populate array of IP addresses.
get_ips("C:\\Path\\to\\file\\IPs.txt")

# Obtain credentials.
username = raw_input("Username: ")
password = getpass()

# Iterate through each IP address in the 'ips' array.
for ip in ips:
    # Call 'make_connection' function to establish connection to the device.
    net_connect = make_connection(ip, username, password)
    # Pull the hostname from the config.  The hostname will be used as the filename for the backup.
    hostname = net_connect.send_command_expect('show run | section hostname')

    # Open/Create the destination file.  Truncate the section returned for 'hostname' so that only the actual hostname is populated.
    dest_file = open("C:\\Path\\to\\file\\" + hostname[9:] + ".txt", 'w')

    # Obtain the data to be backed up and write it to the 'dest_file'.
    output = net_connect.send_command_expect('show run')
    dest_file.writelines(output)
    dest_file.close