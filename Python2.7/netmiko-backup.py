from getpass import getpass
from netmiko import ConnectHandler
import re

def make_connection(ip, username, password):
    return ConnectHandler(device_type='cisco_ios', ip=ip, username=username, password=password)

def get_ip(input):
    return(re.findall(r'(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?).){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)', input))

def get_ips(file_name):
    for line in open(file_name, 'r').readlines():
        line = get_ip(line)
        for ip in line:
            ips.append(ip)

ips = []

get_ips("C:\\Path\\to\\file\\IPs.txt")

username = raw_input("Username: ")
password = getpass()

for ip in ips:
    net_connect = make_connection(ip, username, password)
    hostname = net_connect.send_command_expect('show run | section hostname')

    # print "Backing up " + hostname[9:] + "."
    dest_file = open("C:\\Path\\to\\file\\" + hostname[9:] + ".txt", 'w')

    output = net_connect.send_command_expect('show run')
    dest_file.writelines(output)
    dest_file.close