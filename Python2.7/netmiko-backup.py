from getpass import getpass
from netmiko import ConnectHandler
import re

password = getpass()

def make_connection(ip, username, password):
    return net_connect = netmiko.ConnectHandler(device_type='cisco_ios', ip=ip, username=username, password=password)

def get_ip (input):
    return(re.findall(r'(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?).){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)', input))

def get_ips (file_name):
    for line in open(file_name, 'r').readlines():
        line = get_ip(line)
    for ip in line:
        ips.append(ip)

def to_doc_a(file_name, variable):
    with open(file_name, 'a') as f:
        f.write(variable)
        f.write('n')

