from getpass import getpass
import netmiko
import re

def make_connection(ip, username, password):
    return netmiko.ConnectHandler(device_type='cisco_ios', ip=ip, username=username, password=password)

def get_ip(input):
    return(re.findall(r'(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?).){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)', input))

def get_ips(file_name):
    for line in open(file_name, 'r').readlines():
        line = get_ip(line)
    for ip in line:
        ips.append(ip)

def to_doc_a(file_name, variable):
    with open(file_name, 'a') as f:
        f.write(variable)
        f.write('n')

def to_doc_w(file_name, variable):
    with open(file_name, 'w') as f:
        f.write(variable)

ips = []

get_ips("C:\IPs.txt")

username = raw_input("Username: ")
password = getpass()
file_name = "C:\results.txt"

to_doc_w(file_name, "")

for ip in ips:
    net_connect = make_connection(ip, username, password)

    output = net_connect.send_command_expect('show run')
    to_doc_a(file_name, output)