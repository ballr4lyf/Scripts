# Script using paramiko library to backup
# HPE v1910 series switches.
# Created by:  Rob Rathbun
# Date Create:  May 8, 2017

import paramiko
import time

# Update these variables to match your environment:
switchIP = '192.168.1.2'
username = 'backup_user'
password = 'backupUserPassw0rd'

# Create SSH client
ssh = paramiko.SSHClient()

# Automatically add untrusted SSH hosts.
ssh.set_missing_host_key_policy(
    paramiko.AutoAddPolicy())

# Connect to switch
ssh.Connect(switchIP, username=username, password=password, look_for_keys=False, allow_agent=False)

# Initiate interactive session with switch.
connection = ssh.invoke_shell()
