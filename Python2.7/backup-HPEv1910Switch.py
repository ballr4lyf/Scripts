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
destination = '\my\backup\destination\file.config'

# Create SSH client
ssh = paramiko.SSHClient()

# Automatically add untrusted SSH hosts.
ssh.set_missing_host_key_policy(
    paramiko.AutoAddPolicy())

# Connect to switch
ssh.Connect(switchIP, username=username, password=password, look_for_keys=False, allow_agent=False)

# Initiate interactive session with switch.
connection = ssh.invoke_shell()

# Enable advanced shell.
connection.send('_cmdline-mode on\n')
connection.send('Y\n')
connection.send('512900\n')

# Disable paging.
connection.send('system-view\n')
connection.send('user-interface vty 0 15\n')
connection.send('screen-lenght 0\n')
connection.send('quit\n')
connection.send('quit\n')

# Show current config.
connection.send('display current-configuration\n')

# Wait to complete.
time.sleep(2)

# Reset paging.
connection.send('system-view\n')
connection.send('user-interface vty 0 15\n')
connection.send('screen-lenght 24\n')
connection.send('quit\n')
connection.send('quit\n')