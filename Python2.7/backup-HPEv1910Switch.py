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
destination = '/my/backup/destination'

# Create SSH client
ssh = paramiko.SSHClient()

# Automatically add untrusted SSH hosts.
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

# Connect to switch
ssh.connect(switchIP, username=username, password=password, look_for_keys=False, allow_agent=False)

# Initiate interactive session with switch.
connection = ssh.invoke_shell()

# Enable advanced shell.
connection.send('_cmdline-mode on\n')
connection.send('Y\n')
connection.send('512900\n')

# Disable paging.
connection.send('system-view\n')
connection.send('user-interface vty 0 15\n')
connection.send('screen-length 0\n')
connection.send('quit\n')
connection.send('quit\n')

# Show current config.
connection.send('display current-configuration\n')

# Wait to complete.
time.sleep(2)

# Save output to temporary file.
output = connection.recv(65535)
tempfile = open(destination + '/switch_' + switchIP + '.config', 'w')
tempfile.write(output)
tempfile.write('\n')
tempfile.close

# Prune extra lines from config.
lines = open(destination + '/switch_' + switchIP + '.config', 'r').readlines()
outfile = open(destination + '/switch_' + switchIP + '.config', 'w')
outfile.writelines(lines[18:-1])    # By default, the config starts at line 18.
outfile.close

# Reset paging.
connection.send('system-view\n')
connection.send('user-interface vty 0 15\n')
connection.send('screen-length 24\n')
connection.send('quit\n')
connection.send('quit\n')
connection.close
