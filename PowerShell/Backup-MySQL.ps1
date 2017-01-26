<#
    Script Name:  Backup-MySQL.ps1
    Created By:  Rob Rathbun
    Created Date:  01/25/2017
    Purpose:  Backup the MySQL database to a file.

    Inspired by the following script:
        http://www.fluxbytes.com/powershell/using-powershell-to-backup-your-mysql-databases/
#>
Class DB {
    [string] $DbName;
    [string] $DbBackupPath;
    [string] $DbFileName;

    DB([string] $DbName, $DbBackupPath, $DbFileName) {
        $this.DbName = $DbName;
        $this.DbBackupPath = $DbBackupPath;
        $this.DbFileName = $DbFileName;
    }
}

$7zipPath = "C:\Program Files\7-Zip\7z.exe"; #Path to 7-zip.
$mySQLDumpLocation = "C:\Program Files\MySQL\MySQL Server 5.6\bin\mysqldump.exe"; #Path to the mysqldump.exe location.

$backupFolder = "C:\MySQL_Backups"; #Folder location where backups will be stored.
$logFile = "MySQL_backups.log"; #File name for the backup log.
$logFileFullPath = [io.path]::Combine($backupFolder, $logFile); #Full path to the backup log.

$serverIP = "xxx.xxx.xxx.xxx"; #IP address (or FQDN) of the MySQL server.
$dbUser = "mysql.username"; #DB Username with access to all DBs being backed up.
$dbPassword = "P@ssw0rd"; #Password for the above user account.

<# In the line following this comment, enter in the name of each database to be backed up.

   Input each database name between the parentheses (e.g. ( )).
   Each database name needs to be surrounded by quotes (e.g. "database1").
   Each database name needs to be separated by a comma (,).
   There is no trailing comma after the last database name.
#>

$DbArrayInit = @("Database1", "Database2","Database3");

# Add DBs to the array that will be used programatically.
$DbArrayFinal = @()
foreach ($namedDB in $DbArrayInit) {
    [void]$DbArrayFinal.Add([DB]::New($namedDB, [io.path]::Combine($backupFolder, $namedDB), $namedDB))
}

# Check for pre-existing log file.  If none exists, create one.
    If (!(Test-Path $backupFolder)) {
        New-Item $backupFolder -ItemType Directory | Out-Null
    }

    # Check log file for existence and size constraints.
    If (!(Test-Path $logFileFullPath)) {
        New-Item $logFileFullPath -ItemType File | Out-Null; #Create file if it does not exist.
    } Elseif ((Get-Item $logFileFullPath).Length -gt 50mb) {
        If (Test-Path ($logFileFullPath + ".old")) {
            Remove-Item ($logFileFullPath + ".old") -Force | Out-Null; #Check for and remove old backup logfile over 50mb.
        }
        Rename-Item -Path $logFileFullPath -NewName ($logFile + ".old") | Out-Null; #Rename existing logfile.
        New-Item $logFileFullPath -ItemType File | Out-Null; #Create a new log file.
    }


