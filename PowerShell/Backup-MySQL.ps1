<#
    Script Name:  Backup-MySQL.ps1
    Created By:  Rob Rathbun
    Created Date:  01/25/2017
    Purpose:  Backup the MySQL database to a file.

    Inspired by the following script:
        http://www.fluxbytes.com/powershell/using-powershell-to-backup-your-mysql-databases/

    Make sure to change the variables on lines 24-43 to match your environment.
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
            Rename-Item -Path $logFileFullPath -NewName ($logFile + ".old") | Out-Null; #Rename existing logfile.
            New-Item $logFileFullPath -ItemType File | Out-Null; #Create a new log file.
        }
    }

Out-File $logFileFullPath -InputObject ([string]::Format("{0} : Starting backup", (Get-Date -Format g))) -Append; #Starting line of log file entry

# Process all databases
foreach ($database in $DbArrayFinal) {
    try {
        Out-File $logFileFullPath -InputObject ([string]::Format("`tBacking up database name `"{0}`".", $database.DbName)) -Append;
        If (!(Test-Path $database.DbBackupPath)) {
            New-Item $database.DbBackupPath -ItemType Directory | Out-Null; #Create backup folder if it doesn't exist.
        }

        #Retain one week's copies (7days - today = 6).
        Get-ChildItem $database.DbBackupPath | sort CreationTime -Descending | select -Skip 6 | Remove-Item -Force

        $backupDate = (Get-Date -Format yyyy-MM-dd).ToString(); #Backup Date format to be used in file naming.
        $savePath = [string]::Format("{0}\{1}_{2}.sql", $database.DbBackupPath, $database.DbName, $backupDate); #Full savefile name and path.
        
        $command = [string]::Format("`"{0}`" -u {1} -p {2} -h {3} --quick --default-character-set=utf8 --routines --events `"{4}`" > `"{5}`"",
            $mySQLDumpLocation,
            $dbUser,
            $dbPassword,
            $serverIP,
            $database.DbName,
            $savePath
        ); #Formated string of mysqldump.exe command to be run.

        $dumpError;
        Invoke-Expression "& $command" -ErrorVariable dumpError; #Execute backup command.

        # Log backup job completion.
        If ($dumpError -eq $null) {
            Out-File $logFileFullPath -InputObject ([string]::Format("`tBackup of database `"{0}`" completed with no error codes", $database.DbName)) -Append;
        } Else {
            Out-File $logFileFullPath -InputObject ([string]::Format("`tBackup of database `"{0}`" completed with error: `"{1}`"", $database.DbName, $dumpError)) -Append;
        }

        #Add backup to *.7z compressed archive.
        $7zFileFullPath = [string]::Format("{0}\{1}_{2}.7z", $database.DbBackupPath, $database.DbName, $backupDate);
        $7zCommand = [string]::Format("`"{0}`" a -t7z `"{1}`" -i!`"{2}`"", $7zipPath, $7zFileFullPath, $savePath);
        $compressError;
        Invoke-Expression "& $7zCommand" -ErrorVariable compressError | Out-Null;

        # Log compression results.  Delete *.sql file if no errors.
        If ($compressError -eq $null) {
            Out-File $logFileFullPath -InputObject ([string]::Format("`tFile `"{0}`" compressed into archive `"{1}`".", $savePath, $7zFileFullPath)) -Append;
            Remove-Item $savePath -Force;
            Out-File $logFileFullPath -InputObject ([string]::Format("`tFile `"{0}`" deleted.", $savePath)) -Append;
        } Else {
            Out-File $logFileFullPath -InputObject "`tCompression failed with error code: $compressError" -Append;
            If (Test-Path $7zFileFullPath) {Remove-Item $7zFileFullPath -Force;}
        }
    }
    Catch [Exception] {
        #Write exception to log file.
        $logEntry = [string]::Format("`tFailed to start backup of database name `"{0}`". Reason: {1}", $database.DbName, $_.Exception.Message);
        Out-File $logFileFullPath -InputObject $logEntry -Append;
    }
}

Out-File $logFileFullPath -InputObject ([string]::Format("{0} : Backup Completed", (Get-Date -Format g))) -Append;