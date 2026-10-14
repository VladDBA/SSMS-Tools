#Requires -RunAsAdministrator

<#

.SYNOPSIS
    Backs up or restores SSMS 22's saved connections (ConnectionMruList)
    by exporting from, or importing into, its privateregistry.bin hive.

.NOTES
    - Requires admin rights.
    - SSMS 22 must be closed.
    - -Restore overwrites the ConnectionMruList key in the SSMS 22 hive, and
      backs up the SSMS 22 privateregistry.bin file before making any changes.
    - A .reg file created from a different SSMS 22 build folder can be restored,
      the key path inside it is rewritten to match the current folder.
    - You cannot use this script to restore connections between different OS user profiles and/or different machines.

  Author: Vlad Drumea (VladDBA)
  Website: https://vladdba.com

.PARAMETER Backup
    Exports ConnectionMruList to the .reg file specified by -BackupFile.

.PARAMETER Restore
    Imports the .reg file specified by -BackupFile into the SSMS 22 hive,
    replacing the existing ConnectionMruList.

.PARAMETER BackupFile
    Full path of the .reg file to create (-Backup) or read (-Restore).

.PARAMETER OSUserName
    (Optional) Specify the OS user profile name to use for locating SSMS configuration folders.

.EXAMPLE
    PS C:\> .\Backup-Restore-SSMS22Connections.ps1 -Backup -BackupFile "D:\Backups\SSMS22_Connections.reg"

.EXAMPLE
    PS C:\> .\Backup-Restore-SSMS22Connections.ps1 -Restore -BackupFile "D:\Backups\SSMS22_Connections.reg"

.EXAMPLE
    PS C:\> .\Backup-Restore-SSMS22Connections.ps1 -Backup -BackupFile "D:\Backups\SSMS22_Connections.reg" -OSUserName "OtherUser"

#>

param(
    [switch]$Backup,
    [switch]$Restore,
    [string]$BackupFile,
    [string]$OSUserName
)

if ($Backup -eq $Restore) {
    throw "Specify exactly one of -Backup or -Restore."
}
if ([string]::IsNullOrWhiteSpace($BackupFile)) {
    throw "-BackupFile is required and must be the full path of the .reg file."
}
if (-not [System.IO.Path]::IsPathRooted($BackupFile)) {
    throw "-BackupFile must be a full path. Received: $BackupFile"
}
if ($Restore -and -not (Test-Path -Path $BackupFile -PathType Leaf)) {
    throw "Backup file not found: $BackupFile"
}

## internal variables
$BinFile = "privateregistry.bin"
$MountPoint = "HKEY_LOCAL_MACHINE\SSMSStuff"
$HiveLoaded = $false

$SSMSRoot = if (-not [string]::IsNullOrEmpty($OSUserName)) {
    Join-Path -Path "C:\Users\$OSUserName\AppData\Local" -ChildPath "Microsoft\SSMS"
} else {
    Join-Path -Path $env:LOCALAPPDATA -ChildPath "Microsoft\SSMS"
}
if (-not (Test-Path -Path $SSMSRoot)) {
    throw "The SSMS root directory ($SSMSRoot) does not exist. Ensure SSMS 22 is installed and, if used, that -OSUserName is correct."
}

while (Get-Process "SSMS" -ErrorAction SilentlyContinue | Where-Object { $_.FileVersion -match "^22\." }) {
    Write-Host " SSMS process detected. Please close all SSMS instances to proceed." -ForegroundColor Yellow
    Read-Host -Prompt "Press Enter after all SSMS 22 processes have been closed"
}

function Invoke-Reg {
    param(
        [Parameter(Mandatory)][string[]]$Arguments
    )
    $output = & reg.exe @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "reg $($Arguments -join ' ') failed (Exit $LASTEXITCODE): $($output -join ' ')"
    }
}

# GA (non-preview) SSMS 22 version folders only
$Folder22 = Get-ChildItem -Path $SSMSRoot -Directory |
    Where-Object { $_.Name -match "^22\.0_" } |
    Where-Object {
        $sdkFile = Join-Path -Path $_.FullName -ChildPath "sdk.txt"
        (Test-Path -Path $sdkFile -PathType Leaf) -and
        ((Get-Content -Path $sdkFile -Raw).Trim() -eq "UsePreviews=False")
    }
if (@($Folder22).Count -ne 1) {
    throw "Expected exactly one SSMS 22 folder, found $(@($Folder22).Count)."
}
$FName22 = $Folder22.Name
$Hive22 = Join-Path -Path $Folder22.FullName -ChildPath $BinFile
if (-not (Test-Path -Path $Hive22 -PathType Leaf)) {
    throw "Hive file not found: $Hive22"
}
Write-Host " Found hive file:`n  $Hive22" -ForegroundColor Green

$KeyPath = "$MountPoint\Software\Microsoft\SSMS\$FName22\ConnectionMruList"

if ($Restore) {
    Write-Host " This will overwrite SSMS 22's saved connections with those from:`n  $BackupFile" -ForegroundColor Yellow
    $confirmation = Read-Host -Prompt "Do you want to continue? (Y/N)"
    if ($confirmation -ne "Y") {
        Write-Host " Operation cancelled by user."
        exit
    }

    $RegContents = Get-Content -Path $BackupFile -Raw
    $KeyPattern = "(?i)HKEY_LOCAL_MACHINE\\[^\\\]]+\\Software\\Microsoft\\SSMS\\[^\\\]]+\\ConnectionMruList"
    if ($RegContents -notmatch "(?im)^\[$KeyPattern\]\s*$") {
        throw "$BackupFile does not look like a ConnectionMruList backup created by this script."
    }

    $HiveBackup = "$Hive22.$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
    Copy-Item -Path $Hive22 -Destination $HiveBackup -Force
    Write-Host " Backup created at:`n  $HiveBackup" -ForegroundColor Green

    $TempReg = Join-Path -Path $env:TEMP -ChildPath "SSMS22_ConnectionMruList_$([guid]::NewGuid().ToString('N')).reg"
}

try {
    Write-Host "`n Loading SSMS 22 hive..."
    Invoke-Reg -Arguments @("load", $MountPoint, $Hive22)
    $HiveLoaded = $true

    if ($Backup) {
        $BackupDir = Split-Path -Path $BackupFile -Parent
        if (-not (Test-Path -Path $BackupDir)) {
            New-Item -Path $BackupDir -ItemType Directory -Force | Out-Null
        }
        Write-Host " Exporting ConnectionMruList to $BackupFile..."
        Invoke-Reg -Arguments @("export", $KeyPath, $BackupFile, "/y", "/reg:64")
        Write-Host "`n Backup completed successfully." -ForegroundColor Green
    } else {
        # point the backup's key path at this machine's current SSMS 22 folder
        $RegContents = $RegContents -replace $KeyPattern, $KeyPath
        $RegContents | Set-Content -Path $TempReg -Encoding Unicode

        # reg import merges, so remove the existing key to get an overwrite
        & reg.exe query $KeyPath /reg:64 *> $null
        if ($LASTEXITCODE -eq 0) {
            Write-Host " Removing existing ConnectionMruList..."
            Invoke-Reg -Arguments @("delete", $KeyPath, "/f", "/reg:64")
        }

        Write-Host " Importing $BackupFile into SSMS 22 hive..."
        Invoke-Reg -Arguments @("import", $TempReg, "/reg:64")
        Write-Host "`n Restore completed successfully." -ForegroundColor Green
        Write-Host " Start SSMS 22 and verify that your connections have been restored."
        Write-Host " If you encounter any issues, close SSMS 22 and replace $Hive22 with $HiveBackup."
    }
} catch {
    Write-Host " Something went wrong: $_" -ForegroundColor Red
} finally {
    if ($TempReg -and (Test-Path -Path $TempReg)) {
        Remove-Item -Path $TempReg -Force
    }
    if ($HiveLoaded) {
        Write-Host " Unloading SSMS 22 hive..."
        [System.GC]::Collect()
        try {
            Invoke-Reg -Arguments @("unload", $MountPoint)
        } catch {
            Write-Host " Failed to unload hive: $_" -ForegroundColor Red
            Write-Host " You may need to unload $MountPoint manually using Registry Editor." -ForegroundColor Red
        }
    }
}
