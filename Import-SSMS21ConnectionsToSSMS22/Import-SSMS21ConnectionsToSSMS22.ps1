#Requires -RunAsAdministrator

<#

.SYNOPSIS
    Dynamically locate the SSMS version folders whose sdk.txt contains
    "UsePreviews=False", then import the ConnectionMruList key from SSMS 21's
    privateregistry.bin file to the SSMS 22 one.

.NOTES
    - Requires admin rights.
    - Works only if both SSMS 21 and SSMS 22 are installed side-by-side (aka on the same machine).
    - Backs up the SSMS 22 privateregistry.bin file before making any changes.

  Author: Vlad Drumea (VladDBA)
  Website: https://vladdba.com
  Date: November 17, 2025

.LINK
    For more information, visit: https://vladdba.com/2025/11/17/import-saved-connections-from-ssms-21-to-ssms-22/  

.PARAMETER OSUserName
    (Optional) Specify the OS user profile name to use for locating SSMS configuration folders.
    
.EXAMPLE
    PS C:\> .\Import-SSMS21ConnectionsToSSMS22.ps1
    Runs the script to import saved connections from SSMS 21 to SSMS 22.

.EXAMPLE
    PS C:\> .\Import-SSMS21ConnectionsToSSMS22.ps1 -OSUserName "OtherUser"
    Runs the script to import saved connections from SSMS 21 to SSMS 22
    for the specified OS user profile.    

#>


param(
    [Parameter(Position = 0, Mandatory = $False)]
    [string]$OSUserName
)

Write-Host " This will overwrite SSMS 22's saved connections with those from SSMS 21." -Fore Yellow
Write-Host " Make sure both SSMS 21 and 22 are closed before proceeding." -Fore Yellow
# confirm with the user before proceeding
$confirmation = Read-Host -Prompt "Do you want to continue? (Y/N)"
if ($confirmation -ne "Y") {
    Write-Host " Operation cancelled by user."
    exit
}

## internal variables
$BinFile = 'privateregistry.bin'
$RegFileName = 'SSMS21_ConnectionMruList.reg'
# both hives will be mounted under this mount point
$MountPoint = 'HKEY_LOCAL_MACHINE\SSMSStuff'
$MountPointTest = 'HKLM:\SSMSStuff'

# where SSMS 21 and 22 related configuration folders live
$SSMSRoot = if (!([string]::IsNullOrEmpty($OSUserName))) {
    Join-Path -Path "C:\Users\$OSUserName\AppData\Local" -ChildPath 'Microsoft\SSMS'
} else {
    Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\SSMS'
}
if (-not (Test-Path -Path $SSMSRoot)) {
    if (!([string]::IsNullOrEmpty($OSUserName))) {
        throw "The SSMS root directory ($SSMSRoot) does not exist. Ensure SSMS is installed and the specified user name ($OSUserName) is correct."
    } else {
        throw "The SSMS root directory ($SSMSRoot) does not exist. Ensure SSMS is installed."
    }
}

# just to make sure SSMS is not running
while (Get-Process "SSMS" -ErrorAction SilentlyContinue | Where-Object { $_.FileVersion -match '^(21|22)\.' }) {
    Write-Host " SSMS process detected. Please close all SSMS instances to proceed." -Fore Yellow
    Read-Host -Prompt "Press Enter after all SSMS 21/22 processes have been closed" 
}

# a little helper function to invoke reg.exe commands
function Invoke-Reg {
    param(
        [Parameter(Mandatory)][string] $Command,
        [Parameter(Mandatory)][string[]] $Arguments
    )
    $argsLine = $Arguments -join ' '
    $proc = Start-Process -FilePath 'reg.exe' -ArgumentList $Command, $argsLine `
        -NoNewWindow -PassThru -Wait -RedirectStandardError err.txt
    if ($proc.ExitCode -ne 0) {
        $err = Get-Content -Path err.txt -Raw
        Remove-Item err.txt -Force
        throw "reg $Command $argsLine failed (Exit $($proc.ExitCode)): $err"
    }
    Remove-Item err.txt -Force
}

# use sdk.txt to figure out which folders belong to GA versions (not previews)
$MatchingDirs = Get-ChildItem -Path $SSMSRoot -Directory |
Where-Object {
    $sdkFile = Join-Path $_.FullName 'sdk.txt'
    Test-Path $sdkFile -PathType Leaf
} |
ForEach-Object {
    $content = (Get-Content -Path (Join-Path $_.FullName 'sdk.txt') -Raw).Trim()
    if ($content -eq 'UsePreviews=False') { $_.FullName }
}
# we can't do anything if we don't have exactly two folders at this point
if ($MatchingDirs.Count -ne 2) {
    throw "Could not locate two matching SSMS version folders. Found $($MatchingDirs.Count)."
}

# figure out which folder is which version (21 vs 22)
$Folder21 = $null
$Folder22 = $null

foreach ($dir in $MatchingDirs) {
    $leaf = Split-Path $dir -Leaf   # e.g. 21.0_56b2001a
    if ($leaf -match '^21\.0_') { 
        $Folder21 = $dir 
        $FName21 = $leaf
    } elseif ($leaf -match '^22\.0_') {
        $Folder22 = $dir 
        $FName22 = $leaf
    }
}

if (-not $Folder21 -or -not $Folder22) {
    throw "Unable to determine which folder is 21 and which is 22. Detected folders:`n$($MatchingDirs -join "`n")"
}

# construct full paths to the privateregistry.bin files
$Hive21 = Join-Path -Path $Folder21 -ChildPath $BinFile
$Hive22 = Join-Path -Path $Folder22 -ChildPath $BinFile
#in case of sapces, quote the paths for reg.exe
$Hive21Safe = '"' + $Hive21 + '"'
$Hive22Safe = '"' + $Hive22 + '"'

# Verify the hive files exist and back them up
foreach ($h in @($Hive21, $Hive22)) {
    if (-not (Test-Path $h -PathType Leaf)) {
        throw "Hive file not found: $h"
    } else {
        Write-Host " Found hive file:`n  $h" -Fore Green
        if ($h -eq $Hive22) {
            Copy-Item -Path $h -Destination "$h.bak" -Force
            Write-Host " Backup created at:`n  $h.bak" -Fore Green
        }
    }
}

# Dude, where's my desktop?
$DesktopPath = [Environment]::GetFolderPath('Desktop')
$RegFilePath = Join-Path -Path $DesktopPath -ChildPath $RegFileName
# have a quoted path for reg.exe
$RegFilePathSafe = '"' + $RegFilePath + '"'
# remove existing .reg file if present
if ( Test-Path $RegFilePath) {
    Remove-Item -Path $RegFilePath -Force
}
try {
    # load hive, export key, unload hive
    Write-Host "`n Loading SSMS 21 hive..."
    Invoke-Reg -Command 'load' -Arguments @($MountPoint, $Hive21Safe) -ErrorAction Stop

    $ExportPath = "$MountPoint\Software\Microsoft\SSMS\$FName21\ConnectionMruList"
    Write-Host " Exporting $RegFileName to Desktop..."
    Invoke-Reg -Command 'export' -Arguments @(
        $ExportPath,
        $RegFilePathSafe,
        '/reg:64'
    ) -ErrorAction Stop

    Write-Host " Unloading SSMS 21 hive..."
    Invoke-Reg -Command 'unload' -Arguments @($MountPoint) -ErrorAction Stop

    # read the exported .rg file, get entries count, and replace the old path with the new one
    $RegContents = Get-Content -Path $RegFilePath -Raw
    $Pattern = '"ConnectionName\d*"\s*='
    $ConnCount = ([regex]::Matches($RegContents, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Count

    Write-Host "`n Found $ConnCount connection entries in the exported $RegFileName file." -Fore Green

    $OldRegPath = "$MountPoint\Software\Microsoft\SSMS\$FName21\ConnectionMruList"
    $NewRegPath = "$MountPoint\Software\Microsoft\SSMS\$FName22\ConnectionMruList"
    # Replace the old path with the new one (Unicode/UTF‑16LE required for .reg files)
    Write-Host " Updating $RegFileName..."
    $RegContents = $RegContents -replace [regex]::Escape($OldRegPath), $NewRegPath
    $RegContents | Set-Content -Path $RegFilePath -Encoding Unicode

    # load SSMS 22 hive and import edited .reg file
    Write-Host "`n Loading SSMS 22 hive..."
    Invoke-Reg -Command 'load' -Arguments @($MountPoint, $Hive22Safe) -ErrorAction Stop

    Write-Host " Importing $RegFileName into SSMS 22 hive..."
    Invoke-Reg -Command 'import' -Arguments @($RegFilePathSafe) -ErrorAction Stop

    # Unload SSMS 22 hive and finish up
    Write-Host " Unloading SSMS 22 hive..."
    Invoke-Reg -Command 'unload' -Arguments @($MountPoint) -ErrorAction Stop
    Remove-Item -Path $RegFilePath -Force

    # that's all, folks
    Write-Host "`n All operations completed successfully." -Fore Green
    Write-Host " Start SSMS 22 and verify that your SSMS 21 connections have been imported."
    Write-Host " If everything looks good, you can delete the backup file."
    Write-Host "`n If you encounter any issues:`n  1. Close SSMS 22 `n  2. Restore SSMS 22's original $BinFile file using the following commands:"
    Write-Host "      Remove-Item -Path `"$Hive22`" -Force" -Fore Yellow
    Write-Host "      Copy-Item -Path `"$Hive22.bak`" ```n      -Destination `"$Hive22`" -Force" -Fore Yellow
    Write-Host "  3. Restart SSMS 22."
} catch {
    Write-Host " Something went wrong: $_" -Fore Red
    # Attempt to unload any loaded hive
    if ( Test-Path $MountPointTest ) {
        Write-Host " Attempting to unload loaded hive..."
        try {
            Invoke-Reg -Command 'unload' -Arguments @($MountPoint) -ErrorAction Stop
            Write-Host " Hive unloaded successfully." -Fore Green
        } catch {
            Write-Host " Failed to unload hive: $_" -Fore Red
            Write-Host " You may need to unload $MountPoint manually using Registry Editor." -Fore Red
        }
    }
} 