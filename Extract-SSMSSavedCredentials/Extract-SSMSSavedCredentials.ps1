#Requires -RunAsAdministrator
#Requires -Version 7.0

<#

.SYNOPSIS
    Extracts and decrypts DPAPI-encrypted connection strings and credentials from SSMS privateregistry.bin files.

.DESCRIPTION
    This script locates privateregistry.bin files from SSMS configuration folders, extracts the ConnectionMruList registry keys, 
    decrypts the DPAPI-protected connection strings, and writes the results to text files.

    The script uses the current user's DPAPI key to decrypt the connection strings and credentials.
    data within the decrypted strings.

    Will not work with privateregistry.bin files from other machines or user profiles, as DPAPI encryption is tied to the user context.

.PARAMETER KeepRegFiles
    A switch parameter that, when specified, prevents the script from deleting the intermediate .reg files created during the process.
    By default, each .reg file is deleted after processing.

.EXAMPLE
    .\Extract-SSMSSavedCredentials.ps1
    Extracts and decrypts connection strings and credentials from all SSMS privateregistry.bin files found in the default SSMS 
    configuration folders. The results are saved to `DecryptedConnectionStrings.txt` and `DecryptedCredentials.txt` in the script's directory.

.EXAMPLE
    .\Extract-SSMSSavedCredentials.ps1 -KeepRegFiles
    Performs the same operation as the previous example, but retains the intermediate .reg files for further inspection.

.NOTES
    - Requires admin rights.
    - Works only for SSMS 21 and SSMS 22.
    - Works only for the current user profile, as DPAPI encryption is user-specific.
    - Requires PowerSell 7.0 or higher.

  Author: Vlad Drumea (VladDBA)
  Website: https://vladdba.com
  Date: November 18, 2025

.INPUTS
    None.

.OUTPUTS
    The script writes the decrypted connection strings and credentials to text files in the script's directory.
    - DecryptedConnectionStrings.txt: Contains the decrypted connection strings in the format:
      Connection Name: Connection String
    - DecryptedCredentials.txt: Contains the decrypted credentials - data source, user ID, and password.

.LINK
    For more information, visit: https://vladdba.com/2025/11/22/powershell-extract-ssms-21-22-saved-connection-information/
#>


param (
    [Parameter(Mandatory = $false)]
    [switch]$KeepRegFiles
)
# helper function to turn a hex string into a byte array
function Convert-HexToByteArray {
    param([string]$Hex)

    $clean = $Hex -replace '\s', '' # strip any whitespace
    if ($clean.Length % 2 -ne 0) {
        throw "Hex string length must be even. Got $($clean.Length) characters."
    }

    $bytes = New-Object byte[] ($clean.Length / 2)
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        $bytes[$i] = [Convert]::ToByte($clean.Substring($i * 2, 2), 16)
    }
    return $bytes
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
        exit
    }
    Remove-Item err.txt -Force
}

## internal variables
$BinFile = 'privateregistry.bin'
$MountPoint = 'HKEY_LOCAL_MACHINE\SSMSStuff'
$MountPointTest = 'HKLM:\SSMSStuff'
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$RawConnStringPath = Join-Path -Path $ScriptPath -ChildPath 'DecryptedConnectionStrings.txt'
$CredentialsPath = Join-Path -Path $ScriptPath -ChildPath 'DecryptedCredentials.txt'

$ConnStrings = @()
$Credentials = @()

$TotalConnStrings = 0
$TotalCredentials = 0

# where SSMS 21 and 22 related configuration folders live
$SSMSRoot = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Microsoft\SSMS'

## checks and confirmations

Write-Host " Make sure all instances of SSMS are closed before proceeding" -Fore Yellow
$Confirmation = Read-Host -Prompt "Do you want to continue? (Y/N)"
if ($Confirmation -ne "Y") {
    Write-Host " Operation cancelled by user."
    exit
}

if (-not (Test-Path -Path $SSMSRoot)) {
    throw "The SSMS root directory ($SSMSRoot) does not exist. Is SSMS even installed?"
}

# just to make sure SSMS is not running
while (Get-Process "SSMS" -ErrorAction SilentlyContinue | Where-Object { $_.FileVersion -match '^(21|22)\.' }) {
    Write-Host " SSMS process detected. Please close all SSMS instances to proceed." -Fore Yellow
    Read-Host -Prompt "Press Enter after all SSMS 21/22 processes have been closed" 
}

## figure out if there's anything to work with
$MatchingDirs = Get-ChildItem -Path $SSMSRoot -Directory |
Where-Object {
    Test-Path -Path (Join-Path $_.FullName $BinFile) -PathType Leaf
} |
Select-Object -ExpandProperty FullName

# did we find any matching dirs?
if ($MatchingDirs.Count -eq 0) {
    throw "No SSMS configuration folders with privateregistry.bin files were found."
}

# how many matching dirs did we find?
$DirCount = $MatchingDirs.Count
if ($DirCount -eq 0) {
    throw "Could not locate any SSMS configuration folders with privateregistry.bin files."
    exit
}
Write-Host "`n Found $DirCount SSMS configuration folders with privateregistry.bin files." -Fore Green    

# process each matching hive
foreach ($Folder in $MatchingDirs) {
    $FName = Split-Path $Folder -Leaf   # will use this as an identifier
    $Hive = Join-Path -Path $Folder -ChildPath $BinFile
    # have a quoted path for reg.exe
    $HiveSafe = '"' + $Hive + '"'
    # make sure there's no mounted hive from a previous run
    if (Test-Path -Path $MountPointTest) {
        Write-Host " Unloading previously mounted hive at $MountPoint..." -Fore Yellow
        Invoke-Reg -Command 'unload' -Arguments @($MountPoint)
    }
    # Mount the hive
    Write-Host "`n Mounting hive from $FName..." -Fore Magenta
    Invoke-Reg -Command 'load' -Arguments @($MountPoint, $HiveSafe)
    
    $RegFileName = "${FName}_ConnectionMruList.reg"
    $RegFilePath = Join-Path -Path $ScriptPath -ChildPath $RegFileName
    # have a quoted path for reg.exe
    $RegFilePathSafe = '"' + $RegFilePath + '"'
    $ExportPath = "$MountPoint\Software\Microsoft\SSMS\$FName\ConnectionMruList"
    
    # make sure there's no leftover .reg file from a previous run
    if (Test-Path -Path $RegFilePath) {
        Remove-Item -Path $RegFilePath -Force
    }

    # Export the relevant key to a .reg file
    Write-Host " Exporting $RegFileName..."
    try {
        Invoke-Reg -Command 'export' -Arguments @(
            $ExportPath,
            $RegFilePathSafe,
            '/reg:64'
        ) -ErrorAction Stop
    } catch {
        Write-Warning "Failed to export $RegFileName" -Fore Red
        # Unload the hive before continuing
        Write-Host " Unloading hive and proceeding..." 
        Invoke-Reg -Command 'unload' -Arguments @($MountPoint)
        continue
    }
    
    # Unload the hive
    Write-Host " Unloading hive..."
    Invoke-Reg -Command 'unload' -Arguments @($MountPoint)
    $ConnStringCounter = 0
    $CredentialsCounter = 0
    $ConnStrings += "`n### Decrypted connections from $FName ###"
    $Credentials += "`n### Decrypted credentials from $FName ###"
    # Load the .reg file
    $Connections = Select-String $RegFilePath -Pattern '"Connection\d' -Raw
    
    if ( $Connections.Count -eq 0) {
        # no need to continue if there's nothing to process
        Write-Host " No saved connections found in $FName - skipping decryption." -Fore Yellow
        # clean up .reg file unless told otherwise
        if ((-not $KeepRegFiles) -and (Test-Path -Path $RegFilePath)) {
            Remove-Item -Path $RegFilePath -Force
        }
        continue
    } else {
        Write-Host " Found $($Connections.Count) saved connections in $FName." -Fore Green
    }
    foreach ($Connection in $Connections) {
        $ConnName = $Connection -replace "`"=.*", "" -replace '"', ''
        $HexBlob = $Connection -replace '"Connection\d+"="', '' -replace '"', ''
        if ([string]::IsNullOrWhiteSpace($HexBlob)) {
            Write-Warning " $ConnName missing a blob - skipping."
            continue
        }

        try {
            # DPAPI decryption
            $bytes = Convert-HexToByteArray -Hex $hexBlob
            $plain = [System.Security.Cryptography.ProtectedData]::Unprotect(
                $bytes,
                $null,
                [System.Security.Cryptography.DataProtectionScope]::CurrentUser)

            $clearText = [Text.Encoding]::Unicode.GetString($plain)

            # find stored credentials and format them nicely
            if ($ClearText -like "*;Persist Security Info=True*") {
                $CredentialsCounter++
                $DataSource = $ClearText -replace ";Persist Security Info=True*.+", ""
                $DataSource = $DataSource -replace ".*Data Source=", ""
                $UserID = $ClearText -replace ".*User ID=", "User ID=" -replace ";Password*.+", ""
                $Pass = $ClearText -replace ".*;Password=", "Password=" -replace ";Pooling=.*", ""
                $Credentials += "## $DataSource `n     $UserID `n     $Pass"
            }

            $ConnStrings += "$clearText"
            $ConnStringCounter++
            
        } catch {
            $msg = $_.Exception.Message
            Write-Warning "Failed to decrypt ${ConnName}: $msg"
            $ConnStrings += "${ConnName}: *** DECRYPTION FAILED ***"
        }
    }

    Write-Host " Decrypted connection strings:" $ConnStringCounter -Fore Green
    Write-Host " Decrypted credentials:" $CredentialsCounter -Fore Green
    $TotalConnStrings += $ConnStringCounter
    $TotalCredentials += $CredentialsCounter

    # clean up .reg file unless told otherwise
    if ((-not $KeepRegFiles) -and (Test-Path -Path $RegFilePath)) {
        Remove-Item -Path $RegFilePath -Force
    }
}
# write output files
$ConnStrings | Set-Content -Path $RawConnStringPath -Encoding Unicode
$Credentials | Set-Content -Path $CredentialsPath -Encoding Unicode
Write-Host "`n Decryption complete." -Fore Green
Write-Host " Total extracted connection strings: $TotalConnStrings" -Fore Green
Write-Host " Total extracted credentials: $TotalCredentials" -Fore Green