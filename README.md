# SSMS-Tools

Small collection of PowerShell utilities and config files for SQL Server Management Studio 21 and 22.

## Overview

This repository provides a small set of Windows PowerShell utilities SQL Server Management Studio (SSMS) 21 and 22, and a sample SSMS 22 configuration JSON.

All scripts operate on local SSMS configuration data, require administrator privileges to load/unload registry hives, and are intended to be run on Windows.

## Scripts

- [Import-SSMS21ConnectionsToSSMS22.ps1](/Import-SSMS21ConnectionsToSSMS22/)  
  Imports saved SSMS 21 connection entries into SSMS 22 by copying the `ConnectionMruList` registry data from one `privateregistry.bin` hive to another.
  More info: [Import saved connections from SSMS 21 to SSMS 22](https://vladdba.com/2025/11/17/import-saved-connections-from-ssms-21-to-ssms-22/)

- [Extract-SSMSSavedCredentials.ps1](/Extract-SSMSSavedCredentials/)
  Locates `privateregistry.bin` files, exports `ConnectionMruList` registries, decrypts DPAPI-protected blobs, and writes human‑readable output files:
  - `DecryptedConnectionStrings.txt`
  - `DecryptedCredentials.txt`  
  Requires PowerShell 7.x (DPAPI) and admin rights.
  More info: [PowerShell script to extract SSMS 21 and 22 saved connection data](https://vladdba.com/2025/11/22/powershell-extract-ssms-21-22-saved-connection-information/)

### Usage

Run from an elevated PowerShell session. Examples:

````powershell
# migrate connections from SSMS21 to SSMS22
.\Import-SSMS21ConnectionsToSSMS22.ps1

# extract and decrypt saved credentials (keep .reg files for inspection)
.\Extract-SSMSSavedCredentials.ps1 -KeepRegFiles
````

## Files

- [My-SSMS22-Config.JSON](/My-SSMS22-Config.JSON)
  My current SQL Server Mangement Studio 22 configuration JSON as described in [this blog post](https://vladdba.com/2025/11/16/my-sql-server-management-studio-22-configuration/)