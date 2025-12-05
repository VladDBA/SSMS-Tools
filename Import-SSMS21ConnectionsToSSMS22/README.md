# Import-SSMS21ConnectionsToSSMS22.ps1

## Overview

This PowerShell script imports saved connections from **SQL Server Management Studio (SSMS) 21** to **SSMS 22** by transferring the `ConnectionMruList` key from the `privateregistry.bin` file of SSMS 21 to SSMS 22. Thus helping you not waste time manually re-adding all your SSMS 21 connections to SSMS 22.

---

## What it does

- Dynamically locates the SSMS version folders based on `sdk.txt` files (in case you had the Preview for SSMS 21 and/or 22 installed at some point).
- Backs up the SSMS 22 `privateregistry.bin` file before making any changes.
- Exports the `ConnectionMruList` key from SSMS 21's `privateregistry.bin` files.
- Updates the exported registry key to match SSMS 22 and imports it into SSMS 22.

---

## Requirements

- **Administrator Rights**: The script requires elevated privileges to modify the registry and access the `privateregistry.bin` files.
- **SSMS 21 and SSMS 22 Installed**: Both versions must be installed on the same machine.
- **PowerShell**: Built and tested for PowerShell 5.1 and newer.

---

## Usage

1. **Close SSMS**: Ensure both SSMS 21 and SSMS 22 are closed before running the script.
2. **Run the Script**:

   ```powershell
   .\Import-SSMS21ConnectionsToSSMS22.ps1
   ```

3. **Optional**: Use the `-OSUserName` parameter to use a different user profile for the import

   ```powershell
   .\Import-SSMS21ConnectionsToSSMS22.ps1 -OSUserName "OtherUser"
   ```

4. **Follow Prompts**: The script will prompt you to confirm before proceeding.
5. **Verify Connections**:
   - Open SSMS 22 and check if your SSMS 21 connections have been imported successfully.

---

## Notes

- **Backup Files**: The script creates backup copies of the `privateregistry.bin` file before making changes. This is stored in the same directory as the original SSMS 22 file with a `.bak` extension.
- **Error Handling**: If an error occurs, the script provides detailed error messages and you can use the previously generated backup to undo any changes.

---

## Troubleshooting

If you encounter issues:

1. Close SSMS 22.
2. Restore the original `privateregistry.bin` file for SSMS 22:

   ```powershell
   Remove-Item -Path "path\to\privateregistry.bin" -Force
   Copy-Item -Path "path\to\privateregistry.bin.bak" -Destination "path\to\privateregistry.bin" -Force
   ```

3. Restart SSMS 22.

---

## More info

For more details, visit the [Import Saved Connections from SSMS 21 to SSMS 22](https://vladdba.com/2025/11/17/import-saved-connections-from-ssms-21-to-ssms-22/) blog post.
