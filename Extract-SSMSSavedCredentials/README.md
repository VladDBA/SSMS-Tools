# Extract-SSMSSavedCredentials.ps1

## Overview

This PowerShell script extracts and decrypts DPAPI-encrypted connection strings and credentials from **SQL Server Management Studio (SSMS)** `privateregistry.bin` files. It retrieves saved connections from SSMS configuration folders, decrypts them, and writes the results to text files for easy access.

---

## What it does

- Locates `privateregistry.bin` files for SSMS 21 and SSMS 22.
- Extracts and decrypts the `ConnectionMruList` registry keys.
- Handles DPAPI-protected connection strings.
- Outputs decrypted connection strings and credentials to text files.
- Provides an option to retain intermediate `.reg` files for debugging.

---

## Requirements

- **Administrator Rights**: The script requires elevated privileges to access and modify the registry.
- **PowerShell 7.0 or newer**: I couldn't get PS 5.1 to work with the `[System.Security.Cryptography.ProtectedData]` class.
- **SSMS Installed**: Works with SSMS 21 and SSMS 22, both Preview and GA releases.

---

## Usage

1. **Close SSMS**: Ensure all instances of SSMS are closed before running the script.
2. **Run the Script**:

   ```powershell
   .\Extract-SSMSSavedCredentials.ps1
   ```

3. **Optional**: Use the `-KeepRegFiles` parameter to retain intermediate `.reg` files:

   ```powershell
   .\Extract-SSMSSavedCredentials.ps1 -KeepRegFiles
   ```

4. **View Results**:
   - `DecryptedConnectionStrings.txt`
   - `DecryptedCredentials.txt`

---

## Notes

- **Backup**: The script does not modify the original `privateregistry.bin` files.
- **Decryption Scope**: Only works for the current user profile and machine due to DPAPI encryption.
- **Output Files**:
  - `DecryptedConnectionStrings.txt` - contains decrypted raw connection strings.
  - `DecryptedCredentials.txt` - contains only data sources with matching user IDs and passwords extracted from the raw connection strings.

---

## Troubleshooting

If you encounter issues:

1. Ensure you are running the script as an administrator.
2. Verify that SSMS 21 or SSMS 22 is installed on your machine.
3. Check for errors in the console output for more details.