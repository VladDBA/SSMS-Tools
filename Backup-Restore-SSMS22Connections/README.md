# Backup-Restore-SSMS22Connections.ps1

## Overview

This PowerShell script backs up and restores the saved connections of **SQL Server Management Studio (SSMS) 22** by exporting the `ConnectionMruList` key from the `privateregistry.bin` file to a `.reg` file, or importing that `.reg` file back into it.

---

## What it does

- Dynamically locates the SSMS 22 version folder based on its `sdk.txt` file (in case you had the Preview for SSMS 22 installed at some point).
- **Backup**: loads the SSMS 22 `privateregistry.bin` hive and exports the `ConnectionMruList` key to the `.reg` file specified by `-BackupFile`.
- **Restore**: backs up the SSMS 22 `privateregistry.bin` file, loads the hive, deletes the existing `ConnectionMruList` key, and imports the `.reg` file specified by `-BackupFile`.
- Updates the key path inside the `.reg` file to match the current SSMS 22 version folder, so backups survive SSMS 22 updates that change the folder name.

---

## Requirements

- **Administrator Rights**: The script requires elevated privileges to load/unload the registry hive and access the `privateregistry.bin` file.
- **SSMS 22 Installed**: GA release (the script ignores Preview folders).
- **PowerShell**: PowerShell 5.1 or newer.

---

## Usage

1. **Close SSMS**: Ensure all instances of SSMS 22 are closed before running the script.
2. **Back up your connections**:

   ```powershell
   .\Backup-Restore-SSMS22Connections.ps1 -Backup -BackupFile "D:\Backups\SSMS22_Connections.reg"
   ```

3. **Restore your connections**:

   ```powershell
   .\Backup-Restore-SSMS22Connections.ps1 -Restore -BackupFile "D:\Backups\SSMS22_Connections.reg"
   ```

4. **Optional**: Use the `-OSUserName` parameter to use a different user profile:

   ```powershell
   .\Backup-Restore-SSMS22Connections.ps1 -Backup -BackupFile "D:\Backups\SSMS22_Connections.reg" -OSUserName "OtherUser"
   ```

5. **Follow Prompts**: When restoring, the script will prompt you to confirm before proceeding.
6. **Verify Connections**:
   - Open SSMS 22 and check that your connections are there.

### Parameters

- `-Backup` - exports `ConnectionMruList` to the file specified by `-BackupFile`. An existing file is overwritten.
- `-Restore` - imports the file specified by `-BackupFile`, replacing the existing `ConnectionMruList`.
- `-BackupFile` - full path of the `.reg` file to create or read.
- `-OSUserName` - (optional) OS user profile name used to locate the SSMS configuration folders.

Exactly one of `-Backup` or `-Restore` must be used.

---

## Notes

- **Backup Files**: Before a restore, the script copies `privateregistry.bin` to the same directory as the original, with a timestamped `.bak` extension (e.g. `privateregistry.bin.20261014_101500.bak`).
- **Overwrite**: A restore replaces the existing `ConnectionMruList`. Connections added after the backup was taken are lost.
- **Restore Scope**: The connection details are encrypted with DPAPI, so a backup can't be restored between different OS user profiles and/or different machines.
- **Error Handling**: If an error occurs, the script provides detailed error messages and unloads the registry hive.

---

## Troubleshooting

If you encounter issues after a restore:

1. Close SSMS 22.
2. Restore the original `privateregistry.bin` file for SSMS 22:

   ```powershell
   Remove-Item -Path "path\to\privateregistry.bin" -Force
   Copy-Item -Path "path\to\privateregistry.bin.<timestamp>.bak" -Destination "path\to\privateregistry.bin" -Force
   ```

3. Restart SSMS 22.

If the script fails to unload the hive, unload `HKEY_LOCAL_MACHINE\SSMSStuff` manually using Registry Editor.
