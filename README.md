# Disable New Outlook Toggle via PowerShell

## Overview
This PowerShell script disables the "New Outlook" toggle for all eligible user profiles on a Windows machine. It modifies the Windows registry settings for Microsoft Outlook to enforce this change.

## How It Works
1. Retrieves all user SIDs from the `HKEY_USERS` registry.
2. Skips system and service accounts.
3. Checks if the Outlook registry path exists for each user.
4. Sets the registry values to:
   - `UseNewOutlook` = `0` in `Preferences`
   - `HideNewOutlookToggle` = `1` in `Options\General`
5. Displays a summary of the changes made.

## Registry Paths Affected
- `HKEY_USERS\<User SID>\Software\Microsoft\Office\16.0\Outlook\Preferences`
  - `UseNewOutlook` (DWORD) = `0`
- `HKEY_USERS\<User SID>\Software\Microsoft\Office\16.0\Outlook\Options\General`
  - `HideNewOutlookToggle` (DWORD) = `1`

## Usage
Run the script as an administrator:

```powershell
powershell -ExecutionPolicy Bypass -File DisableNewOutlook.ps1
```

## Requirements
- Administrator privileges.
- Windows machine with Microsoft Office 2016 (16.0) or later installed.
- PowerShell 5.1 or later.

## Notes
- If the registry paths do not exist, the script will notify you but will not create them.
- Only user profiles with Outlook registry settings present will be modified.

## License
This script is provided as-is without any warranty. Use at your own risk.

## Author
DaanSelen

