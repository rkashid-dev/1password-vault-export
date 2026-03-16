# 1Password CLI — Full Vault Export

A PowerShell script that exports every item from your 1Password vaults to structured, per-category CSV files. Covers all 23 official 1Password item categories and downloads all file attachments and document bodies.

## Features

- Exports all 23 item categories to separate, clearly named CSV files
- Downloads document bodies and embedded file attachments
- Consistent column ordering: identity columns → fixed category columns → dynamic/custom fields
- No data is dropped — duplicate field names get `_raw` or numeric suffixes
- TOTP/OTP fields exported as both raw `otpauth://` URI (importable) and live code
- Archived items optionally included
- Auto re-authentication on session expiry
- Full timestamped export log

## Prerequisites

| Requirement | Notes |
|---|---|
| [1Password desktop app](https://1password.com/downloads/) | Must be installed, running, and **unlocked** |
| [1Password CLI (`op`)](https://developer.1password.com/docs/cli/get-started/) | Must be installed and on `PATH` |
| PowerShell 5.1+ or PowerShell 7+ | Built-in on Windows 10/11 |

## Configuration

Open `Export-1PVault.ps1` and edit the three variables at the top:

```powershell
# Export a single vault by name, or leave "" to export ALL vaults
$VAULT_NAME = ""

# Set to $false to skip archived items
$INCLUDE_ARCHIVE = $true

# Update this to match your Windows username
$OUTPUT_ROOT = "C:\Users\YourName\Desktop"
```

## Usage

```powershell
# Run from PowerShell (may need to adjust execution policy)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Export-1PVault.ps1
```

## Output Structure

```
Desktop\
  1PExport_<timestamp>\
    <VaultName>\
      Logins.csv
      Passwords.csv
      SecureNotes.csv
      ApiCredentials.csv
      CreditCards.csv
      SshKeys.csv
      ... (one file per category found)
    Attachments\
      <VaultName>\
        <ItemTitle>_<ItemID>\
          <filename>
    export_log.txt
```

## Categories Covered

| Category | CSV File |
|---|---|
| Login | `Logins.csv` |
| Password | `Passwords.csv` |
| Secure Note | `SecureNotes.csv` |
| Document | `Documents.csv` |
| API Credential | `ApiCredentials.csv` |
| Amazon Web Services | `AmazonWebServices.csv` |
| Bank Account | `BankAccounts.csv` |
| Credit Card | `CreditCards.csv` |
| Crypto Wallet | `CryptoWallets.csv` |
| Database | `Databases.csv` |
| Driver License | `DriverLicenses.csv` |
| Email Account | `EmailAccounts.csv` |
| Identity | `Identities.csv` |
| Medical Record | `MedicalRecords.csv` |
| Membership | `Memberships.csv` |
| Outdoor License | `OutdoorLicenses.csv` |
| Passport | `Passports.csv` |
| Reward Program | `RewardPrograms.csv` |
| Server | `Servers.csv` |
| Social Security Number | `SocialSecurityNumbers.csv` |
| Software License | `SoftwareLicenses.csv` |
| SSH Key | `SshKeys.csv` |
| Wireless Router | `WirelessRouters.csv` |

Any unrecognised or future categories are written to `Other_<CATEGORY>.csv`.

## CSV Column Order

Every CSV follows the same structure:

1. **Identity columns** — `Vault`, `ItemID`, `Title`, `Category`, `URLs`, `Tags`, `Notes`, `Attachments`, `CreatedAt`, `UpdatedAt`
2. **Fixed columns** — category-specific built-in fields (always present, even if empty)
3. **Dynamic columns** — any custom or extra fields not matched to a fixed slot

## Security Notice

> **This script exports credentials in plaintext.** Treat the output folder with the same care as your 1Password vault. Delete the export files securely once you no longer need them.

## Author

RK | March 2026

## License

MIT
