# ==============================================================================
# 1Password CLI — Full Vault Export Script
# Version: 3.0  |  Author: RK  |  March 2026
#
# Exports every vault item to per-category CSV files.
# Covers ALL 23 official 1Password item categories.
# Downloads all file attachments and DOCUMENT bodies.
# Field IDs sourced directly from: op item template get <category>
#
# CATEGORIES COVERED:
#   LOGIN                 PASSWORD              SECURE_NOTE
#   DOCUMENT              API_CREDENTIAL        AMAZON_WEB_SERVICES
#   BANK_ACCOUNT          CREDIT_CARD           CRYPTO_WALLET
#   DATABASE              DRIVER_LICENSE        EMAIL_ACCOUNT
#   IDENTITY              MEDICAL_RECORD        MEMBERSHIP
#   OUTDOOR_LICENSE       PASSPORT              REWARD_PROGRAM
#   SERVER                SOCIAL_SECURITY_NUMBER SOFTWARE_LICENSE
#   SSH_KEY               WIRELESS_ROUTER
#
# OUTPUT STRUCTURE:
#   Desktop\1PExport_<timestamp>\
#     <VaultName>\
#       Logins.csv
#       Passwords.csv
#       SecureNotes.csv
#       Documents.csv
#       ApiCredentials.csv
#       AmazonWebServices.csv
#       BankAccounts.csv
#       CreditCards.csv
#       CryptoWallets.csv
#       Databases.csv
#       DriverLicenses.csv
#       EmailAccounts.csv
#       Identities.csv
#       MedicalRecords.csv
#       Memberships.csv
#       OutdoorLicenses.csv
#       Passports.csv
#       RewardPrograms.csv
#       Servers.csv
#       SocialSecurityNumbers.csv
#       SoftwareLicenses.csv
#       SshKeys.csv
#       WirelessRouters.csv
#       Other_<CATKEY>.csv        <- any unknown / future category
#     Attachments\
#       <VaultName>\
#         <SafeTitle>_<ItemID>\
#           <filename>
#     export_log.txt
#
# CSV COLUMN ORDER (every file):
#   [Identity cols] Vault | ItemID | Title | Category | URLs | Tags |
#                   Notes | Attachments | CreatedAt | UpdatedAt
#   [Fixed cols]    Category-specific built-in fields (always present)
#   [Dynamic cols]  Any custom / extra fields not matched to a fixed slot
#
# PREREQUISITES:
#   - 1Password desktop app installed, running, and UNLOCKED
#   - op CLI installed and on PATH
#   - PowerShell 5.1+ or PowerShell 7+
# ==============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ==============================================================================
# CONFIGURATION  —  edit these before running
# ==============================================================================

# Set to a vault name to export only that vault, e.g. "Infrastructure"
# Leave as "" to export ALL vaults the account has access to
$VAULT_NAME = ""

# Include items that have been archived?
$INCLUDE_ARCHIVE = $true

# Output root — a timestamped subfolder is created here automatically
# Update the username to match the machine you are running this on
$OUTPUT_ROOT = "C:\Users\Carl.Engineer\Desktop"

# ==============================================================================
# DERIVED PATHS  —  do not edit
# ==============================================================================

$TIMESTAMP       = Get-Date -Format "yyyyMMdd_HHmmss"
$OUTPUT_DIR      = "$OUTPUT_ROOT\1PExport_$TIMESTAMP"
$ATTACHMENTS_DIR = "$OUTPUT_DIR\Attachments"
$LOG_PATH        = "$OUTPUT_DIR\export_log.txt"

# ==============================================================================
# CATEGORY -> CSV FILENAME MAP
# ==============================================================================

$script:CAT_FILE = @{
    "LOGIN"                  = "Logins"
    "PASSWORD"               = "Passwords"
    "SECURE_NOTE"            = "SecureNotes"
    "DOCUMENT"               = "Documents"
    "API_CREDENTIAL"         = "ApiCredentials"
    "AMAZON_WEB_SERVICES"    = "AmazonWebServices"
    "BANK_ACCOUNT"           = "BankAccounts"
    "CREDIT_CARD"            = "CreditCards"
    "CRYPTO_WALLET"          = "CryptoWallets"
    "DATABASE"               = "Databases"
    "DRIVER_LICENSE"         = "DriverLicenses"
    "EMAIL_ACCOUNT"          = "EmailAccounts"
    "IDENTITY"               = "Identities"
    "MEDICAL_RECORD"         = "MedicalRecords"
    "MEMBERSHIP"             = "Memberships"
    "OUTDOOR_LICENSE"        = "OutdoorLicenses"
    "PASSPORT"               = "Passports"
    "REWARD_PROGRAM"         = "RewardPrograms"
    "SERVER"                 = "Servers"
    "SOCIAL_SECURITY_NUMBER" = "SocialSecurityNumbers"
    "SOFTWARE_LICENSE"       = "SoftwareLicenses"
    "SSH_KEY"                = "SshKeys"
    "WIRELESS_ROUTER"        = "WirelessRouters"
}

# ==============================================================================
# FIXED COLUMN DEFINITIONS PER CATEGORY
# These are the named columns that always appear in each category CSV,
# even when empty. Field IDs sourced from: op item template get <category>
# Column order here = column order in the CSV (after identity cols).
# ==============================================================================

$script:CAT_COLS = @{

    # --------------------------------------------------------------------------
    # LOGIN  |  op item template get Login
    # Fixed IDs: username, password, notesPlain
    # OTP:       type=OTP  (no stable built-in ID — matched by type)
    # --------------------------------------------------------------------------
    "LOGIN" = @(
        "Username",         # id: username
        "Password",         # id: password
        "TOTP_URI",         # OTP field .value  = raw otpauth:// URI (importable)
        "OneTimePassword"   # OTP field .totp   = live code at export time
    )

    # --------------------------------------------------------------------------
    # PASSWORD  |  op item template get Password
    # Fixed IDs: password, notesPlain
    # --------------------------------------------------------------------------
    "PASSWORD" = @(
        "Password"          # id: password
    )

    # --------------------------------------------------------------------------
    # SECURE_NOTE  |  op item template get "Secure Note"
    # Fixed IDs: notesPlain only (captured in identity cols as Notes)
    # --------------------------------------------------------------------------
    "SECURE_NOTE" = @()

    # --------------------------------------------------------------------------
    # DOCUMENT  |  op item template get Document
    # File body: op document get <itemId>  (not in fields array)
    # --------------------------------------------------------------------------
    "DOCUMENT" = @(
        "FileName",         # label: filename
        "FileSize"          # label: file size
    )

    # --------------------------------------------------------------------------
    # API_CREDENTIAL  |  op item template get "API Credential"
    # Fixed IDs: username, credential, notesPlain
    # --------------------------------------------------------------------------
    "API_CREDENTIAL" = @(
        "Username",         # id: username
        "Credential",       # id: credential
        "APIType",          # label: type
        "Hostname",         # label: hostname
        "ValidFrom",        # label: valid from
        "Expires",          # label: expires
        "Filename"          # label: filename
    )

    # --------------------------------------------------------------------------
    # AMAZON_WEB_SERVICES  |  op item template get "Amazon Web Services"
    # Fixed IDs: access_key_id, secret_access_key
    # --------------------------------------------------------------------------
    "AMAZON_WEB_SERVICES" = @(
        "AccessKeyID",      # id: access_key_id
        "SecretAccessKey",  # id: secret_access_key
        "DefaultRegion",    # label: default region
        "ConsoleURL",       # label: console dashboard url
        "MFASerial",        # label: mfa serial
        "AccountID"         # label: account id
    )

    # --------------------------------------------------------------------------
    # BANK_ACCOUNT  |  op item template get "Bank Account"
    # Fixed IDs: bankName, owner, accountType, routingNo,
    #            accountNo, swift, iban, telephonePin
    # --------------------------------------------------------------------------
    "BANK_ACCOUNT" = @(
        "BankName",         # id: bankName
        "OwnerName",        # id: owner
        "AccountType",      # id: accountType
        "RoutingNumber",    # id: routingNo
        "AccountNumber",    # id: accountNo
        "SWIFT",            # id: swift
        "IBAN",             # id: iban
        "PIN",              # id: telephonePin
        "PhoneLocal",       # label: phone (local)
        "PhoneTollFree",    # label: phone (toll free)
        "PhoneIntl",        # label: phone (intl)
        "BranchAddress",    # label: branch address
        "BranchPhone"       # label: branch phone
    )

    # --------------------------------------------------------------------------
    # CREDIT_CARD  |  op item template get "Credit Card"
    # Fixed IDs: cardholder, type, ccnum, cvv, expiry, validFrom, pin
    # --------------------------------------------------------------------------
    "CREDIT_CARD" = @(
        "CardholderName",       # id: cardholder
        "CardType",             # id: type
        "CardNumber",           # id: ccnum
        "VerificationNumber",   # id: cvv
        "ExpiryDate",           # id: expiry      (MONTH_YEAR format: YYYYMM)
        "ValidFrom",            # id: validFrom   (MONTH_YEAR format)
        "CreditLimit",          # label: credit limit
        "CashLimit",            # label: cash withdrawal limit
        "InterestRate",         # label: interest rate
        "IssueNumber",          # label: issue number
        "PIN",                  # id: pin
        "IssuingBank",          # label: issuing bank
        "PhoneLocal",           # label: phone (local)
        "PhoneTollFree",        # label: phone (toll free)
        "PhoneIntl",            # label: phone (intl)
        "Website"               # label: website
    )

    # --------------------------------------------------------------------------
    # CRYPTO_WALLET  |  op item template get "Crypto Wallet"
    # Fixed IDs: password, username
    # --------------------------------------------------------------------------
    "CRYPTO_WALLET" = @(
        "RecoveryPhrase",   # label: recovery phrase  (CONCEALED)
        "Password",         # id: password
        "WalletAddress",    # label: wallet address
        "Username"          # id: username
    )

    # --------------------------------------------------------------------------
    # DATABASE  |  op item template get Database
    # Fixed IDs: database_type, hostname, port, database,
    #            username, password, sid, alias, options
    # --------------------------------------------------------------------------
    "DATABASE" = @(
        "DBType",           # id: database_type
        "Hostname",         # id: hostname
        "Port",             # id: port
        "Database",         # id: database
        "Username",         # id: username
        "Password",         # id: password
        "SID",              # id: sid
        "Alias",            # id: alias
        "ConnectionOptions" # id: options
    )

    # --------------------------------------------------------------------------
    # DRIVER_LICENSE  |  op item template get "Driver License"
    # Fixed IDs: fullname, sex, number, birthdate,
    #            expiry_date, class, conditions, state, country, address
    # --------------------------------------------------------------------------
    "DRIVER_LICENSE" = @(
        "FullName",         # id: fullname
        "Sex",              # id: sex
        "LicenseNumber",    # id: number
        "DateOfBirth",      # id: birthdate
        "ExpiryDate",       # id: expiry_date
        "LicenseClass",     # id: class
        "Conditions",       # id: conditions
        "State",            # id: state
        "Country",          # id: country
        "Address"           # id: address
    )

    # --------------------------------------------------------------------------
    # EMAIL_ACCOUNT  |  op item template get "Email Account"
    # Fixed IDs: username, password
    # All other fields matched by label (section-based in the template)
    # --------------------------------------------------------------------------
    "EMAIL_ACCOUNT" = @(
        "Username",         # id: username
        "Password",         # id: password
        "MailFrom",         # label: email
        "Provider",         # label: provider
        "ProviderWebsite",  # label: provider's website
        "AuthMethod",       # label: auth method
        "Security",         # label: security
        "IMAPServer",       # label: imap server        (IMAP section)
        "IMAPPort",         # label: imap port          (IMAP section)
        "SMTPServer",       # label: smtp server        (SMTP section)
        "SMTPPort",         # label: smtp port          (SMTP section)
        "POPServer",        # label: pop server         (POP3 section)
        "POPPort"           # label: pop port           (POP3 section)
    )

    # --------------------------------------------------------------------------
    # IDENTITY  |  op item template get Identity
    # Fixed IDs: firstname, initial, lastname, sex, birthdate,
    #            occupation, company, department, jobtitle, address
    # All phone/email/internet fields matched by label (section-based)
    # --------------------------------------------------------------------------
    "IDENTITY" = @(
        # Personal section
        "FirstName",        # id: firstname
        "Initial",          # id: initial
        "LastName",         # id: lastname
        "Sex",              # id: sex
        "BirthDate",        # id: birthdate
        "Occupation",       # id: occupation
        # Work section
        "Company",          # id: company
        "Department",       # id: department
        "JobTitle",         # id: jobtitle
        # Address section
        "Street",           # id: address
        "City",             # label: city
        "State",            # label: state
        "ZipCode",          # label: zip / postal code
        "Country",          # label: country
        # Phone section
        "DefaultPhone",     # label: default phone
        "HomePhone",        # label: home
        "CellPhone",        # label: cell
        "BusinessPhone",    # label: business
        # Internet section
        "DefaultEmail",     # label: default email
        "HomeEmail",        # label: home email   (distinct from Phone "home")
        "BusinessEmail",    # label: business email
        "Website",          # label: website
        "Username",         # label: username
        "Reminder"          # label: reminder
    )

    # --------------------------------------------------------------------------
    # MEDICAL_RECORD  |  op item template get "Medical Record"
    # No stable built-in IDs in template — all matched by label
    # --------------------------------------------------------------------------
    "MEDICAL_RECORD" = @(
        "Date",                     # label: date
        "Location",                 # label: location
        "HealthcareProfessional",   # label: healthcare professional
        "PatientName",              # label: patient
        "ReasonForVisit",           # label: reason for visit
        "Medication",               # label: medication       (Medication section)
        "Dosage"                    # label: dosage           (Medication section)
    )

    # --------------------------------------------------------------------------
    # MEMBERSHIP  |  op item template get Membership
    # Fixed IDs: org_name, member_name, membership_no,
    #            member_since, expiry_date, website, phone, pin
    # --------------------------------------------------------------------------
    "MEMBERSHIP" = @(
        "GroupOrOrganization",  # id: org_name
        "MemberName",           # id: member_name
        "MemberID",             # id: membership_no
        "MemberSince",          # id: member_since
        "ExpiryDate",           # id: expiry_date
        "Website",              # id: website
        "Phone",                # id: phone
        "PIN"                   # id: pin
    )

    # --------------------------------------------------------------------------
    # OUTDOOR_LICENSE  |  op item template get "Outdoor License"
    # Fixed IDs: name, valid_from, expires,
    #            approved_wildlife, max_catch, state, country
    # --------------------------------------------------------------------------
    "OUTDOOR_LICENSE" = @(
        "FullName",             # id: name
        "ValidFrom",            # id: valid_from
        "ExpiryDate",           # id: expires
        "ApprovedWildlife",     # id: approved_wildlife
        "MaximumQuota",         # id: max_catch
        "State",                # id: state
        "Country"               # id: country
    )

    # --------------------------------------------------------------------------
    # PASSPORT  |  op item template get Passport
    # Fixed IDs: issuing_country, fullname, sex, nationality,
    #            issue_date, expiry_date, number, birthplace, birthdate
    # --------------------------------------------------------------------------
    "PASSPORT" = @(
        "IssuingCountry",   # id: issuing_country
        "FullName",         # id: fullname
        "Sex",              # id: sex
        "Nationality",      # id: nationality
        "IssueDate",        # id: issue_date
        "ExpiryDate",       # id: expiry_date
        "PassportNumber",   # id: number
        "PlaceOfBirth",     # id: birthplace
        "DateOfBirth"       # id: birthdate
    )

    # --------------------------------------------------------------------------
    # REWARD_PROGRAM  |  op item template get "Reward Program"
    # Fixed IDs: company_name, member_name, membership_no, pin, website
    # --------------------------------------------------------------------------
    "REWARD_PROGRAM" = @(
        "CompanyName",      # id: company_name
        "MemberName",       # id: member_name
        "MemberID",         # id: membership_no
        "PIN",              # id: pin
        "MemberSince",      # label: member since
        "ExpiryDate",       # label: expiry date
        "Website",          # id: website
        "PhoneLocal",       # label: phone (local)
        "PhoneTollFree",    # label: phone (toll free)
        "PhoneIntl"         # label: phone (intl)
    )

    # --------------------------------------------------------------------------
    # SERVER  |  op item template get Server
    # Fixed IDs: username, password
    # All other fields matched by label (section-based)
    # --------------------------------------------------------------------------
    "SERVER" = @(
        "URL",              # label: url
        "Username",         # id: username
        "Password",         # id: password
        "HostingProvider",  # label: hosting provider   (Hosting Provider section)
        "AdminConsole",     # label: admin console url  (Admin Console section)
        "AdminUser",        # label: admin console username
        "AdminPass",        # label: admin console password
        "SupportURL",       # label: support url        (Support section)
        "SupportPhone"      # label: support phone
    )

    # --------------------------------------------------------------------------
    # SOCIAL_SECURITY_NUMBER  |  op item template get "Social Security Number"
    # Fixed IDs: name, number
    # --------------------------------------------------------------------------
    "SOCIAL_SECURITY_NUMBER" = @(
        "Name",             # id: name
        "Number"            # id: number
    )

    # --------------------------------------------------------------------------
    # SOFTWARE_LICENSE  |  op item template get "Software License"
    # Fixed IDs: license_key
    # All others matched by label
    # --------------------------------------------------------------------------
    "SOFTWARE_LICENSE" = @(
        "LicenseKey",       # id: license_key
        "Version",          # label: version
        "LicensedTo",       # label: licensed to
        "RegisteredEmail",  # label: registered email
        "Company",          # label: company
        "Publisher",        # label: publisher          (Publisher section)
        "Website",          # label: website
        "OrderNumber",      # label: order number       (Order section)
        "OrderDate",        # label: order date
        "OrderTotal",       # label: retail price
        "SupportEmail"      # label: support email      (Customer section)
    )

    # --------------------------------------------------------------------------
    # SSH_KEY  |  op item template get "SSH Key"
    # Fixed IDs: public_key, private_key, fingerprint
    # --------------------------------------------------------------------------
    "SSH_KEY" = @(
        "PublicKey",        # id: public_key
        "PrivateKey",       # id: private_key
        "Fingerprint",      # id: fingerprint
        "KeyType",          # label: key type
        "Passphrase"        # label: passphrase
    )

    # --------------------------------------------------------------------------
    # WIRELESS_ROUTER  |  op item template get "Wireless Router"
    # All fields matched by label
    # --------------------------------------------------------------------------
    "WIRELESS_ROUTER" = @(
        "BaseStationName",      # label: base station name
        "BaseStationPassword",  # label: base station password
        "AirportID",            # label: airport id
        "NetworkName",          # label: network name       (Wireless Network section)
        "WirelessPassword",     # label: wireless password
        "WirelessSecurity",     # label: wireless security
        "RouterIP",             # label: server / ip address (Router section)
        "AttachmentPassword"    # label: attached storage password
    )
}

# Columns that appear FIRST in every CSV regardless of category
$script:IDENTITY_COLS = @(
    "Vault", "ItemID", "Title", "Category",
    "URLs", "Tags", "Notes", "Attachments",
    "CreatedAt", "UpdatedAt"
)

# ==============================================================================
# WRITE-LOG
# ==============================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$ts] [$Level] $Message"
    $color = switch ($Level) {
        "ERROR"   { "Red" }
        "WARN"    { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "Cyan" }
    }
    Write-Host $entry -ForegroundColor $color
    Add-Content -Path $LOG_PATH -Value $entry -Encoding UTF8
}

# ==============================================================================
# INVOKE-OP
# Calls op with an explicit string[] argument array.
# Retries ONCE after re-authentication on session/auth errors.
# ==============================================================================

function Invoke-Op {
    param(
        [string[]]$Arguments,
        [string]$Description = "op"
    )

    $output   = & op @Arguments 2>&1
    $exitCode = $LASTEXITCODE

    if ($exitCode -eq 0) { return $output }

    $outStr = ($output | Out-String)
    if ($outStr -match "401|403|session|locked|sign in|authenticate|token|not signed in") {
        Write-Log "Auth error on [$Description]. Re-authenticating..." "WARN"
        if (-not (Invoke-Reauth)) {
            Write-Log "Re-auth failed. Skipping: $Description" "ERROR"
            return $null
        }
        $output   = & op @Arguments 2>&1
        $exitCode = $LASTEXITCODE
        if ($exitCode -eq 0) {
            Write-Log "Retry succeeded: $Description" "SUCCESS"
            return $output
        }
        Write-Log "Retry failed: $Description | $($output | Out-String)" "ERROR"
        return $null
    }

    Write-Log "op error [$Description]: $outStr" "WARN"
    return $null
}

function Invoke-Reauth {
    $null = op whoami 2>&1
    if ($LASTEXITCODE -eq 0) { return $true }
    Write-Log "Signing in..." "WARN"
    $token = op signin --raw 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    $env:OP_SESSION = $token
    Write-Log "Re-authentication successful." "SUCCESS"
    return $true
}

# ==============================================================================
# FIELD LOOKUP HELPERS
# ==============================================================================

function Get-FieldById {
    param([object[]]$Fields, [string]$Id)
    $f = $Fields | Where-Object {
        $_.PSObject.Properties['id'] -and $_.id -eq $Id
    } | Select-Object -First 1
    if ($f -and $f.PSObject.Properties['value'] -and $null -ne $f.value) {
        return [string]$f.value
    }
    return ""
}

function Get-FieldByLabel {
    param([object[]]$Fields, [string]$Label)
    $f = $Fields | Where-Object {
        $_.PSObject.Properties['label'] -and ($_.label -ieq $Label)
    } | Select-Object -First 1
    if ($f -and $f.PSObject.Properties['value'] -and $null -ne $f.value) {
        return [string]$f.value
    }
    return ""
}

# ==============================================================================
# EXTRACT-CATEGORYFIELDS
# ==============================================================================

function Extract-CategoryFields {
    param(
        [Parameter(Mandatory)][object]$Detail,
        [Parameter(Mandatory)][string]$Category
    )

    $fixed    = [ordered]@{}
    $extras   = [ordered]@{}
    $fields   = if ($Detail.PSObject.Properties['fields'] -and $Detail.fields) {
        @($Detail.fields)
    } else { @() }

    $consumed = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    function fById ([string]$id) {
        $consumed.Add($id) | Out-Null
        Get-FieldById $fields $id
    }

    function fByLbl ([string]$lbl) {
        $f = $fields | Where-Object {
            $_.PSObject.Properties['label'] -and ($_.label -ieq $lbl)
        } | Select-Object -First 1
        if ($f) {
            if ($f.PSObject.Properties['id'] -and $f.id) {
                $consumed.Add($f.id) | Out-Null
            }
            if ($f.PSObject.Properties['value'] -and $null -ne $f.value) {
                return [string]$f.value
            }
        }
        return ""
    }

    switch ($Category.ToUpper()) {

        "LOGIN" {
            $fixed["Username"]        = fById "username"
            $fixed["Password"]        = fById "password"
            $fixed["TOTP_URI"]        = ""
            $fixed["OneTimePassword"] = ""

            $otpField = $fields | Where-Object {
                ($_.PSObject.Properties['type'] -and $_.type -eq "OTP") -or
                ($_.PSObject.Properties['id']   -and $_.id   -match "^TOTP")
            } | Select-Object -First 1

            if ($otpField) {
                if ($otpField.PSObject.Properties['id'] -and $otpField.id) {
                    $consumed.Add($otpField.id) | Out-Null
                }
                $fixed["TOTP_URI"] = if (
                    $otpField.PSObject.Properties['value'] -and $otpField.value
                ) { [string]$otpField.value } else { "" }

                $fixed["OneTimePassword"] = if (
                    $otpField.PSObject.Properties['totp'] -and $otpField.totp
                ) { [string]$otpField.totp } else { "" }
            }
        }

        "PASSWORD"    { $fixed["Password"] = fById "password" }
        "SECURE_NOTE" { }

        "DOCUMENT" {
            $fixed["FileName"] = fByLbl "filename"
            $fixed["FileSize"] = fByLbl "file size"
        }

        "API_CREDENTIAL" {
            $fixed["Username"]   = fById  "username"
            $fixed["Credential"] = fById  "credential"
            $fixed["APIType"]    = fByLbl "type"
            $fixed["Hostname"]   = fByLbl "hostname"
            $fixed["ValidFrom"]  = fByLbl "valid from"
            $fixed["Expires"]    = fByLbl "expires"
            $fixed["Filename"]   = fByLbl "filename"
        }

        "AMAZON_WEB_SERVICES" {
            $fixed["AccessKeyID"]     = fById  "access_key_id"
            $fixed["SecretAccessKey"] = fById  "secret_access_key"
            $fixed["DefaultRegion"]   = fByLbl "default region"
            $fixed["ConsoleURL"]      = fByLbl "console dashboard url"
            $fixed["MFASerial"]       = fByLbl "mfa serial"
            $fixed["AccountID"]       = fByLbl "account id"
        }

        "BANK_ACCOUNT" {
            $fixed["BankName"]      = fById  "bankName"
            $fixed["OwnerName"]     = fById  "owner"
            $fixed["AccountType"]   = fById  "accountType"
            $fixed["RoutingNumber"] = fById  "routingNo"
            $fixed["AccountNumber"] = fById  "accountNo"
            $fixed["SWIFT"]         = fById  "swift"
            $fixed["IBAN"]          = fById  "iban"
            $fixed["PIN"]           = fById  "telephonePin"
            $fixed["PhoneLocal"]    = fByLbl "phone (local)"
            $fixed["PhoneTollFree"] = fByLbl "phone (toll free)"
            $fixed["PhoneIntl"]     = fByLbl "phone (intl)"
            $fixed["BranchAddress"] = fByLbl "branch address"
            $fixed["BranchPhone"]   = fByLbl "branch phone"
        }

        "CREDIT_CARD" {
            $fixed["CardholderName"]     = fById  "cardholder"
            $fixed["CardType"]           = fById  "type"
            $fixed["CardNumber"]         = fById  "ccnum"
            $fixed["VerificationNumber"] = fById  "cvv"
            $fixed["ExpiryDate"]         = fById  "expiry"
            $fixed["ValidFrom"]          = fById  "validFrom"
            $fixed["CreditLimit"]        = fByLbl "credit limit"
            $fixed["CashLimit"]          = fByLbl "cash withdrawal limit"
            $fixed["InterestRate"]       = fByLbl "interest rate"
            $fixed["IssueNumber"]        = fByLbl "issue number"
            $fixed["PIN"]                = fById  "pin"
            $fixed["IssuingBank"]        = fByLbl "issuing bank"
            $fixed["PhoneLocal"]         = fByLbl "phone (local)"
            $fixed["PhoneTollFree"]      = fByLbl "phone (toll free)"
            $fixed["PhoneIntl"]          = fByLbl "phone (intl)"
            $fixed["Website"]            = fByLbl "website"
        }

        "CRYPTO_WALLET" {
            $fixed["RecoveryPhrase"] = fByLbl "recovery phrase"
            $fixed["Password"]       = fById  "password"
            $fixed["WalletAddress"]  = fByLbl "wallet address"
            $fixed["Username"]       = fById  "username"
        }

        "DATABASE" {
            $fixed["DBType"]            = fById "database_type"
            $fixed["Hostname"]          = fById "hostname"
            $fixed["Port"]              = fById "port"
            $fixed["Database"]          = fById "database"
            $fixed["Username"]          = fById "username"
            $fixed["Password"]          = fById "password"
            $fixed["SID"]               = fById "sid"
            $fixed["Alias"]             = fById "alias"
            $fixed["ConnectionOptions"] = fById "options"
        }

        "DRIVER_LICENSE" {
            $fixed["FullName"]      = fById "fullname"
            $fixed["Sex"]           = fById "sex"
            $fixed["LicenseNumber"] = fById "number"
            $fixed["DateOfBirth"]   = fById "birthdate"
            $fixed["ExpiryDate"]    = fById "expiry_date"
            $fixed["LicenseClass"]  = fById "class"
            $fixed["Conditions"]    = fById "conditions"
            $fixed["State"]         = fById "state"
            $fixed["Country"]       = fById "country"
            $fixed["Address"]       = fById "address"
        }

        "EMAIL_ACCOUNT" {
            $fixed["Username"]        = fById  "username"
            $fixed["Password"]        = fById  "password"
            $fixed["MailFrom"]        = fByLbl "email"
            $fixed["Provider"]        = fByLbl "provider"
            $fixed["ProviderWebsite"] = fByLbl "provider's website"
            $fixed["AuthMethod"]      = fByLbl "auth method"
            $fixed["Security"]        = fByLbl "security"
            $fixed["IMAPServer"]      = fByLbl "imap server"
            $fixed["IMAPPort"]        = fByLbl "imap port"
            $fixed["SMTPServer"]      = fByLbl "smtp server"
            $fixed["SMTPPort"]        = fByLbl "smtp port"
            $fixed["POPServer"]       = fByLbl "pop server"
            $fixed["POPPort"]         = fByLbl "pop port"
        }

        "IDENTITY" {
            $fixed["FirstName"]     = fById  "firstname"
            $fixed["Initial"]       = fById  "initial"
            $fixed["LastName"]      = fById  "lastname"
            $fixed["Sex"]           = fById  "sex"
            $fixed["BirthDate"]     = fById  "birthdate"
            $fixed["Occupation"]    = fById  "occupation"
            $fixed["Company"]       = fById  "company"
            $fixed["Department"]    = fById  "department"
            $fixed["JobTitle"]      = fById  "jobtitle"
            $fixed["Street"]        = fById  "address"
            $fixed["City"]          = fByLbl "city"
            $fixed["State"]         = fByLbl "state"
            $fixed["ZipCode"]       = fByLbl "zip / postal code"
            $fixed["Country"]       = fByLbl "country"
            $fixed["DefaultPhone"]  = fByLbl "default phone"
            $fixed["HomePhone"]     = fByLbl "home"
            $fixed["CellPhone"]     = fByLbl "cell"
            $fixed["BusinessPhone"] = fByLbl "business"
            $fixed["DefaultEmail"]  = fByLbl "default email"
            $fixed["HomeEmail"]     = fByLbl "home email"
            $fixed["BusinessEmail"] = fByLbl "business email"
            $fixed["Website"]       = fByLbl "website"
            $fixed["Username"]      = fByLbl "username"
            $fixed["Reminder"]      = fByLbl "reminder"
        }

        "MEDICAL_RECORD" {
            $fixed["Date"]                   = fByLbl "date"
            $fixed["Location"]               = fByLbl "location"
            $fixed["HealthcareProfessional"] = fByLbl "healthcare professional"
            $fixed["PatientName"]            = fByLbl "patient"
            $fixed["ReasonForVisit"]         = fByLbl "reason for visit"
            $fixed["Medication"]             = fByLbl "medication"
            $fixed["Dosage"]                 = fByLbl "dosage"
        }

        "MEMBERSHIP" {
            $fixed["GroupOrOrganization"] = fById "org_name"
            $fixed["MemberName"]          = fById "member_name"
            $fixed["MemberID"]            = fById "membership_no"
            $fixed["MemberSince"]         = fById "member_since"
            $fixed["ExpiryDate"]          = fById "expiry_date"
            $fixed["Website"]             = fById "website"
            $fixed["Phone"]               = fById "phone"
            $fixed["PIN"]                 = fById "pin"
        }

        "OUTDOOR_LICENSE" {
            $fixed["FullName"]         = fById "name"
            $fixed["ValidFrom"]        = fById "valid_from"
            $fixed["ExpiryDate"]       = fById "expires"
            $fixed["ApprovedWildlife"] = fById "approved_wildlife"
            $fixed["MaximumQuota"]     = fById "max_catch"
            $fixed["State"]            = fById "state"
            $fixed["Country"]          = fById "country"
        }

        "PASSPORT" {
            $fixed["IssuingCountry"] = fById "issuing_country"
            $fixed["FullName"]       = fById "fullname"
            $fixed["Sex"]            = fById "sex"
            $fixed["Nationality"]    = fById "nationality"
            $fixed["IssueDate"]      = fById "issue_date"
            $fixed["ExpiryDate"]     = fById "expiry_date"
            $fixed["PassportNumber"] = fById "number"
            $fixed["PlaceOfBirth"]   = fById "birthplace"
            $fixed["DateOfBirth"]    = fById "birthdate"
        }

        "REWARD_PROGRAM" {
            $fixed["CompanyName"]   = fById  "company_name"
            $fixed["MemberName"]    = fById  "member_name"
            $fixed["MemberID"]      = fById  "membership_no"
            $fixed["PIN"]           = fById  "pin"
            $fixed["MemberSince"]   = fByLbl "member since"
            $fixed["ExpiryDate"]    = fByLbl "expiry date"
            $fixed["Website"]       = fById  "website"
            $fixed["PhoneLocal"]    = fByLbl "phone (local)"
            $fixed["PhoneTollFree"] = fByLbl "phone (toll free)"
            $fixed["PhoneIntl"]     = fByLbl "phone (intl)"
        }

        "SERVER" {
            $fixed["URL"]             = fByLbl "url"
            $fixed["Username"]        = fById  "username"
            $fixed["Password"]        = fById  "password"
            $fixed["HostingProvider"] = fByLbl "hosting provider"
            $fixed["AdminConsole"]    = fByLbl "admin console url"
            $fixed["AdminUser"]       = fByLbl "admin console username"
            $fixed["AdminPass"]       = fByLbl "admin console password"
            $fixed["SupportURL"]      = fByLbl "support url"
            $fixed["SupportPhone"]    = fByLbl "support phone"
        }

        "SOCIAL_SECURITY_NUMBER" {
            $fixed["Name"]   = fById "name"
            $fixed["Number"] = fById "number"
        }

        "SOFTWARE_LICENSE" {
            $fixed["LicenseKey"]      = fById  "license_key"
            $fixed["Version"]         = fByLbl "version"
            $fixed["LicensedTo"]      = fByLbl "licensed to"
            $fixed["RegisteredEmail"] = fByLbl "registered email"
            $fixed["Company"]         = fByLbl "company"
            $fixed["Publisher"]       = fByLbl "publisher"
            $fixed["Website"]         = fByLbl "website"
            $fixed["OrderNumber"]     = fByLbl "order number"
            $fixed["OrderDate"]       = fByLbl "order date"
            $fixed["OrderTotal"]      = fByLbl "retail price"
            $fixed["SupportEmail"]    = fByLbl "support email"
        }

        "SSH_KEY" {
            $fixed["PublicKey"]   = fById  "public_key"
            $fixed["PrivateKey"]  = fById  "private_key"
            $fixed["Fingerprint"] = fById  "fingerprint"
            $fixed["KeyType"]     = fByLbl "key type"
            $fixed["Passphrase"]  = fByLbl "passphrase"
        }

        "WIRELESS_ROUTER" {
            $fixed["BaseStationName"]     = fByLbl "base station name"
            $fixed["BaseStationPassword"] = fByLbl "base station password"
            $fixed["AirportID"]           = fByLbl "airport id"
            $fixed["NetworkName"]         = fByLbl "network name"
            $fixed["WirelessPassword"]    = fByLbl "wireless password"
            $fixed["WirelessSecurity"]    = fByLbl "wireless security"
            $fixed["RouterIP"]            = fByLbl "server / ip address"
            $fixed["AttachmentPassword"]  = fByLbl "attached storage password"
        }

        default {
            $fixed["Username"] = fById "username"
            $fixed["Password"] = fById "password"
        }
    }

    # ── COMMON: Notes
    $notesField = $fields | Where-Object {
        $_.PSObject.Properties['id'] -and $_.id -eq "notesPlain"
    } | Select-Object -First 1

    $fixed["Notes"] = if (
        $notesField -and
        $notesField.PSObject.Properties['value'] -and
        $null -ne $notesField.value -and
        [string]$notesField.value -ne ""
    ) {
        ([string]$notesField.value) -replace "`r`n"," " -replace "`n"," " -replace "`r"," "
    } else { "" }

    $consumed.Add("notesPlain") | Out-Null

    # ── COMMON: Additional OTP fields
    foreach ($f in $fields) {
        if (-not ($f.PSObject.Properties['id'] -and $f.id)) { continue }
        if ($consumed.Contains($f.id)) { continue }

        $isOtp = ($f.PSObject.Properties['type'] -and $f.type -eq "OTP") -or
                 ($f.id -match "^TOTP")
        if (-not $isOtp) { continue }

        $lbl = if ($f.PSObject.Properties['label'] -and $f.label) { $f.label } else { $f.id }
        $uri = if ($f.PSObject.Properties['value'] -and $f.value) { [string]$f.value } else { "" }
        $otp = if ($f.PSObject.Properties['totp']  -and $f.totp)  { [string]$f.totp  } else { "" }

        $extras["TOTP_${lbl}_URI"]  = $uri
        $extras["TOTP_${lbl}_Code"] = $otp
        $consumed.Add($f.id) | Out-Null
    }

    # ── DYNAMIC EXTRAS
    foreach ($f in $fields) {
        $fId  = if ($f.PSObject.Properties['id']    -and $f.id)    { [string]$f.id }    else { "" }
        $fLbl = if ($f.PSObject.Properties['label'] -and $f.label) { [string]$f.label } else { $fId }
        $fVal = if ($f.PSObject.Properties['value'] -and $null -ne $f.value) {
            [string]$f.value
        } else { "" }

        if ($fId -ne "" -and $consumed.Contains($fId)) { continue }
        if ($fId -eq "" -and $fLbl -eq "")             { continue }

        $prefix = ""
        if ($f.PSObject.Properties['section'] -and $f.section -and
            $f.section.PSObject.Properties['label'] -and $f.section.label -and
            [string]$f.section.label -ne "") {
            $prefix = "$($f.section.label)_"
        }

        $colName = "$prefix$fLbl"

        $fixedMatch = $fixed.Keys | Where-Object { $_ -ieq $colName } | Select-Object -First 1
        if ($fixedMatch) { $colName = "${colName}_raw" }

        $colKey = $colName
        $n = 2
        while ($extras.Contains($colKey)) { $colKey = "${colName}_$n"; $n++ }

        $extras[$colKey] = $fVal
        if ($fId -ne "") { $consumed.Add($fId) | Out-Null }
    }

    return [PSCustomObject]@{ Fixed = $fixed; Extras = $extras }
}

# ==============================================================================
# WRITE-CATEGORYCSV
# ==============================================================================

function Write-CategoryCsv {
    param(
        [string]$CsvPath,
        [System.Collections.Generic.List[hashtable]]$Rows,
        [System.Collections.Generic.List[string]]$ExtraKeys,
        [string]$CatKey,
        [ref]$WrittenCount
    )

    if ($Rows.Count -eq 0) {
        Write-Log "  No items for '$CatKey' — skipping CSV." "WARN"
        return
    }

    $fixedCols = if ($script:CAT_COLS.ContainsKey($CatKey.ToUpper())) {
        $script:CAT_COLS[$CatKey.ToUpper()]
    } else { @() }

    $allCols = [System.Collections.Generic.List[string]]::new()
    foreach ($c in $script:IDENTITY_COLS) { $allCols.Add($c) }
    foreach ($c in $fixedCols)            { if (-not $allCols.Contains($c)) { $allCols.Add($c) } }
    foreach ($c in $ExtraKeys)            { if (-not $allCols.Contains($c)) { $allCols.Add($c) } }

    $objects = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($row in $Rows) {
        $obj = New-Object PSCustomObject
        foreach ($col in $allCols) {
            $val = if ($row.Contains($col)) { $row[$col] } else { "" }
            $obj | Add-Member -MemberType NoteProperty -Name $col -Value $val
        }
        $objects.Add($obj)
    }

    $objects | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    Write-Log ("  [OK] {0} -> {1}  ({2} rows, {3} cols)" -f `
        $CatKey, (Split-Path $CsvPath -Leaf), $Rows.Count, $allCols.Count) "SUCCESS"

    $WrittenCount.Value++
}

# ==============================================================================
# INITIALISE OUTPUT DIRECTORIES
# ==============================================================================

New-Item -ItemType Directory -Path $OUTPUT_DIR      -Force | Out-Null
New-Item -ItemType Directory -Path $ATTACHMENTS_DIR -Force | Out-Null

Write-Log "=================================================================" "SUCCESS"
Write-Log "  1Password Full Vault Export v3.0 — Started" "SUCCESS"
Write-Log "  Output : $OUTPUT_DIR" "INFO"
Write-Log "  Archive: $INCLUDE_ARCHIVE" "INFO"
Write-Log "=================================================================" "SUCCESS"

# ==============================================================================
# AUTHENTICATION CHECK
# ==============================================================================

Write-Log "Checking authentication..."
$null = op whoami 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Log "Not authenticated. Attempting sign-in..." "WARN"
    if (-not (Invoke-Reauth)) {
        Write-Log "Authentication failed. Ensure the 1Password app is open and unlocked." "ERROR"
        exit 1
    }
}
Write-Log "Authenticated OK." "SUCCESS"

# ==============================================================================
# VAULT SELECTION
# ==============================================================================

Write-Log "Fetching vault list..."
$vaultListRaw = Invoke-Op -Arguments @("vault", "list", "--format", "json") `
                          -Description "vault list"

if ($null -eq $vaultListRaw) {
    Write-Log "Failed to retrieve vault list. Exiting." "ERROR"
    exit 1
}

$allVaults = $vaultListRaw | ConvertFrom-Json

if ($VAULT_NAME -ne "") {
    $vaults = @($allVaults | Where-Object { $_.name -eq $VAULT_NAME })
    if ($vaults.Count -eq 0) {
        Write-Log "Vault '$VAULT_NAME' not found. Available vaults:" "ERROR"
        $allVaults | ForEach-Object { Write-Log "  - $($_.name)  [$($_.id)]" }
        exit 1
    }
} else {
    $vaults = $allVaults
}

Write-Log "Vaults to process: $($vaults.Count)"
foreach ($v in $vaults) { Write-Log "  -> $($v.name)  [$($v.id)]" }

# ==============================================================================
# MAIN EXPORT LOOP
# ==============================================================================

$totalVaults     = $vaults.Count
$totalItems      = 0
$totalFiles      = 0
$errorCount      = 0
$csvFilesWritten = 0

foreach ($vault in $vaults) {
    $vaultName = $vault.name
    $vaultId   = $vault.id
    $safeVault = $vaultName -replace '[\\/:*?"<>|]', '_'
    $vaultDir  = "$OUTPUT_DIR\$safeVault"

    New-Item -ItemType Directory -Path $vaultDir -Force | Out-Null

    Write-Log ""
    Write-Log "━━━━  Vault: '$vaultName'  [$vaultId]  ━━━━" "SUCCESS"

    $listArgs = @("item", "list", "--vault", $vaultId, "--format", "json")
    if ($INCLUDE_ARCHIVE) { $listArgs += "--include-archive" }

    $itemListRaw = Invoke-Op -Arguments $listArgs `
                             -Description "item list [$vaultName]"
    if ($null -eq $itemListRaw) {
        Write-Log "Failed to list items in '$vaultName'. Skipping vault." "ERROR"
        $errorCount++
        continue
    }

    $items     = @($itemListRaw | ConvertFrom-Json)
    $itemCount = $items.Count
    Write-Log "Found $itemCount item(s)  (include-archive=$INCLUDE_ARCHIVE)"

    $catData   = @{}
    $itemIndex = 0

    foreach ($item in $items) {
        $itemIndex++
        $itemId    = $item.id
        $itemTitle = $item.title
        $catKey    = ($item.category).ToUpper()
        $totalItems++

        Write-Log ("  [{0}/{1}] [{2}] '{3}'" -f $itemIndex, $itemCount, $catKey, $itemTitle)

        $detailRaw = Invoke-Op `
            -Arguments   @("item", "get", $itemId, "--vault", $vaultId, "--format", "json") `
            -Description "item get '$itemTitle' [$itemId]"

        if ($null -eq $detailRaw) {
            Write-Log "  FAILED to retrieve '$itemTitle' — skipping." "ERROR"
            $errorCount++
            continue
        }

        $detail = $detailRaw | ConvertFrom-Json

        $tags = if ($detail.PSObject.Properties['tags'] -and $detail.tags) {
            $detail.tags -join "; "
        } else { "" }

        $createdAt = if ($detail.PSObject.Properties['created_at'] -and $detail.created_at) {
            [string]$detail.created_at
        } else { "" }

        $updatedAt = if ($detail.PSObject.Properties['updated_at'] -and $detail.updated_at) {
            [string]$detail.updated_at
        } else { "" }

        $urlList = [System.Collections.Generic.List[string]]::new()
        if ($detail.PSObject.Properties['urls'] -and $detail.urls) {
            foreach ($u in $detail.urls) {
                if ($u.PSObject.Properties['href'] -and $u.href) {
                    $urlList.Add([string]$u.href)
                }
            }
        }
        $urlStr = $urlList -join "; "

        $extracted = Extract-CategoryFields -Detail $detail -Category $catKey
        $catFixed  = $extracted.Fixed
        $catExtras = $extracted.Extras

        $savedFiles    = [System.Collections.Generic.List[string]]::new()
        $safeTitle     = $itemTitle -replace '[\\/:*?"<>|]', '_'
        $attachItemDir = "$ATTACHMENTS_DIR\$safeVault\${safeTitle}_$itemId"

        if ($catKey -eq "DOCUMENT") {
            New-Item -ItemType Directory -Path $attachItemDir -Force | Out-Null

            $docName = if ($catFixed["FileName"] -ne "") {
                $catFixed["FileName"] -replace '[\\/:*?"<>|]', '_'
            } else { $safeTitle }

            $docPath = "$attachItemDir\$docName"

            Write-Log "    Downloading document body: $docName"
            $null = Invoke-Op `
                -Arguments   @("document", "get", $itemId, "--vault", $vaultId, "--out-file", $docPath) `
                -Description "document get body '$itemTitle'"

            if (Test-Path $docPath) {
                $savedFiles.Add($docName)
                $totalFiles++
                Write-Log "    Saved: $docPath" "SUCCESS"
            } else {
                Write-Log "    FAILED: document body for '$itemTitle'" "ERROR"
                $errorCount++
            }
        }

        if ($detail.PSObject.Properties['files'] -and $detail.files -and
            $detail.files.Count -gt 0) {

            New-Item -ItemType Directory -Path $attachItemDir -Force | Out-Null

            foreach ($file in $detail.files) {
                $fileId   = $file.id
                $fileName = if ($file.PSObject.Properties['name'] -and $file.name) {
                    $file.name -replace '[\\/:*?"<>|]', '_'
                } else { $fileId }

                $filePath = "$attachItemDir\$fileName"

                Write-Log "    Downloading attachment: $fileName"
                $null = Invoke-Op `
                    -Arguments   @("document", "get", $fileId, "--vault", $vaultId, "--out-file", $filePath) `
                    -Description "document get attachment '$fileName'"

                if (Test-Path $filePath) {
                    $savedFiles.Add($fileName)
                    $totalFiles++
                    Write-Log "    Saved: $filePath" "SUCCESS"
                } else {
                    Write-Log "    FAILED: attachment '$fileName' on '$itemTitle'" "ERROR"
                    $errorCount++
                }
            }
        }

        $row = [ordered]@{
            "Vault"       = $vaultName
            "ItemID"      = $itemId
            "Title"       = $itemTitle
            "Category"    = $item.category
            "URLs"        = $urlStr
            "Tags"        = $tags
            "Notes"       = if ($catFixed.Contains("Notes")) { $catFixed["Notes"] } else { "" }
            "Attachments" = ($savedFiles -join "; ")
            "CreatedAt"   = $createdAt
            "UpdatedAt"   = $updatedAt
        }

        foreach ($k in $catFixed.Keys)  { if ($k -eq "Notes") { continue }; $row[$k] = $catFixed[$k] }
        foreach ($k in $catExtras.Keys) { $row[$k] = $catExtras[$k] }

        if (-not $catData.ContainsKey($catKey)) {
            $catData[$catKey] = @{
                Rows      = [System.Collections.Generic.List[hashtable]]::new()
                ExtraKeys = [System.Collections.Generic.List[string]]::new()
            }
        }

        $catData[$catKey].Rows.Add($row)

        foreach ($k in $catExtras.Keys) {
            if (-not $catData[$catKey].ExtraKeys.Contains($k)) {
                $catData[$catKey].ExtraKeys.Add($k)
            }
        }
    }

    Write-Log ""
    Write-Log "Writing CSVs for vault '$vaultName'..."

    $writtenRef = [ref]$csvFilesWritten

    foreach ($catKey in ($catData.Keys | Sort-Object)) {
        $csvName = if ($script:CAT_FILE.ContainsKey($catKey)) {
            $script:CAT_FILE[$catKey]
        } else {
            "Other_$($catKey -replace '[\\/:*?"<>|]', '_')"
        }

        Write-CategoryCsv `
            -CsvPath      "$vaultDir\$csvName.csv" `
            -Rows         $catData[$catKey].Rows `
            -ExtraKeys    $catData[$catKey].ExtraKeys `
            -CatKey       $catKey `
            -WrittenCount $writtenRef
    }

    Write-Log "Vault '$vaultName' complete.  CSVs -> $vaultDir" "SUCCESS"
}

# ==============================================================================
# SUMMARY
# ==============================================================================

Write-Log ""
Write-Log "=================================================================" "SUCCESS"
Write-Log "  EXPORT COMPLETE" "SUCCESS"
Write-Log "  Vaults Processed  : $totalVaults"
Write-Log "  Total Items       : $totalItems"
Write-Log "  Files Downloaded  : $totalFiles"
Write-Log "  CSV Files Written : $csvFilesWritten"
Write-Log "  Errors            : $errorCount"
Write-Log "-----------------------------------------------------------------"
Write-Log "  Output Folder     : $OUTPUT_DIR"
Write-Log "  Attachments       : $ATTACHMENTS_DIR"
Write-Log "  Log               : $LOG_PATH"
Write-Log "=================================================================" "SUCCESS"

if ($errorCount -gt 0) {
    Write-Log "One or more errors occurred. Review the log: $LOG_PATH" "WARN"
}
