# hMailServer Data Export Automation

A comprehensive, professional-grade PowerShell 7 utility for extracting domains, accounts, aliases, and distribution lists from a local hMailServer instance. This tool supports multiple industrial-standard export formats, including CSV, JSON, Markdown, JavaScript, PHP, and XML, facilitating seamless integration with web applications, migration tools, and administrative dashboards.

[![PowerShell](https://img.shields.io/badge/PowerShell-7.0%2B-blue.svg)](https://microsoft.com/powershell)
[![hMailServer](https://img.shields.io/badge/hMailServer-5.6%2B-green.svg)](https://www.hmailserver.com)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://microsoft.com/windows)

## 📋 Table of Contents

1. [Features](#1-features)
2. [Prerequisites](#2-prerequisites)
3. [Installation](#3-installation)
4. [Configuration](#4-configuration)
5. [Usage](#5-usage)
    - 5.1 [Basic Syntax](#51-basic-syntax)
    - 5.2 [Export Formats](#52-export-formats)
    - 5.3 [Operation Modes](#53-operation-modes)
    - 5.4 [Examples](#54-examples)
6. [How It Works](#6-how-it-works)
    - 6.1 [COM Interface Integration](#61-com-interface-integration)
    - 6.2 [Output Directory Resolution](#62-output-directory-resolution)
    - 6.3 [Advanced Logging (Mode B)](#63-advanced-logging)
7. [Troubleshooting](#7-troubleshooting)
    - 7.1 [Common Issues](#71-common-issues)
    - 7.2 [Debug Mode](#72-debug-mode)
8. [Contributing](#8-contributing)
9. [License](#9-license)
10. [Acknowledgments](#10-acknowledgments)
11. [Author & Support](#11-author--support)

## 1. Features

- **Multi-Format Export Support**: Simultaneously export data to industrial formats: `CSV`, `JSON`, `Markdown`, `JS`, `PHP`, and `XML`.
- **Comprehensive Data Retrieval**: Extracts accounts, aliases, distribution lists, and list members with granular detail.
- **Dynamic Directory Resolution**: Intelligent three-tier priority system for managing output locations.
- **Professional Logging**: Detailed, color-coded console output with optional persistence to log files.
- **Security-First Approach**: Secure handling of hMailServer administrator credentials with interactive prompt support.
- **Robust Error Handling**: Built-in validation for COM object availability, write permissions, and filesystem paths.
- **Automated Directory Management**: Automatically verifies and creates target directories if they are missing.
- **Advanced Mode (Mode B)**: Provides deep-dive extraction of distribution list members and full execution tracing.

## 2. Prerequisites

- **Operating System**: Windows (tested on Windows Server 2019/2022, Windows 10/11)
- **PowerShell**: Version 7.0 or higher (Cross-platform Core)
- **hMailServer**: Local installation with the COM API enabled (`hMailServer.Application` object)
- **Permissions**: Administrator rights on the local machine (to interact with hMailServer COM)
- **Dependencies**: 
    - Access to the hMailServer Administrator password.
    - Write permissions to the target output directory.

## 3. Installation

### 3.1 Download the Script

Clone the repository to your local administrative workstation:

```powershell
git clone https://github.com/paulmann/PowerShell_Export-hMailServerData.git
cd PowerShell_Export-hMailServerData
```

### 3.2 Verify Environment

Ensure that PowerShell 7 is active and that hMailServer service is running:

```powershell
$PSVersionTable.PSVersion # Should be 7.x
Get-Service -Name hMailServer # Should be Running
```

## 4. Configuration

The script uses a prioritized configuration model. You can modify the variables at the top of `Export-hMailServerData.ps1` for persistent settings.

### 4.1 Global Settings

```powershell
$HMAILADMINUSER = "Administrator"
$OUTPUTDIR = "C:\Export" # Global fallback directory
```

### 4.2 Format-Specific Directories

You can direct different formats to different locations by setting specific variables:

```powershell
$PHPOUTPUTDIR = "C:\Web\Includes"
$JSONOUTPUTDIR = "C:\App\Data"
```

### 4.3 Credentials

If `$HMAILPASSWORD` is left empty, the script will securely prompt for the password at runtime.

## 5. Usage

### 5.1 Basic Syntax

```powershell
./Export-hMailServerData.ps1 [-ExportFormat <Format>] [-Mode <A|B>] [-OutputDir <Path>] [-DebugMode <0|1>]
```

### 5.2 Export Formats

| Format     | Extension | Use Case                                  |
|-----------|-----------|--------------------------------------------|
| `CSV`     | `.csv`    | Spreadsheet analysis / Bulk imports        |
| `JSON`    | `.json`   | Web APIs / Modern application data         |
| `Markdown`| `.md`     | Documentation / GitHub Wiki content        |
| `JS`      | `.js`     | JavaScript modules / Frontend integration  |
| `PHP`     | `.php`    | PHP Constants / Backend inclusion          |
| `XML`     | `.xml`    | Legacy systems / Structured data           |

#### 5.2.1 Sample output: Accounts

All formats below show the same logical dataset: two example accounts from `example.com` with the exact field set produced by `Get-DomainAccounts`.[file:1]

| Format | Sample output (excerpt) |
|--------|-------------------------|
| `CSV` | `Domain,Address,Active,IsAD,MaxSizeMB,PasswordExpirationEnabled,PasswordExpires,VacationMessageIsOn,ForwardEnabled,ForwardAddress,SignatureEnabled,AdminLevel,PersonFirstName,PersonLastName,Created,LastLogon`<br>`example.com,user1@example.com,True,False,2048,False,,False,False,,False,0,John,Doe,2026-04-14T10:15:30Z,2026-04-14T11:00:00Z`<br>`example.com,user2@example.com,True,False,1024,False,,False,False,,False,0,Jane,Doe,2026-04-14T10:16:02Z,2026-04-14T11:05:00Z` |
| `JSON` | `[`<br>`  {`<br>`    "Domain": "example.com",`<br>`    "Address": "user1@example.com",`<br>`    "Active": true,`<br>`    "IsAD": false,`<br>`    "MaxSizeMB": 2048,`<br>`    "PasswordExpirationEnabled": false,`<br>`    "PasswordExpires": null,`<br>`    "VacationMessageIsOn": false,`<br>`    "ForwardEnabled": false,`<br>`    "ForwardAddress": null,`<br>`    "SignatureEnabled": false,`<br>`    "AdminLevel": 0,`<br>`    "PersonFirstName": "John",`<br>`    "PersonLastName": "Doe",`<br>`    "Created": "2026-04-14T10:15:30Z",`<br>`    "LastLogon": "2026-04-14T11:00:00Z"`<br>`  },`<br>`  {`<br>`    "Domain": "example.com",`<br>`    "Address": "user2@example.com",`<br>`    "Active": true,`<br>`    "IsAD": false,`<br>`    "MaxSizeMB": 1024,`<br>`    "PasswordExpirationEnabled": false,`<br>`    "PasswordExpires": null,`<br>`    "VacationMessageIsOn": false,`<br>`    "ForwardEnabled": false,`<br>`    "ForwardAddress": null,`<br>`    "SignatureEnabled": false,`<br>`    "AdminLevel": 0,`<br>`    "PersonFirstName": "Jane",`<br>`    "PersonLastName": "Doe",`<br>`    "Created": "2026-04-14T10:16:02Z",`<br>`    "LastLogon": "2026-04-14T11:05:00Z"`<br>`  }`<br>`]` |
| `Markdown` | `Markdown export renders the same accounts dataset as a GitHub‑friendly table like the one you are reading now, with the exact same columns as in the CSV/JSON examples.` |
| `JS` | `/**`<br>` * Auto-generated hMailServer export`<br>` * Format: JavaScript Module`<br>` */`<br>`const example_com_accounts = [`<br>`  {`<br>`    Domain: "example.com",`<br>`    Address: "user1@example.com",`<br>`    Active: true,`<br>`    IsAD: false,`<br>`    MaxSizeMB: 2048,`<br>`    PasswordExpirationEnabled: false,`<br>`    PasswordExpires: null,`<br>`    VacationMessageIsOn: false,`<br>`    ForwardEnabled: false,`<br>`    ForwardAddress: null,`<br>`    SignatureEnabled: false,`<br>`    AdminLevel: 0,`<br>`    PersonFirstName: "John",`<br>`    PersonLastName: "Doe",`<br>`    Created: "2026-04-14T10:15:30Z",`<br>`    LastLogon: "2026-04-14T11:00:00Z"`<br>`  },`<br>`  {`<br>`    Domain: "example.com",`<br>`    Address: "user2@example.com",`<br>`    Active: true,`<br>`    IsAD: false,`<br>`    MaxSizeMB: 1024,`<br>`    PasswordExpirationEnabled: false,`<br>`    PasswordExpires: null,`<br>`    VacationMessageIsOn: false,`<br>`    ForwardEnabled: false,`<br>`    ForwardAddress: null,`<br>`    SignatureEnabled: false,`<br>`    AdminLevel: 0,`<br>`    PersonFirstName: "Jane",`<br>`    PersonLastName: "Doe",`<br>`    Created: "2026-04-14T10:16:02Z",`<br>`    LastLogon: "2026-04-14T11:05:00Z"`<br>`  }`<br>`];`<br>`export default example_com_accounts;` |
| `PHP` | `<?php`<br>`/**`<br>` * Auto-generated hMailServer export`<br>` * Format: PHP Constants`<br>` */`<br>`const DATA_EXAMPLE_COM_ACCOUNTS = [`<br>`  [`<br>`    'Domain' => 'example.com',`<br>`    'Address' => 'user1@example.com',`<br>`    'Active' => true,`<br>`    'IsAD' => false,`<br>`    'MaxSizeMB' => 2048,`<br>`    'PasswordExpirationEnabled' => false,`<br>`    'PasswordExpires' => null,`<br>`    'VacationMessageIsOn' => false,`<br>`    'ForwardEnabled' => false,`<br>`    'ForwardAddress' => null,`<br>`    'SignatureEnabled' => false,`<br>`    'AdminLevel' => 0,`<br>`    'PersonFirstName' => 'John',`<br>`    'PersonLastName' => 'Doe',`<br>`    'Created' => '2026-04-14T10:15:30Z',`<br>`    'LastLogon' => '2026-04-14T11:00:00Z'`<br>`  ],`<br>`  [`<br>`    'Domain' => 'example.com',`<br>`    'Address' => 'user2@example.com',`<br>`    'Active' => true,`<br>`    'IsAD' => false,`<br>`    'MaxSizeMB' => 1024,`<br>`    'PasswordExpirationEnabled' => false,`<br>`    'PasswordExpires' => null,`<br>`    'VacationMessageIsOn' => false,`<br>`    'ForwardEnabled' => false,`<br>`    'ForwardAddress' => null,`<br>`    'SignatureEnabled' => false,`<br>`    'AdminLevel' => 0,`<br>`    'PersonFirstName' => 'Jane',`<br>`    'PersonLastName' => 'Doe',`<br>`    'Created' => '2026-04-14T10:16:02Z',`<br>`    'LastLogon' => '2026-04-14T11:05:00Z'`<br>`  ],`<br>`];`<br>`return DATA_EXAMPLE_COM_ACCOUNTS;` |
| `XML` | `<?xml version="1.0" encoding="UTF-8"?>`<br>`<export>`<br>`  <item>`<br>`    <Domain>example.com</Domain>`<br>`    <Address>user1@example.com</Address>`<br>`    <Active>1</Active>`<br>`    <IsAD>0</IsAD>`<br>`    <MaxSizeMB>2048</MaxSizeMB>`<br>`    <PasswordExpirationEnabled>0</PasswordExpirationEnabled>`<br>`    <PasswordExpires />`<br>`    <VacationMessageIsOn>0</VacationMessageIsOn>`<br>`    <ForwardEnabled>0</ForwardEnabled>`<br>`    <ForwardAddress />`<br>`    <SignatureEnabled>0</SignatureEnabled>`<br>`    <AdminLevel>0</AdminLevel>`<br>`    <PersonFirstName>John</PersonFirstName>`<br>`    <PersonLastName>Doe</PersonLastName>`<br>`    <Created>2026-04-14T10:15:30Z</Created>`<br>`    <LastLogon>2026-04-14T11:00:00Z</LastLogon>`<br>`  </item>`<br>`  <item>`<br>`    <Domain>example.com</Domain>`<br>`    <Address>user2@example.com</Address>`<br>`    <Active>1</Active>`<br>`    <IsAD>0</IsAD>`<br>`    <MaxSizeMB>1024</MaxSizeMB>`<br>`    <PasswordExpirationEnabled>0</PasswordExpirationEnabled>`<br>`    <PasswordExpires />`<br>`    <VacationMessageIsOn>0</VacationMessageIsOn>`<br>`    <ForwardEnabled>0</ForwardEnabled>`<br>`    <ForwardAddress />`<br>`    <SignatureEnabled>0</SignatureEnabled>`<br>`    <AdminLevel>0</AdminLevel>`<br>`    <PersonFirstName>Jane</PersonFirstName>`<br>`    <PersonLastName>Doe</PersonLastName>`<br>`    <Created>2026-04-14T10:16:02Z</Created>`<br>`    <LastLogon>2026-04-14T11:05:00Z</LastLogon>`<br>`  </item>`<br>`</export>` |

#### 5.2.2 Sample output: Aliases

This dataset matches the object shape returned by `Get-DomainAliases` (`Domain`, `Name`, `Address`, `Value`, `Active`, `Created`).[file:1]

Example aliases for `example.com`:

- `sales@example.com` → `user1@example.com`
- `support@example.com` → `user2@example.com`

| Format | Sample output (excerpt) |
|--------|-------------------------|
| `CSV` | `Domain,Name,Address,Value,Active,Created`<br>`example.com,sales,sales@example.com,user1@example.com,True,2026-04-14T10:20:00Z`<br>`example.com,support,support@example.com,user2@example.com,True,2026-04-14T10:21:00Z` |
| `JSON` | `[`<br>`  {`<br>`    "Domain": "example.com",`<br>`    "Name": "sales",`<br>`    "Address": "sales@example.com",`<br>`    "Value": "user1@example.com",`<br>`    "Active": true,`<br>`    "Created": "2026-04-14T10:20:00Z"`<br>`  },`<br>`  {`<br>`    "Domain": "example.com",`<br>`    "Name": "support",`<br>`    "Address": "support@example.com",`<br>`    "Value": "user2@example.com",`<br>`    "Active": true,`<br>`    "Created": "2026-04-14T10:21:00Z"`<br>`  }`<br>`]` |
| `Markdown` | `Markdown export renders the same accounts dataset as a GitHub‑friendly table like the one you are reading now, with the exact same columns as in the CSV/JSON examples.` |
| `JS` | `const example_com_aliases = [`<br>`  { Domain: "example.com", Name: "sales", Address: "sales@example.com", Value: "user1@example.com", Active: true, Created: "2026-04-14T10:20:00Z" },`<br>`  { Domain: "example.com", Name: "support", Address: "support@example.com", Value: "user2@example.com", Active: true, Created: "2026-04-14T10:21:00Z" }`<br>`];`<br>`export default example_com_aliases;` |
| `PHP` | `const DATA_EXAMPLE_COM_ALIASES = [`<br>`  ['Domain' => 'example.com', 'Name' => 'sales', 'Address' => 'sales@example.com', 'Value' => 'user1@example.com', 'Active' => true, 'Created' => '2026-04-14T10:20:00Z'],`<br>`  ['Domain' => 'example.com', 'Name' => 'support', 'Address' => 'support@example.com', 'Value' => 'user2@example.com', 'Active' => true, 'Created' => '2026-04-14T10:21:00Z'],`<br>`];` |
| `XML` | `<export>`<br>`  <item>`<br>`    <Domain>example.com</Domain>`<br>`    <Name>sales</Name>`<br>`    <Address>sales@example.com</Address>`<br>`    <Value>user1@example.com</Value>`<br>`    <Active>1</Active>`<br>`    <Created>2026-04-14T10:20:00Z</Created>`<br>`  </item>`<br>`  <item>`<br>`    <Domain>example.com</Domain>`<br>`    <Name>support</Name>`<br>`    <Address>support@example.com</Address>`<br>`    <Value>user2@example.com</Value>`<br>`    <Active>1</Active>`<br>`    <Created>2026-04-14T10:21:00Z</Created>`<br>`  </item>`<br>`</export>` |

#### 5.2.3 Sample output: Distribution list members

In advanced mode (`Mode = 'B'`) the script also exports distribution list members using the shape returned by `Get-DomainDistributionLists().Members` (`Domain`, `DistributionList`, `RecipientAddress`).[file:1]

Example: list `team@example.com` with two recipients.

| Format | Sample output (excerpt) |
|--------|-------------------------|
| `CSV` | `Domain,DistributionList,RecipientAddress`<br>`example.com,team@example.com,user1@example.com`<br>`example.com,team@example.com,user2@example.com` |
| `JSON` | `[`<br>`  {`<br>`    "Domain": "example.com",`<br>`    "DistributionList": "team@example.com",`<br>`    "RecipientAddress": "user1@example.com"`<br>`  },`<br>`  {`<br>`    "Domain": "example.com",`<br>`    "DistributionList": "team@example.com",`<br>`    "RecipientAddress": "user2@example.com"`<br>`  }`<br>`]` |
| `Markdown` | `Markdown export renders the same accounts dataset as a GitHub‑friendly table like the one you are reading now, with the exact same columns as in the CSV/JSON examples.` |
| `JS` | `const example_com_distribution_list_members = [`<br>`  { Domain: "example.com", DistributionList: "team@example.com", RecipientAddress: "user1@example.com" },`<br>`  { Domain: "example.com", DistributionList: "team@example.com", RecipientAddress: "user2@example.com" }`<br>`];`<br>`export default example_com_distribution_list_members;` |
| `PHP` | `const DATA_EXAMPLE_COM_DISTRIBUTION_LIST_MEMBERS = [`<br>`  ['Domain' => 'example.com', 'DistributionList' => 'team@example.com', 'RecipientAddress' => 'user1@example.com'],`<br>`  ['Domain' => 'example.com', 'DistributionList' => 'team@example.com', 'RecipientAddress' => 'user2@example.com'],`<br>`];` |
| `XML` | `<export>`<br>`  <item>`<br>`    <Domain>example.com</Domain>`<br>`    <DistributionList>team@example.com</DistributionList>`<br>`    <RecipientAddress>user1@example.com</RecipientAddress>`<br>`  </item>`<br>`  <item>`<br>`    <Domain>example.com</Domain>`<br>`    <DistributionList>team@example.com</DistributionList>`<br>`    <RecipientAddress>user2@example.com</RecipientAddress>`<br>`  </item>`<br>`</export>` |

### 5.3 Operation Modes

| Mode | Level | Description |
|------|-------|-------------|
| `A` | Standard | Primary data export (Accounts, Aliases, Lists) |
| `B` | Advanced | Includes DL members and detailed file logging |

### 5.4 Examples

**Standard JSON export to default directory:**
```powershell
./Export-hMailServerData.ps1 -ExportFormat JSON
```

**Advanced PHP export with custom output path:**
```powershell
./Export-hMailServerData.ps1 -ExportFormat PHP -Mode B -OutputDir D:\Data
```

**Markdown export with debug logging enabled:**
```powershell
./Export-hMailServerData.ps1 -ExportFormat Markdown -DebugMode 1
```

## 6. How It Works

### 6.1 COM Interface Integration

The script utilizes the `hMailServer.Application` COM object to communicate directly with the mail server engine. It performs authenticated sessions to traverse the object hierarchy:
1. **Domains**: Iterates through all hosted domains.
2. **Accounts**: Collects addresses, quotas, and signature states.
3. **Aliases**: Maps naming redirects.
4. **Distribution Lists**: Resolves list logic and (in Mode B) member relationships.

### 6.2 Output Directory Resolution

The script implements a strict three-tier priority logic to determine where files are saved:
1. **Command Line**: `-OutputDir` parameter (Highest priority).
2. **Format Variable**: e.g., `$PHPOUTPUTDIR` if set in the script.
3. **Global Fallback**: The `$OUTPUTDIR` variable.

### 6.3 Advanced Logging (Mode B)

When running in **Mode B**, the script generates a timestamped log file in the output directory. This log captures every step of the COM object interaction, providing a full audit trail for troubleshooting large-scale exports.

## 7. Troubleshooting

### 7.1 Common Issues

**"Failed to create COM object"**:
- Ensure hMailServer is installed on the local machine.
- Verify the COM library is registered (run hMailServer Administrator once).

**"Authentication Failed"**:
- Check that the `$HMAILADMINUSER` matches your admin login.
- Ensure the password entered (or in `$HMAILPASSWORD`) is correct.

**"Write permission denied"**:
- Run PowerShell as Administrator.
- Verify the NTFS permissions on the target `$OUTPUTDIR`.

### 7.2 Debug Mode

Enable verbose output for deep inspection:

```powershell
./Export-hMailServerData.ps1 -ExportFormat CSV -DebugMode 1
```

This provides:
- COM object lifecycle events.
- Domain-by-domain processing status.
- Serialization and filesystem IO details.

## 8. Contributing

Professional contributions are welcome! If you find a bug or have a feature request:
1. Fork the repository.
2. Create a feature branch.
3. Submit a Pull Request with a clear description of changes.

## 9. License

This project is licensed under the PolyForm Noncommercial License 1.0.0 - see the [LICENSE](LICENSE) file for details.

## 10. Acknowledgments

- **hMailServer Team** for providing a robust COM API for automation.
- **PowerShell Community** for the inspiration on modular script design.

## 11. Author & Support

**Mikhail Deynekin**
- 🌐 **Website**: [deynekin.com](https://deynekin.com)
- 📧 **Email**: [mid1977@gmail.com](mailto:mid1977@gmail.com)
- 🐙 **GitHub**: [@paulmann](https://github.com/paulmann)

### Getting Help

- 📖 **Documentation**: Read this README thoroughly.
- 🐛 **Bug Reports**: [Open an issue](https://github.com/paulmann/PowerShell_Export-hMailServerData/issues/new)
- 💡 **Feature Requests**: [Request features](https://github.com/paulmann/PowerShell_Export-hMailServerData/issues/new)

---

**Professional Note**: Always verify your export results in a staging environment before integrating with production databases or web services.
