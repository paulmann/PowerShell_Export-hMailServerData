# hMailServer Data Export Automation

A comprehensive, professional-grade PowerShell 7 utility for extracting domains, accounts, aliases, and distribution lists from a local hMailServer instance. This tool supports multiple industrial-standard export formats, including CSV, JSON, Markdown, JavaScript, PHP, and XML, facilitating seamless integration with web applications, migration tools, and administrative dashboards.

[![PowerShell](https://img.shields.io/badge/PowerShell-7.0%2B-blue.svg)](https://microsoft.com/powershell)
[![hMailServer](https://img.shields.io/badge/hMailServer-5.6%2B-green.svg)](https://www.hmailserver.com)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
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

| Format | Extension | Use Case |
|--------|-----------|----------|
| `CSV` | `.csv` | Spreadsheet analysis / Bulk imports |
| `JSON` | `.json` | Web APIs / Modern application data |
| `Markdown` | `.md` | Documentation / GitHub Wiki content |
| `JS` | `.js` | JavaScript modules / Frontend integration |
| `PHP` | `.php` | PHP Constants / Backend inclusion |
| `XML` | `.xml` | Legacy systems / Structured data |

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

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

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
