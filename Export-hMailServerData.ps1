<#
.SYNOPSIS
    Export hMailServer domains, accounts, aliases, and distribution lists to multiple file formats.

.DESCRIPTION
    This script connects to a local hMailServer installation via COM interface and exports:
    - All mailboxes per domain to [domain]_accounts.[ext]
    - All aliases per domain to [domain]_aliases.[ext]
    - All distribution lists per domain to [domain]_distribution-lists.[ext]

    Additionally, in Advanced mode (Mode = 'B') it exports:
    - All distribution list members per domain to [domain]_distribution-list-members.[ext]
    - Detailed log to a separate .log file

    Supported export formats: CSV, JSON, Markdown, JavaScript (JS), PHP constants, XML.
    
    OUTPUT DIRECTORY RESOLUTION (in order of priority):
    1. Command-line parameter -OutputDir (for current export format only)
    2. Script variable named $<FORMAT>_OUTPUT_DIR (e.g., $PHP_OUTPUT_DIR when -ExportFormat PHP)
    3. Global fallback variable $OUTPUT_DIR

.AUTHOR
    Mikhail Deynekin <mikhail@deynekin.com>

.SITE
    https://deynekin.com

.LINK
    https://github.com/paulmann

.REQUIREMENTS
    - PowerShell 7.0 or higher
    - hMailServer installed locally (COM object 'hMailServer.Application' must be available)
    - Administrator credentials for hMailServer
    - Write permissions to the target output directory

.EXAMPLE
    .\Export-hMailServerData.ps1 -ExportFormat PHP
    # Uses $PHP_OUTPUT_DIR if set, otherwise falls back to $OUTPUT_DIR

.EXAMPLE
    .\Export-hMailServerData.ps1 -ExportFormat PHP -OutputDir "C:\inetpub\wwwroot\"
    # Forces output to C:\inetpub\wwwroot\ regardless of script variables

.EXAMPLE
    .\Export-hMailServerData.ps1 -ExportFormat JSON -Mode B -DebugMode 0
    # Exports JSON to $JSON_OUTPUT_DIR (or $OUTPUT_DIR) in advanced mode

.NOTES
    Version: 2.1.1
    Date: 2026-04-14
    License: MIT
    Changelog:
        - 2.1.1: XML/Markdown fix
        - 2.1.0: Simplified output directory logic: single -OutputDir param + format-specific vars
        - 2.0.0: Added multi-format export support (CSV/JSON/MD/JS/PHP/XML)
        - 1.2.0: Initial version with CSV-only export
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet(0,1)]
    [int]$DebugMode = 1,

    [Parameter(Mandatory = $false)]
    [ValidateSet('A','B')]
    [string]$Mode = 'B',

    [Parameter(Mandatory = $false)]
    [ValidateSet('CSV','JSON','Markdown','JS','PHP','XML')]
    [string]$ExportFormat = 'CSV',

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if (-not [System.IO.Path]::IsPathRooted($_)) {
            throw "OutputDir must be an absolute path. Received: '$_'"
        }
        return $true
    })]
    [string]$OutputDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =============================================================================
# CONSTANTS / CONFIGURATION
# =============================================================================
$DEBUG            = $DebugMode
$HMAIL_PASSWORD   = ''
$OUTPUT_DIR       = 'C:\mail\Logs'  # Global fallback for all formats
$HMAIL_ADMIN_USER = 'Administrator'

# Format-specific output directories (override $OUTPUT_DIR per format)
# Set to $null or empty string to use global $OUTPUT_DIR
$CSV_OUTPUT_DIR       = $null  # e.g., 'C:\exports\csv'
$JSON_OUTPUT_DIR      = $null  # e.g., 'C:\exports\json'
$Markdown_OUTPUT_DIR  = $null  # e.g., 'C:\exports\markdown'
$JS_OUTPUT_DIR        = $null  # e.g., 'C:\exports\javascript'
$PHP_OUTPUT_DIR       = $null  # e.g., 'C:\exports\php'
$XML_OUTPUT_DIR       = $null  # e.g., 'C:\exports\xml'

# Format-specific configuration (extensions, encoders, etc.)
$FormatConfig = @{
    CSV      = @{ Extension = '.csv'; ContentType = 'text/csv'; Encoder = 'UTF8'; BOM = $false }
    JSON     = @{ Extension = '.json'; ContentType = 'application/json'; Encoder = 'UTF8'; Depth = 10; BOM = $false }
    Markdown = @{ Extension = '.md'; ContentType = 'text/markdown'; Encoder = 'UTF8'; BOM = $false }
    JS       = @{ Extension = '.js'; ContentType = 'application/javascript'; Encoder = 'UTF8'; BOM = $true }
    PHP      = @{ Extension = '.php'; ContentType = 'application/x-httpd-php'; Encoder = 'UTF8'; BOM = $true }
    XML      = @{ Extension = '.xml'; ContentType = 'application/xml'; Encoder = 'UTF8'; Depth = 10; BOM = $false }
}

# =============================================================================
# ENCODING HELPER
# =============================================================================
function Get-FileEncoding {
    [CmdletBinding()]
    [OutputType([System.Text.UTF8Encoding])]
    param(
        [Parameter(Mandatory = $false)]
        [bool]$WithBom = $false
    )
    
    # .NET 8+/PowerShell 7.4: UTF8Encoding(bool encoderShouldEmitUTF8Identifier, bool throwOnInvalidBytes)
    return [System.Text.UTF8Encoding]::new($WithBom, $false)
}


# =============================================================================
# DIRECTORY RESOLUTION ENGINE
# =============================================================================

function Resolve-OutputDirectory {
    <#
    .SYNOPSIS
        Resolves the effective output directory for the current export format.
    
    .DESCRIPTION
        Implements a simple three-tier priority system:
        1. Command-line parameter -OutputDir (if provided)
        2. Script variable $<FORMAT>_OUTPUT_DIR (e.g., $PHP_OUTPUT_DIR)
        3. Global fallback $OUTPUT_DIR
    
    .OUTPUTS
        [string] The resolved absolute path to the output directory
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    # Priority 1: Command-line parameter
    if (-not [string]::IsNullOrWhiteSpace($OutputDir)) {
        Write-DebugLog -Message "Output directory resolved from -OutputDir parameter: '$OutputDir'"
        return $OutputDir
    }

    # Priority 2: Format-specific script variable
    $varName = "${ExportFormat}_OUTPUT_DIR"
    $varValue = Get-Variable -Name $varName -ValueOnly -ErrorAction SilentlyContinue
    if (-not [string]::IsNullOrWhiteSpace($varValue)) {
        Write-DebugLog -Message "Output directory resolved from script variable $$varName: '$varValue'"
        return $varValue
    }

    # Priority 3: Global fallback
    Write-DebugLog -Message "Output directory resolved from global fallback OUTPUT_DIR: '$OUTPUT_DIR'"
    return $OUTPUT_DIR
}

function Ensure-OutputDirectory {
    <#
    .SYNOPSIS
        Ensures the resolved output directory exists and is writable.
    #>
    [CmdletBinding()]
    param()

    $targetDir = Resolve-OutputDirectory
    
    if (-not (Test-Path -LiteralPath $targetDir)) {
        Write-Log -Message "Output directory '$targetDir' does not exist. Creating..." -Level 'INFO'
        try {
            New-Item -Path $targetDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Log -Message "Directory '$targetDir' successfully created." -Level 'OK'
        }
        catch {
            Write-Log -Message "Failed to create directory '$targetDir': $($_.Exception.Message)" -Level 'ERROR'
            throw
        }
    }
    else {
        Write-DebugLog -Message "Output directory '$targetDir' already exists."
    }

    # Verify write permissions with a test file
    try {
        $testFile = Join-Path -Path $targetDir -ChildPath ".write_test_$(Get-Random)"
        [System.IO.File]::WriteAllText($testFile, '')
        Remove-Item -LiteralPath $testFile -Force -ErrorAction Stop
        Write-DebugLog -Message "Write permission verified for '$targetDir'"
    }
    catch {
        Write-Log -Message "Write permission denied for directory '$targetDir': $($_.Exception.Message)" -Level 'ERROR'
        throw
    }
}

# =============================================================================
# GLOBAL STATISTICS
# =============================================================================
$script:Stats = [ordered]@{
    DomainsTotal               = 0
    DomainsSucceeded           = 0
    DomainsFailed              = 0
    AccountsExported           = 0
    AliasesExported            = 0
    DistributionListsExported  = 0
    DistributionListRecipients = 0
    DlMembersRowsExported      = 0
    FilesWritten               = 0
    Errors                     = 0
    StartTime                  = Get-Date
    EndTime                    = $null
}

# =============================================================================
# LOGGING
# =============================================================================

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'DEBUG', 'SUMMARY')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $prefix = switch ($Level) {
        'INFO'    { '[INFO]' }
        'OK'      { '[ OK ]' }
        'WARN'    { '[WARN]' }
        'ERROR'   { '[ERR ]' }
        'DEBUG'   { '[DBG ]' }
        'SUMMARY' { '[SUM ]' }
    }

    $color = switch ($Level) {
        'OK'      { 'Green' }
        'WARN'    { 'Yellow' }
        'ERROR'   { 'Red' }
        'DEBUG'   { 'DarkGray' }
        'SUMMARY' { 'Cyan' }
        default   { 'White' }
    }

    $line = "$timestamp $prefix $Message"

    if (-not ($Level -eq 'DEBUG' -and $DEBUG -ne 1)) {
        Write-Host $line -ForegroundColor $color
    }

    if ($Mode -eq 'B') {
        try {
            Add-Content -Path $global:LogFilePath -Value $line
        }
        catch {
            # Intentionally ignored to avoid recursive logging failures
        }
    }
}

function Write-DebugLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Log -Message $Message -Level 'DEBUG'
}

# =============================================================================
# SECURITY / INPUT
# =============================================================================

function Convert-SecureStringToPlainText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Security.SecureString]$SecureString
    )

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Get-HMailPassword {
    [CmdletBinding()]
    param()

    if ([string]::IsNullOrWhiteSpace($HMAIL_PASSWORD)) {
        Write-Log -Message 'HMAIL_PASSWORD constant is empty. Prompting for hMailServer Administrator password.' -Level 'INFO'
        $securePassword = Read-Host 'Enter hMailServer Administrator password' -AsSecureString
        return (Convert-SecureStringToPlainText -SecureString $securePassword)
    }

    Write-DebugLog -Message 'Using password from HMAIL_PASSWORD constant (insecure; remove it after validation).'
    return $HMAIL_PASSWORD
}

# =============================================================================
# FILESYSTEM HELPERS
# =============================================================================

function Get-SafeFileNamePart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $invalidChars = [IO.Path]::GetInvalidFileNameChars()
    $safe = $Text

    foreach ($char in $invalidChars) {
        $safe = $safe.Replace($char, '_')
    }

    $safe = $safe -replace '[\\/:*?"<>|]', '_'
    return $safe.Trim()
}

function Convert-ToObjectArray {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$InputObject
    )

    if ($null -eq $InputObject) {
        return [object[]]@()
    }

    if ($InputObject -is [string]) {
        return [object[]]@($InputObject)
    }

    if ($InputObject -is [System.Array]) {
        return [object[]]$InputObject
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
        return [object[]]@($InputObject)
    }

    return [object[]]@($InputObject)
}

# =============================================================================
# FORMAT-SPECIFIC ESCAPING & SERIALIZATION HELPERS
# =============================================================================

function Escape-ForMarkdown {
    [CmdletBinding()]
    param([Parameter(Mandatory=$false)][AllowNull()][AllowEmptyString()][string]$Text)
    
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $Text -replace '\\', '\\' -replace '\|', '\\|' -replace '\*', '\*' -replace '_', '\_' `
          -replace '\[', '\[' -replace '\]', '\]' -replace '\(', '\(' -replace '\)', '\)' `
          -replace '#', '\#' -replace '\+', '\+' -replace '\-', '\-' -replace '\.', '\.' `
          -replace '`', '\`' -replace '<', '&lt;' -replace '>', '&gt;'
}

function Escape-ForJS {
    [CmdletBinding()]
    param([Parameter(Mandatory=$false)][AllowNull()][AllowEmptyString()][string]$Text)
    
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $Text -replace '\\', '\\' -replace '"', '\"' -replace "'", "\'" `
          -replace "`n", '\n' -replace "`r", '\r' -replace "`t", '\t' `
          -replace '</', '<\/' -replace '[\x00-\x1f]', { param($m) "\u{0:X4}" -f [int]$m.Value[0] }
}

function Escape-ForPHP {
    [CmdletBinding()]
    param([Parameter(Mandatory=$false)][AllowNull()][AllowEmptyString()][string]$Text)
    
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $Text -replace '\\', '\\\\' -replace "'", "\\'"
}

function Escape-ForXML {
    [CmdletBinding()]
    param([Parameter(Mandatory=$false)][AllowNull()][AllowEmptyString()][string]$Text)
    
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' `
          -replace '"', '&quot;' -replace "'", '&apos;'
}

function ConvertTo-NestedValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        [object]$Value,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet('JSON','JS','PHP','XML')]
        [string]$Format,
        
        [Parameter(Mandatory=$false)]
        [int]$Depth = 0,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxDepth = 10
    )
    
    if ($null -eq $Value -or $Value -is [DBNull]) {
        return $(switch ($Format) {
            'JSON' { 'null' }
            'JS'   { 'null' }
            'PHP'  { 'null' }
            'XML'  { '' }
            default { '' }
        })
    }
    
    if ($Depth -gt $MaxDepth) { return '"[MAX_DEPTH_EXCEEDED]"' }
    
    if ($Value -is [bool]) {
        return $(switch ($Format) {
            'JSON' { if ($Value) { 'true' } else { 'false' } }
            'JS'   { if ($Value) { 'true' } else { 'false' } }
            'PHP'  { if ($Value) { 'true' } else { 'false' } }
            'XML'  { if ($Value) { '1' } else { '0' } }
        })
    }
    
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        return $Value.ToString()
    }
    
    if ($Value -is [DateTime]) {
        $iso = $Value.ToString('o')
        return $(switch ($Format) {
            'JSON' { "`"$iso`"" }
            'JS'   { "`"$iso`"" }
            'PHP'  { "`"$iso`"" }
            'XML'  { (Escape-ForXML -Text $iso) }
        })
    }
    
    if ($Value -is [string]) {
        return $(switch ($Format) {
            'JSON' { 
                $escaped = $Value -replace '\\','\\' -replace '"','\"' -replace "`n",'\n' -replace "`r",'\r' -replace "`t",'\t'
                "`"$escaped`""
            }
            'JS'   { "`"$(Escape-ForJS -Text $Value)`"" }
            'PHP'  { "'$(Escape-ForPHP -Text $Value)'" }
            'XML'  { (Escape-ForXML -Text $Value) }
        })
    }
    
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = @()
        foreach ($item in $Value) {
            $items += (ConvertTo-NestedValue -Value $item -Format $Format -Depth ($Depth + 1) -MaxDepth $MaxDepth)
        }
        return $(switch ($Format) {
            'JSON' { "[$($items -join ',')]" }
            'JS'   { "[$($items -join ',')]" }
            'PHP'  { "array($($items -join ','))" }
            'XML'  { "<items>$($items -join '')</items>" }
        })
    }
    
    if ($Value -is [PSCustomObject] -or $Value -is [hashtable]) {
        $props = @()
        $hash = if ($Value -is [hashtable]) { 
            $Value 
        } else { 
            $h = @{}
            $Value.PSObject.Properties | Where-Object { $_.Name } | ForEach-Object { 
                $h[$_.Name] = $_.Value 
            }
            $h
        }
        
        foreach ($key in $hash.Keys) {
            if ($null -eq $key) { continue }
            $val = ConvertTo-NestedValue -Value $hash[$key] -Format $Format -Depth ($Depth + 1) -MaxDepth $MaxDepth
            $safeKey = $key -replace '[^a-zA-Z0-9_]', '_'
            $props += $(switch ($Format) {
                'JSON' { "`"$safeKey`":$val" }
                'JS'   { "$safeKey`: $val" }
                'PHP'  { "'$(Escape-ForPHP -Text $key)' => $val" }
                'XML'  { "<$safeKey>$val</$safeKey>" }
            })
        }
        return $(switch ($Format) {
            'JSON' { "{$($props -join ',')}" }
            'JS'   { "{ $($props -join ', ') }" }
            'PHP'  { "array($($props -join ','))" }
            'XML'  { "<object>$($props -join '')</object>" }
        })
    }
    
    $str = if ($null -eq $Value) { '' } else { [string]$Value }
    return $(switch ($Format) {
        'JSON' { 
            $escaped = $str -replace '\\','\\' -replace '"','\"'
            "`"$escaped`""
        }
        'JS'   { "`"$(Escape-ForJS -Text $str)`"" }
        'PHP'  { "'$(Escape-ForPHP -Text $str)'" }
        'XML'  { (Escape-ForXML -Text $str) }
    })
}

# =============================================================================
# UNIVERSAL EXPORT ENGINE
# =============================================================================

function Export-DataGeneric {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Data,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet('CSV','JSON','Markdown','JS','PHP','XML')]
        [string]$Format,

        [Parameter(Mandatory = $false)]
        [switch]$AllowEmpty = $true
    )

    try {
        [object[]]$dataToExport = Convert-ToObjectArray -InputObject $Data

        if ($null -eq $dataToExport -or $dataToExport.Count -eq 0) {
            if ($AllowEmpty) {
                Write-DebugLog -Message "No data for '$Path'. Creating empty $Format file."
                $null = Export-EmptyFile -Path $Path -Format $Format
                $script:Stats.FilesWritten++
                Write-Log -Message "$Format file successfully written (empty): $Path" -Level 'OK'
                return $true
            }
            else {
                Write-Log -Message "Skipping $Format export for '$Path' (no data and AllowEmpty = false)." -Level 'INFO'
                return $false
            }
        }

        $content = switch ($Format) {
            'CSV'      { ConvertTo-CsvContent -Data $dataToExport }
            'JSON'     { ConvertTo-JsonContent -Data $dataToExport }
            'Markdown' { ConvertTo-MarkdownContent -Data $dataToExport }
            'JS'       { ConvertTo-JSContent -Data $dataToExport -Path $Path }
            'PHP'      { ConvertTo-PHPContent -Data $dataToExport -Path $Path }
            'XML'      { ConvertTo-XMLContent -Data $dataToExport }
        }

        $config = $FormatConfig[$ExportFormat]
        $bomEnabled = $config.BOM ?? $false  # Null-coalescing: defaults to $false if key missing
        $encoding = Get-FileEncoding -WithBom $bomEnabled

        [System.IO.File]::WriteAllText($Path, $content, $encoding)
        
        $script:Stats.FilesWritten++
        Write-Log -Message "$Format file successfully written: $Path" -Level 'OK'
        return $true
    }
    catch {
        $script:Stats.Errors++
        Write-Log -Message "Failed to write $Format file '$Path': $($_.Exception.Message)" -Level 'ERROR'
        Write-DebugLog -Message "$Format export exception details: $($_ | Out-String)"
        return $false
    }
}

function Export-EmptyFile {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][ValidateSet('CSV','JSON','Markdown','JS','PHP','XML')][string]$Format
    )
    
    $content = $(switch ($Format) {
        'CSV'      { "`r`n" }
        'JSON'     { "[]" }
        'Markdown' { "" }
        'JS'       { "const data = [];`r`nexport default data;`r`n" }
        'PHP'      {
@'
<?php
/**
 * Auto-generated empty export file
 * Returns empty array for safe inclusion
 */
return [];
'@
        }
        'XML'      { '<?xml version="1.0" encoding="UTF-8"?><root/>' }
    })
    
    $config = $FormatConfig[$Format]
    $bomEnabled = $config.BOM ?? $false
    $encoding = Get-FileEncoding -WithBom $bomEnabled
    
    [System.IO.File]::WriteAllText($Path, $content, $encoding)
}

function ConvertTo-CsvContent {
    param([Parameter(Mandatory=$true)][object[]]$Data)
    
    if ($Data.Count -eq 0) { return "`r`n" }
    
    $flattened = $Data | ForEach-Object {
        $obj = [ordered]@{}
        $properties = $_.PSObject.Properties | Where-Object { $_.Name -and $_.Name -ne '' }
        foreach ($prop in $properties) {
            $val = $prop.Value
            if ($null -eq $val -or $val -is [DBNull]) {
                $obj[$prop.Name] = ''
            } elseif ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                $obj[$prop.Name] = ($val | Where-Object { $null -ne $_ } | Join-String '; ')
            } elseif ($val -is [PSCustomObject] -or $val -is [hashtable]) {
                $obj[$prop.Name] = ($val | ConvertTo-Json -Compress -Depth 2 -ErrorAction SilentlyContinue) ?? ''
            } else {
                $obj[$prop.Name] = $val
            }
        }
        [PSCustomObject]$obj
    }
    
    $flattened | ConvertTo-Csv -NoTypeInformation -Delimiter ','
}

function ConvertTo-JsonContent {
    param([Parameter(Mandatory=$true)][object[]]$Data)
    
    $config = $FormatConfig['JSON']
    $Data | ConvertTo-Json -Depth $config.Depth -Compress:$false
}

function ConvertTo-MarkdownContent {
    param([Parameter(Mandatory=$true)][object[]]$Data)
    
    if ($Data.Count -eq 0) { return "" }
    
    $headers = $Data[0].PSObject.Properties | Where-Object { $_.Name -and $_.Name -ne '' } | ForEach-Object { $_.Name }
    
    if ($headers.Count -eq 0) { return "| _no_data_ |`r`n| --- |" }
    
    $headerRow = "| " + ($headers | ForEach-Object { Escape-ForMarkdown -Text $_ } | Join-String ' | ') + " |"
    $separator = "| " + (':' * $headers.Count | Join-String ': | ') + " |"
    
    $rows = $Data | ForEach-Object {
        $row = $_
        "| " + ($headers | ForEach-Object { 
            $val = $row.$_
            $str = if ($null -eq $val -or $val -is [DBNull]) { 
                '' 
            } elseif ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) { 
                $val -join '; ' 
            } else { 
                [string]$val 
            }
            Escape-ForMarkdown -Text $str 
        } | Join-String ' | ') + " |"
    }
    
    @($headerRow, $separator) + $rows | Join-String "`r`n"
}

function ConvertTo-JSContent {
    param(
        [Parameter(Mandatory=$true)][object[]]$Data,
        [Parameter(Mandatory=$true)][string]$Path
    )
    
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $varName = $fileName -replace '[^a-zA-Z0-9_]', '_'
    
    $items = $Data | ForEach-Object {
        $props = @()
        $properties = $_.PSObject.Properties | Where-Object { $_.Name -and $_.Name -ne '' }
        foreach ($prop in $properties) {
            $propName = $prop.Name
            $propValue = $prop.Value
            $serializedValue = ConvertTo-NestedValue -Value $propValue -Format 'JS' -Depth 1
            $props += "$propName`: $serializedValue"
        }
        "{ $($props -join ', ') }"
    }
    
    $jsTemplate = @'
/**
 * Auto-generated hMailServer export
 * Format: JavaScript Module
 * File: {FILE}
 * Generated: {DATE}
 */
const {VAR} = [
{ITEMS}
];
export default {VAR};
'@
    
    $content = $jsTemplate -replace '\{FILE\}', $fileName
    $content = $content -replace '\{DATE\}', (Get-Date -Format 'o')
    $content = $content -replace '\{VAR\}', $varName
    $content = $content -replace '\{ITEMS\}', ($items -join ",`r`n")
    
    return $content
}

function ConvertTo-PHPContent {
    param(
        [Parameter(Mandatory=$true)][object[]]$Data,
        [Parameter(Mandatory=$true)][string]$Path
    )
    
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $constName = ('DATA_' + ($fileName -replace '[^a-zA-Z0-9_]', '_').ToUpper())
    
    $items = $Data | ForEach-Object {
        $props = @()
        $properties = $_.PSObject.Properties | Where-Object { $_.Name -and $_.Name -ne '' }
        foreach ($prop in $properties) {
            $propName = $prop.Name
            $propValue = $prop.Value
            $escapedName = Escape-ForPHP -Text $propName
            $serializedValue = ConvertTo-NestedValue -Value $propValue -Format 'PHP' -Depth 1
            $props += "'$escapedName' => $serializedValue"
        }
        "array($($props -join ','))"
    }
    
    $phpHeader = @'
<?php
/**
 * Auto-generated hMailServer export
 * Format: PHP Constants
 * File: {FILE}
 * Generated: {DATE}
 * 
 * Usage: require_once '{FILE}'; // Returns data array
 */

const {CONST} = [
'@
    
    $phpFooter = @'
];

return {CONST};


1;
'@
    
    $content = $phpHeader -replace '\{FILE\}', $fileName
    $content = $content -replace '\{DATE\}', (Get-Date -Format 'o')
    $content = $content -replace '\{CONST\}', $constName
    $content += "`r`n$($items -join ",`r`n")"
    $footer = $phpFooter -replace '\{CONST\}', $constName
    $content += "`r`n$footer"
    
    return $content
}

function ConvertTo-XMLContent {
    param([Parameter(Mandatory=$true)][object[]]$Data)
    
    if ($Data.Count -eq 0) { 
        return '<?xml version="1.0" encoding="UTF-8"?><root/>' 
    }
    
    $rootName = 'export'
    $itemName = 'item'
    
    $itemsXml = $Data | ForEach-Object {
        $props = @()
        $properties = $_.PSObject.Properties | Where-Object { $_.Name -and $_.Name -ne '' }
        foreach ($prop in $properties) {
            $propName = $prop.Name
            $propValue = $prop.Value
            $val = ConvertTo-NestedValue -Value $propValue -Format 'XML' -Depth 1
            $safeName = $propName -replace '[^a-zA-Z0-9_\-.]', '_'
            if ($safeName -notmatch '^[a-zA-Z_]') { 
                $safeName = "field_$safeName"  # Ensure XML tag starts with letter/underscore
            }
            $props += "<$safeName>$val</$safeName>"
        }
        "<$itemName>$($props -join '')</$itemName>"
    }
    
    $xmlTemplate = @'
<?xml version="1.0" encoding="UTF-8"?>
<{ROOT}>
{ITEMS}
</{ROOT}>
'@
    
    $content = $xmlTemplate -replace '\{ROOT\}', $rootName
    $content = $content -replace '\{ITEMS\}', ($itemsXml -join "`r`n")
    
    return $content
}

# =============================================================================
# BACKWARD-COMPATIBLE CSV EXPORT WRAPPER
# =============================================================================

function Export-CsvSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Data,
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [switch]$AllowEmpty = $true
    )
    return Export-DataGeneric -Data $Data -Path $Path -Format 'CSV' -AllowEmpty:$AllowEmpty
}

# =============================================================================
# HMAILSERVER CONNECTION
# =============================================================================

function Connect-HMailServer {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Password)

    try {
        Write-Log -Message 'Creating COM object: hMailServer.Application...' -Level 'INFO'
        $application = New-Object -ComObject 'hMailServer.Application' -ErrorAction Stop
        Write-Log -Message 'COM object hMailServer.Application created successfully.' -Level 'OK'

        Write-Log -Message "Authenticating as '$HMAIL_ADMIN_USER'..." -Level 'INFO'
        $null = $application.Authenticate($HMAIL_ADMIN_USER, $Password)
        Write-Log -Message 'Authentication successful.' -Level 'OK'

        return $application
    }
    catch {
        $script:Stats.Errors++
        Write-Log -Message "Failed to connect to hMailServer: $($_.Exception.Message)" -Level 'ERROR'
        Write-DebugLog -Message "Connection exception stack: $($_.ScriptStackTrace)"
        throw
    }
}

# =============================================================================
# SAFE PROPERTY ACCESS
# =============================================================================

function Get-PropertySafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Object,
        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    try {
        return $Object.$PropertyName
    }
    catch {
        $knownMissing = @('PasswordExpirationEnabled','PasswordExpires','Created','LastLogon')
        if ($PropertyName -notin $knownMissing) {
            Write-DebugLog -Message "Get-PropertySafe: Failed to read '$PropertyName' on '$($Object.GetType().FullName)': $($_.Exception.Message)"
        }
        return $null
    }
}

# =============================================================================
# DATA EXPORT HELPERS
# =============================================================================

function Get-DomainAccounts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Domain,
        [Parameter(Mandatory = $true)] [string]$DomainName
    )

    $accounts = @()
    try {
        $accountCount = [int]$Domain.Accounts.Count
        Write-DebugLog -Message "Domain '${DomainName}': Found ${accountCount} account(s)"

        for ($i = 0; $i -lt $accountCount; $i++) {
            $account = $Domain.Accounts.Item($i)
            Write-DebugLog -Message "  Processing account #${i}: Address='$(Get-PropertySafe -Object $account -PropertyName 'Address')'"

            $accounts += [PSCustomObject]@{
                Domain                    = $DomainName
                Address                   = Get-PropertySafe -Object $account -PropertyName 'Address'
                Active                    = Get-PropertySafe -Object $account -PropertyName 'Active'
                IsAD                      = Get-PropertySafe -Object $account -PropertyName 'IsAD'
                MaxSizeMB                 = Get-PropertySafe -Object $account -PropertyName 'MaxSize'
                PasswordExpirationEnabled = Get-PropertySafe -Object $account -PropertyName 'PasswordExpirationEnabled'
                PasswordExpires           = Get-PropertySafe -Object $account -PropertyName 'PasswordExpires'
                VacationMessageIsOn       = Get-PropertySafe -Object $account -PropertyName 'VacationMessageIsOn'
                ForwardEnabled            = Get-PropertySafe -Object $account -PropertyName 'ForwardEnabled'
                ForwardAddress            = Get-PropertySafe -Object $account -PropertyName 'ForwardAddress'
                SignatureEnabled          = Get-PropertySafe -Object $account -PropertyName 'SignatureEnabled'
                AdminLevel                = Get-PropertySafe -Object $account -PropertyName 'AdminLevel'
                PersonFirstName           = Get-PropertySafe -Object $account -PropertyName 'PersonFirstName'
                PersonLastName            = Get-PropertySafe -Object $account -PropertyName 'PersonLastName'
                Created                   = Get-PropertySafe -Object $account -PropertyName 'Created'
                LastLogon                 = Get-PropertySafe -Object $account -PropertyName 'LastLogon'
            }
        }
        $script:Stats.AccountsExported += $accounts.Count
        Write-Log -Message "Domain '${DomainName}': Prepared $($accounts.Count) account(s) for export" -Level 'OK'
        return $accounts
    }
    catch {
        $script:Stats.Errors++
        Write-Log -Message "Failed to read accounts for domain '${DomainName}': $($_.Exception.Message)" -Level 'ERROR'
        Write-DebugLog -Message "Account export exception: $($_ | Out-String)"
        throw
    }
}

function Get-DomainAliases {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Domain,
        [Parameter(Mandatory = $true)] [string]$DomainName
    )

    $aliases = @()
    try {
        $aliasCount = [int]$Domain.Aliases.Count
        Write-DebugLog -Message "Domain '${DomainName}': Found ${aliasCount} alias(es)"

        for ($i = 0; $i -lt $aliasCount; $i++) {
            $alias = $Domain.Aliases.Item($i)
            $name    = Get-PropertySafe -Object $alias -PropertyName 'Name'
            $value   = Get-PropertySafe -Object $alias -PropertyName 'Value'
            $active  = Get-PropertySafe -Object $alias -PropertyName 'Active'
            $created = Get-PropertySafe -Object $alias -PropertyName 'Created'

            $fullAddress = if ($name -and $name -match '@') { $name }
            elseif ($name) { "${name}@${DomainName}" }
            else { $null }

            Write-DebugLog -Message "  Processing alias #${i}: Name='${name}', Value='${value}', Active='${active}'"

            $aliases += [PSCustomObject]@{
                Domain  = $DomainName
                Name    = $name
                Address = $fullAddress
                Value   = $value
                Active  = $active
                Created = $created
            }
        }
        $script:Stats.AliasesExported += $aliases.Count
        Write-Log -Message "Domain '${DomainName}': Prepared $($aliases.Count) alias(es) for export" -Level 'OK'
        return $aliases
    }
    catch {
        $script:Stats.Errors++
        Write-Log -Message "Failed to read aliases for domain '${DomainName}': $($_.Exception.Message)" -Level 'ERROR'
        Write-DebugLog -Message "Alias export exception: $($_ | Out-String)"
        throw
    }
}

function Get-DomainDistributionLists {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] $Domain,
        [Parameter(Mandatory = $true)] [string]$DomainName
    )

    $distributionLists = @()
    $distributionListMembers = @()

    try {
        $listCount = [int]$Domain.DistributionLists.Count
        Write-DebugLog -Message "Domain '${DomainName}': Found ${listCount} distribution list(s)"

        for ($i = 0; $i -lt $listCount; $i++) {
            $list = $Domain.DistributionLists.Item($i)
            $address     = Get-PropertySafe -Object $list -PropertyName 'Address'
            $active      = Get-PropertySafe -Object $list -PropertyName 'Active'
            $requireAuth = Get-PropertySafe -Object $list -PropertyName 'RequireAddressAuthentication'
            $mode        = Get-PropertySafe -Object $list -PropertyName 'Mode'
            $created     = Get-PropertySafe -Object $list -PropertyName 'Created'

            $recipients = @()
            $recipientCountReported = 0
            try { $recipientCountReported = [int]$list.Recipients.Count }
            catch {
                Write-DebugLog -Message "  Distribution list '${address}': failed to read Recipients.Count: $($_.Exception.Message)"
                $recipientCountReported = 0
            }

            Write-DebugLog -Message "  Processing distribution list #${i}: Address='${address}', Recipients.Count=${recipientCountReported}"

            for ($r = 0; $r -lt $recipientCountReported; $r++) {
                try {
                    $recipient = $list.Recipients.Item($r)
                    $rcptAddr = Get-PropertySafe -Object $recipient -PropertyName 'RecipientAddress'
                    if (-not [string]::IsNullOrWhiteSpace($rcptAddr)) {
                        $recipients += $rcptAddr
                        $distributionListMembers += [PSCustomObject]@{
                            Domain           = $DomainName
                            DistributionList = $address
                            RecipientAddress = $rcptAddr
                        }
                        Write-DebugLog -Message "    Recipient #${r}: '${rcptAddr}'"
                    }
                    else {
                        Write-DebugLog -Message "    Recipient #${r}: empty or unreadable RecipientAddress"
                    }
                }
                catch {
                    $script:Stats.Errors++
                    Write-DebugLog -Message "    Failed to read recipient #${r} for list '${address}': $($_.Exception.Message)"
                }
            }

            $actualRecipientCount = $recipients.Count
            $script:Stats.DistributionListRecipients += $actualRecipientCount

            $distributionLists += [PSCustomObject]@{
                Domain         = $DomainName
                Address        = $address
                Active         = $active
                RequireAuth    = $requireAuth
                Mode           = $mode
                RecipientCount = $actualRecipientCount
                Recipients     = ($recipients -join ';')
                Created        = $created
            }
        }

        $script:Stats.DistributionListsExported += $distributionLists.Count
        $script:Stats.DlMembersRowsExported += $distributionListMembers.Count
        Write-Log -Message "Domain '${DomainName}': Prepared $($distributionLists.Count) distribution list(s) for export" -Level 'OK'
        Write-DebugLog -Message "Domain '${DomainName}': Prepared $($distributionListMembers.Count) distribution list member row(s)"

        return [PSCustomObject]@{ Lists = $distributionLists; Members = $distributionListMembers }
    }
    catch {
        $script:Stats.Errors++
        Write-Log -Message "Failed to read distribution lists for domain '${DomainName}': $($_.Exception.Message)" -Level 'ERROR'
        Write-DebugLog -Message "Distribution list export exception: $($_ | Out-String)"
        throw
    }
}

# =============================================================================
# MAIN EXECUTION
# =============================================================================

$hMailApp = $null

try {
    Write-Log -Message '=== hMailServer Data Export Started ===' -Level 'INFO'
    Write-DebugLog -Message "Configuration: DEBUG=${DEBUG}, Mode='${Mode}', ExportFormat='${ExportFormat}'"
    Write-DebugLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
    
    # Resolve and ensure output directory BEFORE any file operations
    $activeOutputDir = Resolve-OutputDirectory
    Write-Log -Message "Active output directory for ${ExportFormat}: ${activeOutputDir}" -Level 'INFO'
    Ensure-OutputDirectory
    
    # Initialize log file path in the resolved directory
    $global:LogFilePath = Join-Path -Path $activeOutputDir -ChildPath ("hMailServerExport_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

    if ($Mode -eq 'B') {
        Write-Log -Message "Advanced mode enabled (Mode='B'). Log file: ${global:LogFilePath}" -Level 'INFO'
    }
    
    Write-Log -Message "Export format: ${ExportFormat} (extension: $($FormatConfig[$ExportFormat].Extension))" -Level 'INFO'

    $password = Get-HMailPassword
    $hMailApp = Connect-HMailServer -Password $password

    Write-Log -Message 'Retrieving domains list...' -Level 'INFO'
    $domains = $hMailApp.Domains
    $domainCount = [int]$domains.Count
    $script:Stats.DomainsTotal = $domainCount
    Write-Log -Message "Found ${domainCount} domain(s) in hMailServer" -Level 'OK'

    for ($d = 0; $d -lt $domainCount; $d++) {
        try {
            $domain = $domains.Item($d)
            $domainName = [string](Get-PropertySafe -Object $domain -PropertyName 'Name')

            if ([string]::IsNullOrWhiteSpace($domainName)) {
                $domainName = "domain_${d}"
            }

            $safeDomainName = Get-SafeFileNamePart -Text $domainName
            $ext = $FormatConfig[$ExportFormat].Extension

            Write-Log -Message "Processing domain '${domainName}' ($($d + 1) of ${domainCount})..." -Level 'INFO'
            Write-DebugLog -Message "Domain properties: Name='${domainName}', Active='$(Get-PropertySafe -Object $domain -PropertyName 'Active')', MaxSize='$(Get-PropertySafe -Object $domain -PropertyName 'MaxSize')'"

            $accounts = Get-DomainAccounts -Domain $domain -DomainName $domainName
            $aliases  = Get-DomainAliases  -Domain $domain -DomainName $domainName
            $distributionData = Get-DomainDistributionLists -Domain $domain -DomainName $domainName
            $dlists  = $distributionData.Lists
            $members = $distributionData.Members

            # Build file paths using the resolved output directory
            $accountsFile = Join-Path -Path $activeOutputDir -ChildPath "${safeDomainName}_accounts${ext}"
            $aliasesFile  = Join-Path -Path $activeOutputDir -ChildPath "${safeDomainName}_aliases${ext}"
            $dlistsFile   = Join-Path -Path $activeOutputDir -ChildPath "${safeDomainName}_distribution-lists${ext}"

            $null = Export-DataGeneric -Data $accounts -Path $accountsFile -Format $ExportFormat
            $null = Export-DataGeneric -Data $aliases  -Path $aliasesFile  -Format $ExportFormat
            $null = Export-DataGeneric -Data $dlists   -Path $dlistsFile   -Format $ExportFormat

            if ($Mode -eq 'B') {
                $dlMembersFile = Join-Path -Path $activeOutputDir -ChildPath "${safeDomainName}_distribution-list-members${ext}"
                $null = Export-DataGeneric -Data $members -Path $dlMembersFile -Format $ExportFormat -AllowEmpty:$false
            }

            $script:Stats.DomainsSucceeded++
            Write-Log -Message "Domain '${domainName}' processed successfully" -Level 'OK'
        }
        catch {
            $script:Stats.DomainsFailed++
            $script:Stats.Errors++
            Write-Log -Message "Failed to process domain at index ${d}: $($_.Exception.Message)" -Level 'ERROR'
            Write-DebugLog -Message "Domain processing exception: $($_ | Out-String)"
        }
    }
}
catch {
    $script:Stats.Errors++
    Write-Log -Message "Critical execution error: $($_.Exception.Message)" -Level 'ERROR'
    Write-DebugLog -Message "Critical exception details: $($_ | Out-String)"
}
finally {
    if ($null -ne $hMailApp) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hMailApp) | Out-Null
            Write-DebugLog -Message 'COM object released successfully'
        }
        catch {
            Write-DebugLog -Message "Warning: Failed to release COM object: $($_.Exception.Message)"
        }
    }

    $script:Stats.EndTime = Get-Date
    $duration = New-TimeSpan -Start $script:Stats.StartTime -End $script:Stats.EndTime

    Write-Host ''
    Write-Log -Message '=== Export Summary ===' -Level 'SUMMARY'
    Write-Log -Message "Start time                 : $($script:Stats.StartTime)" -Level 'SUMMARY'
    Write-Log -Message "End time                   : $($script:Stats.EndTime)" -Level 'SUMMARY'
    Write-Log -Message "Duration                   : $($duration.ToString('hh\:mm\:ss'))" -Level 'SUMMARY'
    Write-Log -Message "Domains total              : $($script:Stats.DomainsTotal)" -Level 'SUMMARY'
    Write-Log -Message "Domains succeeded          : $($script:Stats.DomainsSucceeded)" -Level 'SUMMARY'
    Write-Log -Message "Domains failed             : $($script:Stats.DomainsFailed)" -Level 'SUMMARY'
    Write-Log -Message "Accounts exported          : $($script:Stats.AccountsExported)" -Level 'SUMMARY'
    Write-Log -Message "Aliases exported           : $($script:Stats.AliasesExported)" -Level 'SUMMARY'
    Write-Log -Message "Distribution lists exported: $($script:Stats.DistributionListsExported)" -Level 'SUMMARY'
    Write-Log -Message "DL recipients exported     : $($script:Stats.DistributionListRecipients)" -Level 'SUMMARY'
    Write-Log -Message "DL member rows exported    : $($script:Stats.DlMembersRowsExported)" -Level 'SUMMARY'
    Write-Log -Message "${ExportFormat} files written          : $($script:Stats.FilesWritten)" -Level 'SUMMARY'
    Write-Log -Message "Total errors               : $($script:Stats.Errors)" -Level 'SUMMARY'
    Write-Log -Message "Active output directory    : $(Resolve-OutputDirectory)" -Level 'SUMMARY'

    if ($Mode -eq 'B') {
        Write-Log -Message "Log file                   : ${global:LogFilePath}" -Level 'SUMMARY'
    }

    Write-Log -Message '=== Export Completed ===' -Level 'SUMMARY'

    if ($script:Stats.Errors -gt 0) {
        Write-Warning "Script completed with $($script:Stats.Errors) error(s). Check logs above."
        exit 1
    }
    else {
        Write-Log -Message 'All operations completed successfully.' -Level 'OK'
        exit 0
    }
}
