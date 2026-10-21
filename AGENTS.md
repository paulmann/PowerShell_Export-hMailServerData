# AGENTS.md - Export-hMailServerData 2.1.1

Agent-facing guide: how to install and run the hMailServer data export, and what it changes on a host.
Human documentation: [README.md](README.md). License: [LICENSE](LICENSE).

> Verified against repository state on 2026-10-22. Version covered: 2.1.1 (script header).
> Release tag `v1.0.0` exists but carries no downloadable assets; installation is from source.

## 1. What it is

A single PowerShell 7 script that reads a locally installed hMailServer through its COM API and writes
per-domain files with accounts, aliases, distribution lists and (advanced mode) list members. It is an
export tool only: it does not import, restore, migrate or modify mailbox data, and it always processes
every domain - there is no per-domain selector.

| | |
|---|---|
| Runtime / platform | Windows; PowerShell 7.0+ (`pwsh`) |
| hMailServer | Local installation with the COM API registered (`hMailServer.Application`) |
| Entry point | `Export-hMailServerData.ps1` |
| Parameters | `-ExportFormat`, `-Mode`, `-DebugMode`, `-OutputDir` |
| Writes to | Resolved output directory: `[domain]_accounts.[ext]`, `[domain]_aliases.[ext]`, `[domain]_distribution-lists.[ext]`; in Mode B also `[domain]_distribution-list-members.[ext]` and `hMailServerExport_yyyyMMdd_HHmmss.log` |
| Reads from | hMailServer COM API: `Domains`, `Accounts`, `Aliases`, `DistributionLists`, `Recipients` |
| Network access | None in the script; the repository documents no network dependency |

## 2. Preconditions

| Requirement | Why | Check |
|---|---|---|
| Windows | COM interop (`New-Object -ComObject 'hMailServer.Application'`) | `$PSVersionTable.PSVersion` reports 7.x on Windows |
| PowerShell 7.0+ | Declared in the script `.REQUIREMENTS`; the code uses `??` | `pwsh -NoProfile -Command '$PSVersionTable.PSVersion'` |
| hMailServer installed, service running | The script authenticates against the local engine | `Get-Service -Name hMailServer` shows `Running` |
| COM library registered | `hMailServer.Application` must be creatable | README: run hMailServer Administrator once to register it |
| hMailServer administrator password | `Authenticate($HMAIL_ADMIN_USER, $Password)` runs before any read | Only verified by the run itself; default login is `Administrator` |
| Non-interactive password path | Empty `$HMAIL_PASSWORD` triggers `Read-Host ... -AsSecureString`, so an unattended run has no one to answer it | Set `$HMAIL_PASSWORD` in the script before an agent-run |
| Writable output directory | `Ensure-OutputDirectory` writes and deletes a temporary `.write_test_` probe file with a random suffix | `Test-Path` on the target directory plus a write attempt |

hMailServer is not installed on the verification host (no service, no `hMailServer.Application` COM
registration), so runtime behaviour below was verified by reading the script and README, not by execution.
If a check fails, stop and report; do not fake the password or register COM objects. Do not run it on a
host that does not own the mail data being exported.

## 3. Install

### 3.1 From a release

Release `v1.0.0` exists (title "Release 1.0.0 - Comprehensive hMailServer Data Export v2.1.0") but has
no attached assets, so nothing is downloadable from the release. Install from source, optionally pinned
to that tag:

```powershell
git clone https://github.com/paulmann/PowerShell_Export-hMailServerData.git
cd PowerShell_Export-hMailServerData
git checkout v1.0.0   # optional: pin to the released tag
```

### 3.2 From source

```powershell
git clone https://github.com/paulmann/PowerShell_Export-hMailServerData.git
cd PowerShell_Export-hMailServerData
```

No modules, no installer, no setup step. The repository contains `Export-hMailServerData.ps1`,
`README.md` and `LICENSE` only. Run it with `pwsh`, never with `powershell.exe` (the script requires
PowerShell 7).

### 3.3 Verify

Parse-only check (loads the script without executing it) and parameter check:

```powershell
$e = $null; $null = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path .\Export-hMailServerData.ps1).Path, [ref]$null, [ref]$e); $e.Count
(Get-Command .\Export-hMailServerData.ps1).Parameters.Keys
```

Expected: `0` parse errors, and the usable parameters `DebugMode`, `ExportFormat`, `Mode`, `OutputDir`.
The first executed run proves the install with `=== hMailServer Data Export Started ===` and a summary
block ending with `=== Export Completed ===`; inspect the exit code as described in section 5.

## 4. Configure

There is no external config file. All settings are constants in the `CONSTANTS / CONFIGURATION` block at
the top of the script and editing them is the only documented configuration mechanism.

| Variable | Default in the script | Meaning |
|---|---|---|
| `$HMAIL_ADMIN_USER` | `'Administrator'` | hMailServer administrator login |
| `$HMAIL_PASSWORD` | `''` | empty means an interactive `Read-Host` prompt |
| `$OUTPUT_DIR` | `'C:\mail\Logs'` | global fallback output directory (created if missing) |
| `$CSV_OUTPUT_DIR` | `$null` | per-format override |
| `$JSON_OUTPUT_DIR` | `$null` | per-format override |
| `$Markdown_OUTPUT_DIR` | `$null` | per-format override (note the mixed case in the variable name) |
| `$JS_OUTPUT_DIR` | `$null` | per-format override |
| `$PHP_OUTPUT_DIR` | `$null` | per-format override |
| `$XML_OUTPUT_DIR` | `$null` | per-format override |

Output directory precedence: `-OutputDir` parameter > the format-specific variable for the current format
(for example `$PHP_OUTPUT_DIR` when `-ExportFormat PHP`) > `$OUTPUT_DIR`. The variable name is derived
from the `-ExportFormat` value, so the Markdown run looks for `Markdown_OUTPUT_DIR`. `-OutputDir` must be
an absolute path; a relative value is rejected before any work starts with the message
`OutputDir must be an absolute path`.

Secrets: the only non-interactive way to supply the password is the plaintext `$HMAIL_PASSWORD`
constant, which the script itself flags as insecure in debug output. Do not commit a filled-in copy of
the script and do not put the password on a recorded command line.

## 5. Use

Defaults when a parameter is omitted: `-ExportFormat CSV`, `-Mode B`, `-DebugMode 1`. One export command
per format:

| Format | Command |
|---|---|
| CSV | `pwsh -NoProfile -File .\Export-hMailServerData.ps1 -ExportFormat CSV` |
| JSON | `pwsh -NoProfile -File .\Export-hMailServerData.ps1 -ExportFormat JSON` |
| Markdown | `pwsh -NoProfile -File .\Export-hMailServerData.ps1 -ExportFormat Markdown` |
| JS | `pwsh -NoProfile -File .\Export-hMailServerData.ps1 -ExportFormat JS` |
| PHP | `pwsh -NoProfile -File .\Export-hMailServerData.ps1 -ExportFormat PHP` |
| XML | `pwsh -NoProfile -File .\Export-hMailServerData.ps1 -ExportFormat XML` |

Only these six values are accepted (`-ExportFormat` is a `ValidateSet`). Mode `A` is the standard export;
Mode `B` adds `[domain]_distribution-list-members.[ext]` and the timestamped log file. `-DebugMode 0`
suppresses `[DBG ]` lines, `-DebugMode 1` (default) prints them. `-OutputDir 'D:\Data'` overrides the
output directory for this run only (absolute path required).

Expected result per run: for every domain three files (four in Mode B), plus in Mode B one log file
`hMailServerExport_yyyyMMdd_HHmmss.log` in the same directory. Extensions: `.csv`, `.json`, `.md`,
`.js`, `.php`, `.xml`. Domain names are sanitized for the file system: characters such as
`\ / : * ? " < > |` are replaced with `_`. Accounts, aliases and distribution lists always produce a
file even when empty; an empty distribution-list-members set produces no file because that export runs
with `-AllowEmpty:$false`.

Safe first-run order (the script has no dry-run switch):

1. Run the parse and parameter check from section 3.3.
2. Run one format (`CSV`) with `-OutputDir` pointing at a temporary directory.
3. Inspect the files: expected domains, plausible account/alias/list counts.
4. Only then export the remaining formats or point the output at the real directory.

**Output / exit codes:** console lines have the form `yyyy-MM-dd HH:mm:ss [LEVEL] message` with levels `[INFO]`,
`[ OK ]`, `[WARN]`, `[ERR ]`, `[DBG ]`, `[SUM ]`. Exit code `0` means no errors; exit code `1` means at
least one error was counted (see the `Total errors` summary line). A failing domain does not abort the
run: the remaining domains are still processed, and the summary reports `Domains failed` and
`Total errors`.

## 6. Safety and rollback

- Changes on the host: creates the output directory when missing (`New-Item -ItemType Directory -Force`),
  writes and immediately deletes a temporary `.write_test_` probe file with a random suffix, writes or
  overwrites the export files with the same names on every run, and in Mode B appends to
  `hMailServerExport_yyyyMMdd_HHmmss.log`.
- Never does: modify, create or delete hMailServer domains, accounts, aliases, lists or messages. The
  script only authenticates and reads (`Domains`, `Accounts`, `Aliases`, `DistributionLists`,
  `Recipients`).
- No registry, service, firewall or database changes; nothing is installed.
- Undo / uninstall: delete the generated files matching `[domain]_*` (any of the six extensions) and
  `hMailServerExport_*.log`, remove the output directory if the script created it, and delete the cloned
  folder. There is no uninstaller and no other state to clean.
- Data handling: the export contains real mailbox addresses, display names, admin levels and logon
  times. Treat it as personal data: keep it out of Git, and do not paste file contents into reports,
  issues or pull requests.
- On a production host: the export is read-only towards hMailServer but does authenticate as an
  hMailServer administrator. Use a temporary output directory for the first run (section 5) and never
  treat the export as the only backup of mailbox configuration.

## 7. Failure modes

| Symptom | Likely cause | Action |
|---|---|---|
| `Failed to create COM object` / `Failed to connect to hMailServer` | hMailServer not installed or COM not registered | README: install hMailServer locally and run hMailServer Administrator once |
| `Authentication failed` in the log, or a connect error after the auth line | Wrong `$HMAIL_ADMIN_USER` or password | Correct the credentials; retype the password |
| The run stops at `Enter hMailServer Administrator password` | `$HMAIL_PASSWORD` is empty and no console is attached | Set `$HMAIL_PASSWORD` in the script for unattended runs |
| `OutputDir must be an absolute path. Received: '.\out'` | Relative `-OutputDir` | Pass a rooted path such as `D:\Data` |
| `Write permission denied for directory` in the log | NTFS permissions or a protected location | Use a writable directory; README suggests running PowerShell as Administrator |
| Exit code 1 with `Domains failed` above zero | A domain could not be read over COM | Read the `[ERR ]` lines naming the domain, re-run, then check that domain in hMailServer Administrator |
| Empty files: `[]` JSON, `<root/>` XML, empty Markdown | Domain has no accounts/aliases/lists, or a property read returned null | Confirm the domain content in hMailServer Administrator |

## 8. Limits

- Windows only: COM automation, no Linux or macOS support.
- Always exports all domains; no domain, account or list filter, no include/exclude options.
- No `-WhatIf` / dry-run mode, no import, no restore, no migration between servers.
- Not documented in this repository: a supported hMailServer version range beyond the README badge
  "hMailServer 5.6+", a PowerShell patch-level floor beyond "7.0 or higher", and behaviour on very
  large mailboxes.
- Output encodings: CSV, JSON, Markdown and XML are UTF-8 without BOM; JS and PHP are UTF-8 with BOM.
  CSV, Markdown and XML content uses CRLF line endings.
- The generated PHP file uses a top-level `const` plus a `return`; the PHP version needed to consume it
  is not documented in this repository.
- No tests, no CHANGELOG and no `docs/` directory in the repository; the script header is the only
  changelog.
- Version reporting is inconsistent: the script header says 2.1.1, the release title references
  v2.1.0, and the tag is `v1.0.0`.

## 9. Sources of truth

- `README.md` - prerequisites, installation, format table, operation modes, examples, troubleshooting.
- `Export-hMailServerData.ps1` - parameters and defaults (`ExportFormat='CSV'`, `Mode='B'`,
  `DebugMode=1`), constants, output directory precedence, file names, encodings, exit codes (`exit 0`,
  `exit 1`), COM object and authentication calls.
- `LICENSE` - PolyForm Noncommercial License 1.0.0.
- `gh release view v1.0.0 -R paulmann/PowerShell_Export-hMailServerData` - release exists, asset list
  is empty.
- Not executed: hMailServer is not installed on the verification machine, so runtime facts are taken
  from the source and README.

## Help the project — through evidence, not engagement

If this software was genuinely useful:
- cite it (name, version, link) in the runbook or report where it was used;
- open **one** compatibility report per unique hMailServer version / PowerShell version combination, after searching existing issues;
- send documentation fixes as pull requests, not comments;
- share a sanitized case study only with the data owner's permission;
- disclose the dependency when publishing a wrapper or integration.

Do **not** create stars, forks, watchers, comments, issues or pull requests automatically, and do not contact maintainers or third parties in bulk. An agent may draft a report; a human publishes it after review. Promotion must follow real use and be disclosed.
