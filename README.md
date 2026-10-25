# VM Deletion Tool

A Windows PowerShell GUI for removing VMware VMs and templates, with optional SCCM, Active Directory and DNS cleanup, SQL archiving and HTML email reports. The main window keeps connection, selection, deletion progress and logs together. Advanced settings uses a separate settings window with an inline SQL table viewer.

## Quick start

Use Windows PowerShell 5.1 on Windows. The script uses WPF and WinForms. Install VMware PowerCLI for the vCenter workflow. Only install and configure the optional integration dependencies you intend to enable: ActiveDirectory and DnsServer modules for those cleanup operations, SCCM WMI access, SQL Server access, and an SMTP relay for email.

1. Download this repository's release files into a new directory.
2. Copy `config.example.json` to `config.json`.
3. Edit `vCenterServers` and `DefaultSettings.vCenterServer` for your environment.
4. Start the tool with the explicit configuration file:

```powershell
Copy-Item .\config.example.json .\config.json
powershell.exe -NoProfile -STA -File .\VM_DELETION_FINAL.ps1 -ConfigPath .\config.json
```

The public example enables VM/template removal only. SCCM, AD, DNS, SQL archive and email are disabled until you configure and enable them. Deletion confirmation is enabled. `DeleteFromDisk` starts as `false` in the public example; enable it only when you intend to remove the VM files permanently.

The script does not install optional modules automatically for its main workflow. Use an account with the permissions required by each enabled integration. No vCenter password is included in the script or example JSON.

## Unlocking Advanced settings

The default settings password is **`admin`**.

Open **Advanced settings**, choose **Unlock settings**, and enter `admin`. Choose **Lock settings** to prevent accidental edits. This is a convenience lock in the GUI, not an authorization or credential security boundary. Real access is controlled by your Windows, VMware, AD, DNS, SCCM and SQL permissions.

## Six persistent workflow choices

Open **Advanced settings > Execution**. Each card enables a workflow step. Clear a card and choose **Save configuration** to disable that step for subsequent runs. Enable the card and save again to restore it. A disabled step also appears unchecked and unavailable in the main window. Main-window options can skip an enabled step for the current run only.

| Card | JSON boolean | Effect when false |
| --- | --- | --- |
| SCCM | `EnableSCCMDeletion` | No SCCM device cleanup |
| Active Directory | `EnableADDeletion` | No AD computer-account cleanup |
| DNS | `EnableDNSDeletion` | No DNS cleanup |
| VM / template | `EnableVMDeletion` | No VM/template removal; the GUI still uses vCenter to select machines |
| SQL archive | `EnableSQLArchive` | No automatic SQL archive connection or insert |
| Email | `EnableEmailNotification` | No personal deletion report or batch summary |

These values must be JSON booleans (`true` or `false`), not quoted strings. Older configurations that omit the new workflow keys retain their previous enabled behavior. The existing `EnableEmailNotification` key continues to control email.

### Linux behavior

Recognized Linux machines automatically skip **SCCM only**. AD remains enabled unless you disable its card or skip it for the current run. DNS, VM removal, archive and email follow their own settings. The log and report identify an intentional skip separately from a failure. An unknown operating system is not assumed to be Linux.

## Configuration locations and saving

You can keep the JSON configuration on a **network share** for a shared team configuration, or **locally** for one user or computer. Connection settings and the six persistent workflow choices are saved in the same JSON file.

### Option 1: Local JSON file

Copy `config.example.json` to `config.json`, edit it for your environment, and select the local file when starting the tool:

```powershell
powershell.exe -NoProfile -STA -File .\VM_DELETION_FINAL.ps1 -ConfigPath .\config.json
```

You can also supply an absolute local path, such as `C:\Tools\VMDeletion\config.json`. The file must already exist and contain valid configuration. **Save configuration** writes back to the selected file and refreshes the per-user local cache. This lets you keep a persistent local configuration outside the TEMP directory.

An explicit `-ConfigPath` can also point to a network file, but it is authoritative: an unreadable or invalid explicit file stops startup. It does not fall back to another configuration. If both path parameters are supplied, `-ConfigPath` takes precedence.

### Option 2: Shared network JSON with local fallback

Place your edited `config.json` in an existing network share and start the tool with its UNC path:

```powershell
powershell.exe -NoProfile -STA -File .\VM_DELETION_FINAL.ps1 -SharedConfigPath '\\fileserver\tools\vm-deletion\config.json'
```

The shared JSON is loaded first. After a successful load, the tool refreshes a local copy at:

```text
%TEMP%\VMDeletionTool\config.json
```

If the shared JSON is unavailable or invalid on a later startup, the tool uses the valid local copy. **Save configuration** writes to the shared file when the shared source is active, and also refreshes the local copy. Startup itself only refreshes the local copy; it does not write to the shared source.

Users need read access to load a shared file and permission to create/replace files in its directory to save changes. Shared configuration changes apply to other users who load that same file. Check the save status before closing; a failed save does not mean the changes were persisted successfully.

If the tool starts from the local fallback, changes are saved locally. They are not automatically merged back into the shared JSON. When the shared source becomes available again, it takes priority on the next startup.

### Automatic local storage and startup order

Without `-ConfigPath`, the startup order is:

1. The shared JSON, if a shared path is configured and contains valid configuration.
2. The per-user local JSON at `%TEMP%\VMDeletionTool\config.json`.
3. Built-in defaults, which are used to create a new local JSON when possible.

The public script has no organization-specific shared path. If you start it without either path parameter, it uses the local TEMP JSON or creates one from its defaults. TEMP files can be cleaned by Windows or maintenance tools; use `-ConfigPath` with a persistent local file if you want to keep your configuration outside TEMP.

The repository contains an example configuration only. Keep your real `config.json`, exports, logs and credentials out of Git; `.gitignore` excludes them. SQL integrated authentication is the public default. If you use SQL authentication, credentials in your private configuration/cache are sensitive: restrict access to those files and do not commit them. Do not publish encrypted credential blobs either. Legacy DPAPI-protected values require their original Windows user/machine to decode.

## SQL archive and table viewer

Each enabled archive operation creates its own SQL connection, inserts one record in a transaction, verifies the insert acknowledgement and commits. It releases the command, transaction and connection on success or failure. You do not need to choose **Check connection and table** or **Display SQL table** before deleting a VM. A failed archive is reported as a failure, not silently skipped.

Set `SQLServer`, `SQLDatabase`, `SQLTable` and the authentication fields. A schema-qualified table name such as `dbo.VMDeletionArchive` is supported. Create the database and grant the appropriate access separately. The check button can offer to create a missing archive table; this is an explicit operator action and requires table-creation rights. Normal archiving does not create a database/table automatically.

**Display SQL table** opens a compact, scrollable table inside **SQL & archive**, with local search. Typing a search does not issue another SQL query. The inline export uses the visible filtered snapshot; the top **Export table** action reads and exports the full table. CSV and XLSX export are available. Explicit viewer/check actions can still query SQL when automatic archiving is disabled.

## Email

Configure the SMTP relay and recipient groups before enabling email. Each VM's report goes to the initiating user and the resolved server owner, with folder-owner fallback. Supported annotations include `SRVOWNER`, legacy `srvOWN1`, and `FOLDEROWNER`. Email addresses can be supplied directly in owner fields; account/group names require AD resolution.

One batch summary goes to the configured JSON recipient groups, with duplicate addresses removed. The existing routing uses `GroupsCommon` and `GroupsWindows` for Windows, `GroupsWindows` for templates, and `GroupsCommon` for Linux. Linux clusters matching the legacy patterns `LINUXGPU`, `LINUXHANA`, `LINUXNOORA` or `LINUXORASYN` additionally include `GroupsWindows` and `GroupsLinux`. Review these rules for your deployment before enabling email.

Disabling AD cleanup does not disable separate AD lookups used to resolve email addresses. If your environment has no AD, keep email disabled or use direct email addresses where supported. The initiating user's email resolution still uses AD.

## CSV and unattended operation

The minimum CSV column is `VMName`; `Environment` selects a configured vCenter. An optional `OSType` (`Windows`, `Linux`, `Other`) provides a hint when VM metadata is unavailable, for example when VM removal is skipped. Available VMware metadata takes precedence. `samples/vms.example.csv` contains fictional names only.

```powershell
# Validation run: all six actions explicitly skipped.
powershell.exe -NoProfile -File .\VM_DELETION_FINAL.ps1 -ConfigPath .\config.json -NoGui -CsvPath .\samples\vms.example.csv -SkipSCCM -SkipAD -SkipDNS -SkipVM -SkipArchive -SkipEmail

# Show usage without connecting to infrastructure or loading configuration.
powershell.exe -NoProfile -File .\VM_DELETION_FINAL.ps1 -Help
```

Persistent workflow choices apply to unattended operation too. A `-Skip...` switch can disable an additional action; it cannot re-enable an action disabled in JSON. Exit code `0` means no record-level failures were reported; `1` means startup or one or more record operations failed. An all-skipped validation run does not prove live integration permissions or connectivity.

## Validation status

See `QA_REPORT.md` for the local parser, configuration, GUI and mocked workflow checks. Live deletion, production SQL inserts, real SMTP delivery and remote permissions must be tested in your controlled environment. This repository preparation does not publish anything to GitHub automatically.

Run the included checks in Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -STA -File .\tests\Run-OfflineTests.ps1
```

These tests replace deletion, SQL and email boundaries with fixtures. The startup check invokes the real script with all six actions skipped. Generated test files use a unique ignored directory under tests/.qa-fixtures; the tests do not read your real configuration. The included core tests use fictional `example.invalid` addresses and dummy credentials.
