# Local QA report

Validation date: 2026-10-25. Runtime checks used Windows PowerShell 5.1 on Windows. The runtime fixes were checked in both the operational script and the sanitized public package. This report describes local evidence, not production integration certification.

## Passed checks

- PowerShell 5.1 parser: no syntax errors in the modified operational script or public script.
- PSScriptAnalyzer: no Error diagnostics after removing unused password helpers and passing SQL authentication as `PSCredential`. Legacy Warning diagnostics remain; see below.
- Configuration: six typed workflow booleans, save/reload, explicit-file selection, shared-source save versus cache-only startup, partial legacy JSON compatibility, invalid boolean/numeric/vCenter data rejection, and refusal to silently replace an invalid explicit configuration with another environment.
- Advanced settings GUI: all six cards exist, follow the settings lock, save through the top Save configuration button, survive reopening, disable their corresponding main-window choices, and remain reachable by scrolling at minimum window size.
- Mixed Windows/Linux workflows: Windows can use SCCM; recognized Linux skips SCCM only and still executes AD when selected. Persistently disabled integrations do not call their operational boundary. CLI Skip switches cannot re-enable a disabled integration.
- Intentional skips retain their reason in logs/reports and do not become false failures. Disabling VM removal produces an explicit skipped step.
- SQL archive: each record gets a distinct connection, with no manual connection check or table-view prerequisite. Commit acknowledgement, parameterization, open/begin/insert/zero-acknowledgement/commit failures, rollback, cleanup, and independent recovery for the next record were exercised using fake SQL objects.
- SQL table creation: the previously undefined creation action now has an implementation. Mocked success/failure, schema-qualified identifiers, input rejection and resource cleanup passed. No real table was created.
- Inline SQL viewer: local search, literal special characters, typed values, zero-match export, filtered export, refresh preserving the query and full-table export were checked with fixture data.
- CSV/XLSX exports: empty data, typed cells, UTF-8 text, repeated export and table-identifier validation passed locally.
- Confirmation, Connecting state, and the GUI deletion -> archive -> personal report -> batch report sequence passed with infrastructure operations mocked. SQL failure remains a failed report even when VM removal succeeded.
- Personal and batch email recipient routing, owner/folder fallback, de-duplication, HTML escaping and disabled email behavior were checked with mocked AD/SMTP.
- Public startup: the actual public script's Help worked with an intentionally missing ConfigPath; an all-actions-skipped NoGui invocation returned 0; a missing explicit configuration stopped startup.
- Public content review: organization-specific built-in hosts, domains, shared path and default SQL credentials were removed from the distributed script/configuration. The example uses fictional values, and `.gitignore` excludes real configuration and operational data. Test credentials are dummy fixture values only.

## Remaining warnings and limits

The static analyzer still reports legacy Warning categories: event-handler parameter names (`sender`, `eventArgs`), global state, function naming, unused parameters/assignments, intentional empty catch blocks, Write-Host/alias usage, SCCM WMI use and state-changing functions without a ShouldProcess contract. These were not suppressed, and this release does not claim warning-free code.

Do not interpret the script as supporting `-WhatIf` throughout. No deletion test in this QA contacted VMware, SCCM, AD or DNS. SQL writes, live table creation, SMTP delivery, real permissions, shared-file ACL behavior and production credential handling remain to be verified in a controlled environment. An all-skipped run is validation, not a deletion test.

Private configuration/cache files may contain SQL authentication credentials when that mode is used. The public example uses integrated authentication and has no credentials. The default `admin` settings password is a GUI editing convenience, not an access-control boundary.

## Reproducing the portable checks

```powershell
powershell.exe -NoProfile -STA -File .\tests\Run-OfflineTests.ps1
```

The included tests run in separate processes, mock the operational boundaries, and use unique ignored fixture directories under tests/.qa-fixtures. The additional interactive GUI checks described above were performed locally during preparation.
