$ErrorActionPreference='Stop'
$package=Split-Path -Parent $PSScriptRoot
$scriptPath=Join-Path $package 'VM_DELETION_FINAL.ps1'
$fixtureRoot=Join-Path (Join-Path $PSScriptRoot '.qa-fixtures') ('VMDeletionStartupQa-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixtureRoot)
$config=Get-Content -LiteralPath (Join-Path $package 'config.example.json') -Raw|ConvertFrom-Json
$config.DefaultSettings.LogPath=Join-Path $fixtureRoot 'logs'
$configPath=Join-Path $fixtureRoot 'config.json'
$config|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $configPath -Encoding UTF8
$helpOutput=& powershell.exe -NoProfile -File $scriptPath -ConfigPath (Join-Path $fixtureRoot 'does-not-exist.json') -Help 2>&1
if($LASTEXITCODE -ne 0 -or ($helpOutput -join "`n") -notmatch 'ConfigPath <path-to-config.json>'){throw 'Help tried to load configuration or omitted the new parameters.'}
$helpOutput|Set-Content -LiteralPath (Join-Path $fixtureRoot 'help.txt')
$output=& powershell.exe -NoProfile -File $scriptPath -ConfigPath $configPath -NoGui -CsvPath (Join-Path $package 'samples\vms.example.csv') -SkipSCCM -SkipAD -SkipDNS -SkipVM -SkipArchive -SkipEmail 2>&1
$output|Set-Content -LiteralPath (Join-Path $fixtureRoot 'all-skipped.txt')
if($LASTEXITCODE -ne 0 -or ($output -join "`n") -notmatch 'Completed successfully. ExitCode=0'){throw 'Actual public all-skipped invocation failed.'}
$savedPreference=$ErrorActionPreference;$ErrorActionPreference='Continue'
$failureOutput=& powershell.exe -NoProfile -File $scriptPath -ConfigPath (Join-Path $fixtureRoot 'does-not-exist.json') -NoGui -CsvPath (Join-Path $package 'samples\vms.example.csv') -SkipSCCM -SkipAD -SkipDNS -SkipVM -SkipArchive -SkipEmail 2>&1
$failureCode=$LASTEXITCODE;$ErrorActionPreference=$savedPreference
if($failureCode -eq 0 -or ($failureOutput -join "`n") -notmatch 'Invalid or unreadable configuration'){throw 'Missing explicit configuration did not stop startup.'}
'ACTUAL_PUBLIC_HELP_ALL_SKIPPED_NOGUI_AND_INVALID_CONFIG_STARTUP_OK'
