$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5){throw 'Run these tests in Windows PowerShell 5.1.'}
$source=Join-Path (Split-Path -Parent $PSScriptRoot) 'VM_DELETION_FINAL.ps1'
$errors=$null
$null=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$null,[ref]$errors)
if($errors.Count){throw ($errors|Out-String)}
foreach($test in @('Test-Core.ps1','Test-Preferences.ps1','Test-SqlLifecycle.ps1','Test-Startup.ps1')){
    & powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot $test)
    if($LASTEXITCODE -ne 0){throw "Offline test failed: $test"}
}
'ALL_OFFLINE_TESTS_PASSED'
