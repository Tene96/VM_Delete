$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Windows.Forms,System.Data
$source=[IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) 'VM_DELETION_FINAL.ps1'))
$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$errors)
if($errors.Count){throw ($errors|Out-String)}
foreach($name in @('Get-DeletionStepDefinitions','Test-DeletionStepEnabled','Get-DeletionStepSkipReason','Get-DeletionOSType','Sync-DeletionStepControls','Save-DeletionStepPreferences','Set-DefaultConfiguration','Apply-Config','Save-LocalConfig','Save-Settings','Write-DeletionConfigJson','Load-ConfigFromFile','Wait-JsonFileAvailable','Initialize-Configuration','Get-NoGuiActionSummary','Invoke-HeadlessCsvDeletion','New-VmDeletionReportStepStates','Update-VmDeletionReportStep','ConvertTo-VmDeletionReportRows','Import-VMDeletionCsv')){
    $f=$ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)|Select-Object -First 1
    if(-not $f){throw "Missing $name"};Invoke-Expression $f.Extent.Text
}
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message}}
$script:Logs=New-Object Collections.ArrayList
function Write-Log {param($Message,$Level) [void]$script:Logs.Add([string]$Message)}
function Show-AutoCloseNotice {throw 'Unexpected dialog in fixture.'}
$fixtureRoot=Join-Path (Join-Path $PSScriptRoot '.qa-fixtures') ('VMDeletionQa-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($fixtureRoot)
$script:LocalConfigCacheDir=Join-Path $fixtureRoot 'cache'
$script:LocalConfigCachePath=Join-Path $script:LocalConfigCacheDir 'config.json'
$script:SharedConfigPath=Join-Path $fixtureRoot 'shared.json'
$script:ExplicitConfigPath=Join-Path $fixtureRoot 'explicit.json'
$script:SharedConfigTimeoutMs=1;$script:LocalConfigTimeoutMs=1
$script:ConfigSource='Explicit JSON: fixture'
Set-DefaultConfiguration
foreach($step in (Get-DeletionStepDefinitions)){$script:DefaultSettings[$step.Key]=$true}
$script:DefaultSettings.SQLPassword='fixture-only'
$script:vCenterServers=@(@{Name='Fixture';Server='vc.fixture.invalid'})
foreach($definition in (Get-DeletionStepDefinitions)){
    Assert (Test-DeletionStepEnabled -Step $definition.Name) 'Legacy defaults must keep all six existing steps enabled.'
    Set-Variable -Scope Script -Name $definition.Run -Value (New-Object Windows.Forms.CheckBox)
}
$script:PersistentStepControls=@{}
foreach($definition in (Get-DeletionStepDefinitions)){
    $control=New-Object Windows.Forms.CheckBox;$control.Checked=$true
    $script:PersistentStepControls[$definition.Name]=$control
}
$script:PersistentStepControls.SCCM.Checked=$false
$script:PersistentStepControls.AD.Checked=$false
$script:PersistentStepControls.Archive.Checked=$false
Save-DeletionStepPreferences
Assert (-not $chkSCCM.Checked -and -not $chkSCCM.Enabled -and -not $chkAD.Checked -and -not $chkArchive.Enabled -and $chkVM.Checked) 'Persistent preferences did not disable the main run controls.'
Assert (Save-Settings) 'Could not save fixture config.'
$saved=Get-Content -LiteralPath $script:ExplicitConfigPath -Raw|ConvertFrom-Json
Assert ($saved.DefaultSettings.EnableSQLArchive -is [bool] -and -not $saved.DefaultSettings.EnableSQLArchive -and -not $saved.DefaultSettings.EnableSCCMDeletion) 'Saved JSON lost typed false values.'
$script:DefaultSettings=@{};$script:vCenterServers=@()
Assert (Initialize-Configuration) 'Explicit configuration failed to reload.'
Assert (-not (Test-DeletionStepEnabled -Step Archive) -and -not (Test-DeletionStepEnabled -Step AD) -and (Test-DeletionStepEnabled -Step VM)) 'Persisted values changed after restart.'
# Backward compatibility: a partial old JSON inherits operational defaults without erasing settings on rejection.
$legacy='{"DefaultSettings":{"SCCMServer":"sccm.fixture.invalid"},"vCenterServers":[{"Name":"Fixture","Server":"vc.fixture.invalid"}]}'|ConvertFrom-Json
Assert (Apply-Config -Config $legacy) 'Old partial JSON no longer loads.'
Assert ((Test-DeletionStepEnabled -Step SCCM) -and (Test-DeletionStepEnabled -Step AD) -and (Test-DeletionStepEnabled -Step Archive)) 'Missing new keys were treated as disabled.'
$before=$script:DefaultSettings
Assert (-not (Apply-Config -Config ('{"DefaultSettings":{"EnableADDeletion":"false"}}'|ConvertFrom-Json))) 'String false was accepted as a JSON boolean.'
Assert ([object]::ReferenceEquals($before,$script:DefaultSettings)) 'Rejected JSON changed active settings.'
Assert (-not (Apply-Config -Config ('{"DefaultSettings":{"SMTPPort":999999}}'|ConvertFrom-Json))) 'Out-of-range numeric value accepted.'
Assert (-not (Apply-Config -Config ('{"DefaultSettings":{"EnableVMDeletion":true},"vCenterServers":[{"Name":""}]}'|ConvertFrom-Json))) 'Invalid vCenter entry accepted.'
# Shared configuration is updated only by an explicit save; startup only caches it.
$script:ExplicitConfigPath='';$script:ConfigSource='Shared JSON: fixture'
$script:DefaultSettings.EnableADDeletion=$false
Assert (Save-Settings) 'Shared fixture save failed.'
$shared=Get-Content -LiteralPath $script:SharedConfigPath -Raw|ConvertFrom-Json
Assert (-not $shared.DefaultSettings.EnableADDeletion) 'Shared-source preference would be lost on restart.'
$script:DefaultSettings.EnableADDeletion=$true
Assert (Save-LocalConfig) 'Cache-only save failed.'
$shared=Get-Content -LiteralPath $script:SharedConfigPath -Raw|ConvertFrom-Json
Assert (-not $shared.DefaultSettings.EnableADDeletion) 'Startup cache wrote to shared source.'
$script:ExplicitConfigPath=Join-Path $fixtureRoot 'missing-explicit.json'
$failed=$false;try{[void](Initialize-Configuration)}catch{$failed=$true}
Assert $failed 'Explicit missing config silently fell back to another environment.'
'PERSISTENT_BOOLEAN_PREFERENCES_SAVE_RELOAD_LEGACY_VALIDATION_AND_SOURCE_POLICY_OK'

foreach($definition in (Get-DeletionStepDefinitions)){$script:DefaultSettings[$definition.Key]=$true}
Assert ((Get-DeletionStepSkipReason -Step SCCM -OSType Linux) -match 'Linux') 'Linux SCCM was not skipped.'
Assert ((Get-DeletionStepSkipReason -Step AD -OSType Linux) -eq '') 'Linux AD was automatically skipped.'
Assert ((Get-DeletionStepSkipReason -Step SCCM -OSType Windows) -eq '') 'Windows SCCM was skipped.'
Assert ((Get-DeletionOSType -OperatingSystem 'Ubuntu Linux (64-bit)') -eq 'Linux') 'Linux OS detection failed.'
Assert ((Get-DeletionOSType -GuestId 'windows2019srv_64Guest') -eq 'Windows') 'Windows OS detection failed.'
$script:CurrentVmDeletionReportRecord=@{VMName='linux-fixture';OSType='Linux';Errors=@();ReportSteps=(New-VmDeletionReportStepStates)}
Update-VmDeletionReportStep -TaskIndex 1 -Status Skipped -Message 'SCCM is skipped for Linux.'
foreach($key in @('AD','DNS','VMware','Archived')){$script:CurrentVmDeletionReportRecord.ReportSteps[$key].Status='Success'}
$row=ConvertTo-VmDeletionReportRows -SuccessList @($script:CurrentVmDeletionReportRecord)
Assert ($null -eq $row.SCCM -and -not $row.Error -and $row.StepDescriptions.SCCM -match 'Linux') 'Intentional Linux skip was rendered as a report failure.'
'LINUX_SKIPS_ONLY_SCCM_AND_REPORT_RETAINS_SKIP_REASON_OK'

# Actual NoGui loop with all external operations replaced at their boundaries.
$SkipSCCM=$false;$SkipAD=$false;$SkipDNS=$false;$SkipVM=$false;$SkipArchive=$false;$SkipEmail=$false
$Reason='Fixture';$script:RunEnvironment='Fixture'
$script:Calls=@{SCCM=0;AD=0;DNS=0;VM=0;Archive=0;Email=0;Connect=0}
$script:DefaultSettings.DeleteFromDisk=$false
function Get-Module {param($Name,$ErrorAction) return @{Name=$Name}}
function Set-PowerCLIConfiguration {param($InvalidCertificateAction,[switch]$Confirm,$ErrorAction)}
function Import-VMDeletionCsv {param($Path) @([pscustomobject]@{VMName='FIXTURE-LINUX';Environment='Fixture';OSType='Linux'},[pscustomobject]@{VMName='FIXTURE-WINDOWS';Environment='Fixture';OSType='Windows'})}
function Get-VCenterServerForEnvironment {param($EnvironmentName) return @{Name='Fixture';Server='vc.fixture.invalid'}}
function Connect-VIServer {param($Server,$ErrorAction) Assert ($Server -eq 'vc.fixture.invalid') 'Unexpected connection';$script:Calls.Connect++}
function Disconnect-VIServer {param($Server,[switch]$Force,[switch]$Confirm,$ErrorAction)}
function Get-VMDetailsForArchive {param($VMName) return @{VMName=$VMName;ServerName=$VMName;OperatingSystem=$(if($VMName -eq 'FIXTURE-LINUX'){'Ubuntu Linux'}else{'Windows Server'})}}
function Remove-SCCMDevice {param($ComputerName,$SCCMServer,$SiteCode) Assert ($ComputerName -eq 'FIXTURE-WINDOWS') 'SCCM was called for Linux';$script:Calls.SCCM++;return @{Success=$true;Found=$true}}
function Get-ADSearchBaseForEnvironment {param($EnvironmentName) return ''}
function Remove-ADDevice {param($ComputerName,$DomainController,$SearchBase) $script:Calls.AD++;return @{Success=$true;Found=$true}}
function Remove-DNSRecords {param($ComputerName,$DNSServer,$DNSZone) $script:Calls.DNS++;return @{Success=$true;Found=$true}}
function Get-Template {param($Name,$ErrorAction) return $null}
function Get-VM {param($Name,$ErrorAction) return @{Name=$Name;PowerState='PoweredOff'}}
function Remove-VM {param($VM,[switch]$Confirm,[switch]$DeletePermanently,$ErrorAction) $script:Calls.VM++}
function Save-VMDeletionToArchive {param($VMDetails,$Reason,$SCCMDeleted,$ADDeleted,$DNSDeleted,$VMDeleted,$Status) $script:Calls.Archive++;return @{Success=$true}}
function Send-VMDeletionNotification {param($VMDetails,$Reason,$OSType,$DeletionSuccess) $script:Calls.Email++}
$result=Invoke-HeadlessCsvDeletion -Path 'fixture.csv'
Assert ($result -eq 0 -and $script:Calls.SCCM -eq 1 -and $script:Calls.AD -eq 2 -and $script:Calls.VM -eq 2 -and $script:Calls.Archive -eq 2 -and $script:Calls.Email -eq 2) 'Mixed OS NoGui step policy is wrong.'
foreach($definition in (Get-DeletionStepDefinitions)){$script:DefaultSettings[$definition.Key]=$false}
foreach($key in @($script:Calls.Keys)){$script:Calls[$key]=0}
$result=Invoke-HeadlessCsvDeletion -Path 'fixture.csv'
Assert ($result -eq 0 -and @($script:Calls.Values|Where-Object{$_ -ne 0}).Count -eq 0) 'Globally disabled steps still called an infrastructure boundary.'
$summary=Get-NoGuiActionSummary
Assert ($summary.Enabled -eq 'None') 'NoGui summary ignores persistent flags.'
foreach($definition in (Get-DeletionStepDefinitions)){$script:DefaultSettings[$definition.Key]=$true}
$SkipSCCM=$true;$SkipAD=$true;$SkipDNS=$true;$SkipVM=$true;$SkipArchive=$true;$SkipEmail=$true
$result=Invoke-HeadlessCsvDeletion -Path 'fixture.csv'
Assert ($result -eq 0 -and @($script:Calls.Values|Where-Object{$_ -ne 0}).Count -eq 0) 'CLI Skip switches re-enabled an operation.'
'NOGUI_MIXED_OS_PERSISTENT_FLAGS_AND_CLI_SKIP_PRECEDENCE_OK'
foreach($control in $script:PersistentStepControls.Values){$control.Dispose()}
foreach($definition in (Get-DeletionStepDefinitions)){(Get-Variable $definition.Run).Value.Dispose()}
