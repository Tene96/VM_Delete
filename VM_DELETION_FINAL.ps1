#requires -Version 5.1

### Created By Tene ###
### January 2026 ###
### Version:1.0
<#
.SYNOPSIS
    GUI tool to delete a VM/Template and clean up related records from vCenter, SCCM, AD, and DNS.

.DESCRIPTION
    VM Deletion Tool supports two operating modes:

    1. Interactive GUI mode
       Run the script without parameters.
       The GUI lets the operator select a vCenter, load VMs, import a CSV list,
       select deletion actions, confirm deletion, review progress, and check archive data.

       Example:
       powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\VM_DELETION_FINAL.ps1"

    2. Scheduled / headless mode for Task Scheduler
       Run the script with -NoGui and -CsvPath.
       No GUI windows are opened.
       The script reads a CSV file, groups servers by Environment/vCenter,
       connects to the relevant vCenter when VM deletion is enabled,
       performs the requested cleanup steps, writes logs, and returns an exit code.

       Example:
       powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\VM_DELETION_FINAL.ps1" -NoGui -CsvPath "C:\Input\vms.csv" -Environment PROD

       Safe validation example without deleting anything:
       powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\VM_DELETION_FINAL.ps1" -NoGui -CsvPath "C:\Input\vms.csv" -Environment PROD -SkipSCCM -SkipAD -SkipDNS -SkipVM -SkipArchive -SkipEmail

    Exit codes in NoGui mode:
       0 = completed without VM-level failures
       1 = one or more failures occurred, or startup validation failed

    Required CSV structure:
       VMName
       SERVER01
       SERVER02

    Recommended CSV structure:
       VMName,Environment
       SERVER01,PROD
       SERVER02,MGMT
       SERVER03,VDI

    Alternative CSV structure:
       VMName,vCenter
       SERVER01,Prod
       SERVER02,Mgmt
       SERVER03,Vdi

    CSV columns:
       VMName      Required. VM or Template name to process.
       Environment Optional. Used to select vCenter and AD OU policy.
       vCenter     Optional. Can be used instead of Environment.

    Environment / AD OU behavior:
       AD deletion uses ADSearchBaseByEnvironment when available.
       PROD and MGMT use the same default OU.
       VDI uses a separate default OU.
       These values can be overridden in config.json.

    Default AD OU map in code:
       ADSearchBaseByEnvironment = @{
           PROD = "OU=Servers,DC=system,DC=co,DC=il"
           MGMT = "OU=Servers,DC=system,DC=co,DC=il"
           VDI  = "OU=VDI,DC=system,DC=co,DC=il"
       }

    Configuration files:
       Local JSON:
           $env:TEMP\VMDeletionTool\config.json

       Network JSON:
           Tene\VMDeletionTool\config.json

       Startup behavior:
           The script waits briefly for JSON availability.
           If no local or network JSON is found, it uses the default settings embedded in the code.
           In GUI mode, a short auto-closing notice is shown.
           In NoGui mode, no notice window is shown.

    Main cleanup actions:
       SCCM    Delete SCCM device record.
       AD      Delete computer object from the OU selected by Environment.
       DNS     Delete A/PTR/AAAA records.
       VM      Delete VM or Template from vCenter.
       Archive Write result to SQL archive table when SQL is configured and checked.
       Email   Send deletion notification when email is enabled.

    NoGui skip switches:
       -SkipSCCM
       -SkipAD
       -SkipDNS
       -SkipVM
       -SkipArchive
       -SkipEmail

    Task Scheduler notes:
       Use a service account with the required permissions for the enabled actions.
       VMware PowerCLI is required when VM deletion is not skipped.
       ActiveDirectory module is required when AD deletion is not skipped.
       DnsServer module is required when DNS deletion is not skipped.
       SCCM access is required when SCCM deletion is not skipped.
       SQL connectivity is required when Archive is not skipped.

.PARAMETER Help
    Shows NoGui usage examples, CSV format, and exit codes, then exits without opening the GUI.

.PARAMETER NoGui
    Runs in headless mode without opening the GUI. Intended for Task Scheduler.

.PARAMETER CsvPath
    Path to a CSV file containing at least a VMName column.

.PARAMETER Environment
    Default environment used when the CSV does not include Environment or vCenter.
    Common values: PROD, MGMT, VDI.

.PARAMETER Reason
    Reason text written to archive/email in NoGui mode.

.PARAMETER ConfigPath
    Explicit JSON configuration file. Save configuration writes back to this file.
    An invalid explicit file stops startup; it never falls back to another environment.

.PARAMETER SharedConfigPath
    Optional override for the shared JSON source. Without ConfigPath, source order is
    shared JSON, local TEMP JSON, and built-in defaults.

.PARAMETER SkipSCCM
    Skips SCCM cleanup.

.PARAMETER SkipAD
    Skips AD cleanup.

.PARAMETER SkipDNS
    Skips DNS cleanup.

.PARAMETER SkipVM
    Skips VM/Template deletion from vCenter.

.PARAMETER SkipArchive
    Skips SQL archive write.

.PARAMETER SkipEmail
    Skips email notification.

.NOTES
    Requires: VMware PowerCLI, SCCM module, ActiveDirectory module, DNS Server tools
    v2.1 - Added Template deletion support + Linux cluster email logic
#>
param(
    [switch]$Help,
    [switch]$NoGui,
    [string]$CsvPath,
    [string]$Environment = "PROD",
    [string]$Reason = "Scheduled CSV deletion",
    [switch]$SkipSCCM,
    [switch]$SkipAD,
    [switch]$SkipDNS,
    [switch]$SkipVM,
    [switch]$SkipArchive,
    [switch]$SkipEmail,
    [string]$ConfigPath,
    [string]$SharedConfigPath
)


# Native WPF needs STA. Headless/help execution keeps its original host.
if (-not $NoGui -and -not $Help -and [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $wizardHostArgs = @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',$PSCommandPath)
    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        if ($entry.Value -is [System.Management.Automation.SwitchParameter]) {
            if ($entry.Value.IsPresent) { $wizardHostArgs += '-' + $entry.Key }
        }
        else { $wizardHostArgs += '-' + $entry.Key; $wizardHostArgs += [string]$entry.Value }
    }
    & (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') @wizardHostArgs
    return
}

CLS

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Global Variables and Settings

# ============================================
# Configuration Sources
# ============================================
# 1. Shared JSON config - primary source of truth.
$script:SharedConfigPath = ''
if($PSBoundParameters.ContainsKey('SharedConfigPath')){$script:SharedConfigPath=$SharedConfigPath}
$script:ExplicitConfigPath=if([string]::IsNullOrWhiteSpace($ConfigPath)){''}else{[IO.Path]::GetFullPath($ConfigPath)}
$script:SharedConfigTimeoutMs = 8000

# 2. Per-user TEMP JSON cache - used only when shared config is unavailable.
$script:LocalConfigCacheDir = Join-Path $env:TEMP 'VMDeletionTool'
$script:LocalConfigCachePath = Join-Path $script:LocalConfigCacheDir 'config.json'
$script:LocalConfigTimeoutMs = 1000

# 3. Built-in defaults are used only if shared and local JSON are unavailable or invalid.
$script:ConfigSource = ''
# ============================================
# Global Variables
# ============================================
$script:vcConnected = $false
$script:allVMs = @()
$script:sqlTableExists = $false
$script:vCenterServers = @()
$script:DefaultSettings = @{}
$script:SelectedVMNames = [System.Collections.ArrayList]@()
$script:IsUpdatingList = $false
$script:allVMsData = @()
$script:allTemplatesData = @()
$script:IncludeTemplates = $false
$script:settingsPassword = "admin"
$script:RunEnvironment = $Environment
$script:CSVImportData = @()
$script:txtLogs = $null
$script:logsTab = $null
$script:IsShowingLogErrorNotice = $false

$tasks = @("Connect to vCenter", "Delete SCCM Record", "Delete AD Record", "Delete DNS Record", "Delete VM", "Archive to SQL", "Send Notification")
$statusLabels = @()
$resultLabels = @()
$yPos = 50

# ============================================
# Configuration Functions
# ============================================


function Wait-JsonFileAvailable {
    param(
        [string]$Path,
        [int]$TimeoutMs = 1000,
        [int]$IntervalMs = 100
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    do {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            return $Path
        }

        Start-Sleep -Milliseconds $IntervalMs
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)

    return $null
}

function Show-AutoCloseNotice {
    param(
        [string]$Message,
        [string]$Title = "VM Deletion Tool",
        [int]$DurationMs = 2200
    )

    try {
        $noticeForm = New-Object System.Windows.Forms.Form
        $noticeForm.Text = $Title
        $noticeForm.Size = New-Object System.Drawing.Size(460, 120)
        $noticeForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $noticeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
        $noticeForm.MaximizeBox = $false
        $noticeForm.MinimizeBox = $false
        $noticeForm.TopMost = $true
        $noticeForm.BackColor = [System.Drawing.Color]::FromArgb(54, 54, 68)

        $noticeLabel = New-Object System.Windows.Forms.Label
        $noticeLabel.Location = New-Object System.Drawing.Point(18, 18)
        $noticeLabel.Size = New-Object System.Drawing.Size(410, 48)
        $noticeLabel.Text = $Message
        $noticeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $noticeLabel.ForeColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
        $noticeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $noticeForm.Controls.Add($noticeLabel)

        $noticeTimer = New-Object System.Windows.Forms.Timer
        $noticeTimer.Interval = $DurationMs
        $noticeTimer.Add_Tick({
            $noticeTimer.Stop()
            $noticeForm.Close()
        })
        $noticeForm.Add_Shown({ $noticeTimer.Start() })
        [void]$noticeForm.ShowDialog()
        $noticeTimer.Dispose()
        $noticeForm.Dispose()
    }
    catch {
        return
    }
}

function Load-ConfigFromFile {
    param(
        [string]$Path,
        [int]$TimeoutMs = 1000
    )
    
    try {
        $availablePath = Wait-JsonFileAvailable -Path $Path -TimeoutMs $TimeoutMs
        if (-not $availablePath) { return $null }
        
        $jsonContent = Get-Content -LiteralPath $availablePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($jsonContent)) { return $null }

        return ($jsonContent | ConvertFrom-Json)
    }
    catch {
        ##Write-Host "Failed to load config from $Path : $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Get-DeletionStepDefinitions {
    @(
        @{Name='SCCM';Key='EnableSCCMDeletion';Editor='aStepSccm';Run='chkSCCM';Wizard='wSCCM';Title='SCCM';Description='Remove the device record. Skipped for Linux.'},
        @{Name='AD';Key='EnableADDeletion';Editor='aStepAd';Run='chkAD';Wizard='wAD';Title='Active Directory';Description='Remove the computer account, including Linux when selected.'},
        @{Name='DNS';Key='EnableDNSDeletion';Editor='aStepDns';Run='chkDNS';Wizard='wDNS';Title='DNS';Description='Remove the server DNS records.'},
        @{Name='VM';Key='EnableVMDeletion';Editor='aStepVm';Run='chkVM';Wizard='wVM';Title='VM / template';Description='Remove the VM or template from vCenter.'},
        @{Name='Archive';Key='EnableSQLArchive';Editor='aStepSql';Run='chkArchive';Wizard='wArchive';Title='SQL archive';Description='Write each result using an independent SQL connection.'},
        @{Name='Email';Key='EnableEmailNotification';Editor='aStepEmail';Run='chkEmail';Wizard='wEmail';Title='Email';Description='Send personal reports and the batch summary.'}
    )
}

function Test-DeletionStepEnabled {
    param([ValidateSet('SCCM','AD','DNS','VM','Archive','Email')][string]$Step)
    $definition = Get-DeletionStepDefinitions | Where-Object { $_.Name -eq $Step } | Select-Object -First 1
    if (-not $script:DefaultSettings.ContainsKey($definition.Key)) { return $true }
    if ($script:DefaultSettings[$definition.Key] -isnot [bool]) { throw "Setting '$($definition.Key)' must be a JSON boolean (true or false)." }
    return $script:DefaultSettings[$definition.Key]
}

function Get-DeletionOSType {
    param([string]$Preferred, [string]$OperatingSystem, [string]$GuestId)
    $evidence = "$OperatingSystem $GuestId"
    if ($evidence -match 'Windows') { return 'Windows' }
    if ($evidence -match 'Linux|Ubuntu|CentOS|Red Hat|RHEL|Debian|SUSE|Fedora|Rocky|AlmaLinux|Oracle Linux') { return 'Linux' }
    if ($Preferred -in @('Windows','Linux','Other')) { return $Preferred }
    return 'Unknown'
}

function Get-DeletionStepSkipReason {
    param([ValidateSet('SCCM','AD','DNS','VM','Archive','Email')][string]$Step, [string]$OSType, [bool]$Selected=$true)
    if (-not (Test-DeletionStepEnabled -Step $Step)) { return 'Disabled in Advanced settings > Execution.' }
    if (-not $Selected) { return 'Not selected for this run.' }
    if ($Step -eq 'SCCM' -and $OSType -eq 'Linux') { return 'SCCM is skipped for Linux.' }
    return ''
}

function Sync-DeletionStepControls {
    param([bool]$Available=$true, [switch]$ResetSelection)
    foreach ($definition in (Get-DeletionStepDefinitions)) {
        $variable = Get-Variable -Name $definition.Run -ErrorAction SilentlyContinue
        if (-not $variable -or -not $variable.Value) { continue }
        $control = $variable.Value
        $enabled = Test-DeletionStepEnabled -Step $definition.Name
        if ($ResetSelection -or -not $enabled) { $control.Checked=$enabled }
        $control.Enabled=($Available -and $enabled)
    }
}

function Save-DeletionStepPreferences {
    foreach ($definition in (Get-DeletionStepDefinitions)) {
        $script:DefaultSettings[$definition.Key] = [bool]$script:PersistentStepControls[$definition.Name].Checked
    }
    Sync-DeletionStepControls -ResetSelection
}

function Write-DeletionConfigJson {
    param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)][string]$Json)
    $absolute=[IO.Path]::GetFullPath($Path)
    $directory=[IO.Path]::GetDirectoryName($absolute)
    [void][IO.Directory]::CreateDirectory($directory)
    $temporary=Join-Path $directory ([IO.Path]::GetRandomFileName()+'.tmp')
    $backup=Join-Path $directory ([IO.Path]::GetRandomFileName()+'.bak')
    try {
        [IO.File]::WriteAllText($temporary,$Json,(New-Object Text.UTF8Encoding($true)))
        if([IO.File]::Exists($absolute)){[IO.File]::Replace($temporary,$absolute,$backup)}
        else{[IO.File]::Move($temporary,$absolute)}
    }
    finally {
        foreach($artifact in @($temporary,$backup)){if([IO.File]::Exists($artifact)){[IO.File]::Delete($artifact)}}
    }
}
function Apply-Config {
    param($Config)
    try {
        if($null -eq $Config){return $false}
        $names=@($Config.PSObject.Properties.Name)
        $settingsSource=if($names -contains 'DefaultSettings'){$Config.DefaultSettings}else{$Config}
        if(-not $settingsSource -or $settingsSource -is [string] -or $settingsSource -is [array]){return $false}
        $previousSettings=$script:DefaultSettings; $previousVcs=$script:vCenterServers
        try {
            Set-DefaultConfiguration
            $candidate=@{}+$script:DefaultSettings
            $candidateVcs=$script:vCenterServers
        }
        finally {$script:DefaultSettings=$previousSettings; $script:vCenterServers=$previousVcs}
        $copied=0
        foreach($property in $settingsSource.PSObject.Properties){
            if($property.Name -in @('vCenterServers','DefaultSettings','LastSaved','SavedBy')){continue}
            $candidate[$property.Name]=$property.Value; $copied++
        }
        if($copied -eq 0){return $false}
        foreach($definition in (Get-DeletionStepDefinitions)){
            if($definition.Name -ne 'Email' -and $settingsSource.PSObject.Properties.Name -notcontains $definition.Key){$candidate[$definition.Key]=$true}
        }
        foreach($key in @('EnableLogging','ConfirmDeletion','DeleteFromDisk','SQLIntegratedAuth','SMTPUseSSL')+@((Get-DeletionStepDefinitions).Key)){
            if($candidate[$key] -isnot [bool]){return $false}
        }
        foreach($rule in @(@('PowerOffTimeout',10,300),@('SMTPPort',1,65535))){
            $value=0
            if(-not [int]::TryParse([string]$candidate[$rule[0]],[ref]$value) -or $value -lt $rule[1] -or $value -gt $rule[2]){return $false}
            $candidate[$rule[0]]=$value
        }
        if($names -contains 'vCenterServers'){
            if(-not $Config.vCenterServers -or $Config.vCenterServers -is [string]){return $false}
            $candidateVcs=@()
            foreach($vc in @($Config.vCenterServers)){
                if([string]::IsNullOrWhiteSpace($vc.Name) -or [string]::IsNullOrWhiteSpace($vc.Server)){return $false}
                $candidateVcs+=@{Name=[string]$vc.Name;Server=[string]$vc.Server}
            }
            if(@($candidateVcs | Group-Object Name | Where-Object Count -gt 1).Count){return $false}
        }
        elseif($previousVcs){$candidateVcs=$previousVcs}
        if($candidate.SQLPasswordEncrypted -eq $true -and -not [string]::IsNullOrEmpty($candidate.SQLPassword)){
            $pointer=[IntPtr]::Zero
            try {
                $secure=ConvertTo-SecureString -String $candidate.SQLPassword -ErrorAction Stop
                $pointer=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
                $candidate.SQLPassword=[Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
            }
            finally {if($pointer -ne [IntPtr]::Zero){[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)}}
            $candidate.Remove('SQLPasswordEncrypted')
        }
        if([string]::IsNullOrWhiteSpace($candidate.LogPath)){$candidate.LogPath=$script:LocalConfigCacheDir}
        $script:DefaultSettings=$candidate; $script:vCenterServers=$candidateVcs
        return $true
    }
    catch {return $false}
}
function Initialize-Configuration {
    $script:ConfigSource = ""
    if(-not [string]::IsNullOrWhiteSpace($script:ExplicitConfigPath)){
        $explicitConfig=Load-ConfigFromFile -Path $script:ExplicitConfigPath -TimeoutMs $script:LocalConfigTimeoutMs
        if(-not $explicitConfig -or -not (Apply-Config -Config $explicitConfig)){throw "Invalid or unreadable configuration: $script:ExplicitConfigPath"}
        $script:ConfigSource="Explicit JSON: $script:ExplicitConfigPath"
        return $true
    }

    # 1. Try the shared network config first.
    $networkConfig = Load-ConfigFromFile -Path $script:SharedConfigPath -TimeoutMs $script:SharedConfigTimeoutMs
    if ($networkConfig -and (Apply-Config -Config $networkConfig)) {
        $script:ConfigSource = "Shared JSON: $script:SharedConfigPath"
        Save-LocalConfig | Out-Null
        return $true
    }

    # 2. If network is unavailable, use a valid local TEMP config.
    $localConfig = Load-ConfigFromFile -Path $script:LocalConfigCachePath -TimeoutMs $script:LocalConfigTimeoutMs
    if ($localConfig -and (Apply-Config -Config $localConfig)) {
        $script:ConfigSource = "TEMP JSON: $script:LocalConfigCachePath"
        return $true
    }

    # 3. If TEMP is missing, empty, or invalid, rebuild it from script defaults.
    Set-DefaultConfiguration
    if (Save-LocalConfig) {
        $script:ConfigSource = "Built-in defaults: recreated TEMP JSON"
        return $true
    }

    $script:ConfigSource = "Built-in defaults: TEMP JSON create failed"
    if (-not $NoGui) {
        Show-AutoCloseNotice -Message "Network JSON not available and TEMP config could not be created.`nUsing default settings from code."
    }
    return $true
}
function Set-DefaultConfiguration {
    $script:vCenterServers=@(@{Name='PROD';Server='vcenter.example.invalid'})
    $script:DefaultSettings=@{
        vCenterServer='vcenter.example.invalid'
        EnableSCCMDeletion=$false
        EnableADDeletion=$false
        EnableDNSDeletion=$false
        EnableVMDeletion=$true
        EnableSQLArchive=$false
        EnableEmailNotification=$false
        SCCMServer='';SCCMSiteCode='';DNSServer='';DNSZone='';DomainController=''
        EmailNotification='';LogPath=(Join-Path $env:TEMP 'VMDeletionTool')
        EnableLogging=$true;ConfirmDeletion=$true;PowerOffTimeout=30;DeleteFromDisk=$false
        SQLServer='';SQLDatabase='';SQLTable='dbo.VMDeletionArchive'
        SQLIntegratedAuth=$true;SQLUsername='';SQLPassword=''
        SMTPServer='';SMTPPort=25;SMTPUseSSL=$false;SMTPFrom=''
        EmailSubjectPrefix='[VM Deletion]'
        ADSearchBase='';ADSearchBaseByEnvironment=@{PROD=''}
        GroupsCommon='';GroupsWindows='';GroupsLinux='';AdminAlertEmail=''
    }
}


function Save-LocalConfig {
    param([switch]$PersistSource)
    try {
        $settingsToSave=@{}
        foreach($key in $script:DefaultSettings.Keys){$settingsToSave[$key]=$script:DefaultSettings[$key]}
        # Runtime credentials are already decoded. Do not label plaintext as a DPAPI blob.
        $settingsToSave.Remove('SQLPasswordEncrypted')
        $config=@{vCenterServers=$script:vCenterServers;DefaultSettings=$settingsToSave;LastSaved=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss');SavedBy="$env:USERDOMAIN\$env:USERNAME"}
        $json=$config | ConvertTo-Json -Depth 10
        $destination=$script:LocalConfigCachePath
        if($PersistSource){
            if(-not [string]::IsNullOrWhiteSpace($script:ExplicitConfigPath)){$destination=$script:ExplicitConfigPath}
            elseif($script:ConfigSource -like 'Shared JSON:*'){$destination=$script:SharedConfigPath}
        }
        Write-DeletionConfigJson -Path $destination -Json $json
        if($destination -ne $script:LocalConfigCachePath){Write-DeletionConfigJson -Path $script:LocalConfigCachePath -Json $json}
        return $true
    }
    catch {
        Write-Warning "Configuration was not fully saved: $($_.Exception.Message)"
        return $false
    }
}

function Show-ResetConfigDialog {
    #Write-Host "Network path: $script:SharedConfigPath" -ForegroundColor Cyan
    #Write-Host "Network exists: $(Test-Path $script:SharedConfigPath)" -ForegroundColor Cyan
    #Write-Host "Local path: $script:LocalConfigCachePath" -ForegroundColor Cyan
    #Write-Host "Local exists: $(Test-Path $script:LocalConfigCachePath)" -ForegroundColor Cyan

    $resetForm = New-Object System.Windows.Forms.Form
    $resetForm.Text = "Reset Configuration"
    $resetForm.Size = New-Object System.Drawing.Size(450, 220)
    $resetForm.StartPosition = "CenterParent"
    $resetForm.FormBorderStyle = "FixedDialog"
    $resetForm.MaximizeBox = $false
    $resetForm.MinimizeBox = $false
    $resetForm.BackColor = [System.Drawing.Color]::FromArgb(54, 54, 68)
    $resetForm.TopMost = $true
    
    $lblQuestion = New-Object System.Windows.Forms.Label
    $lblQuestion.Location = New-Object System.Drawing.Point(20, 20)
$lblQuestion.Size = New-Object System.Drawing.Size(380, 30)
    $lblQuestion.Text = "Select configuration source to reset from:"
    $lblQuestion.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $lblQuestion.ForeColor = [System.Drawing.Color]::FromArgb(238, 240, 244)
    $resetForm.Controls.Add($lblQuestion)
    
    $btnNetwork = New-Object System.Windows.Forms.Button
    $btnNetwork.Location = New-Object System.Drawing.Point(20, 60)
    $btnNetwork.Size = New-Object System.Drawing.Size(180, 50)
    $btnNetwork.Text = "Network (Original)"
    $btnNetwork.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnNetwork.FlatStyle = "Flat"
    $btnNetwork.BackColor = [System.Drawing.Color]::FromArgb(45, 135, 183)
    $btnNetwork.ForeColor = [System.Drawing.Color]::White
    $btnNetwork.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnNetwork.Enabled = [bool](Wait-JsonFileAvailable -Path $script:SharedConfigPath -TimeoutMs $script:SharedConfigTimeoutMs)
    $btnNetwork.Add_Click({
        $resetForm.Tag = "Network"
        $resetForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $resetForm.Close()
    })
    $resetForm.Controls.Add($btnNetwork)
    
    $btnLocal = New-Object System.Windows.Forms.Button
    $btnLocal.Location = New-Object System.Drawing.Point(210, 60)
    $btnLocal.Size = New-Object System.Drawing.Size(180, 50)
    $btnLocal.Text = "Local (My Settings)"
    $btnLocal.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnLocal.FlatStyle = "Flat"
    $btnLocal.BackColor = [System.Drawing.Color]::FromArgb(62, 178, 122)
    $btnLocal.ForeColor = [System.Drawing.Color]::White
    $btnLocal.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnLocal.Enabled = [bool](Wait-JsonFileAvailable -Path $script:LocalConfigCachePath)
    $btnLocal.Add_Click({
        $resetForm.Tag = "Local"
        $resetForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $resetForm.Close()
    })
    $resetForm.Controls.Add($btnLocal)
    
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Location = New-Object System.Drawing.Point(20, 120)
$lblNote.Size = New-Object System.Drawing.Size(420, 20)
    $lblNote.Text = "Network = Original settings | Local = Your saved settings"
    $lblNote.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblNote.ForeColor = [System.Drawing.Color]::FromArgb(188, 192, 202)
    $resetForm.Controls.Add($lblNote)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(160, 145)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Text = "Cancel"
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.Add_Click({
        $resetForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $resetForm.Close()
    })
    $resetForm.Controls.Add($btnCancel)
    
    $result = $resetForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $resetForm.Tag
    }
    return $null
}

function Reset-Configuration {
    if($script:ExplicitConfigPath){
        $config=Load-ConfigFromFile -Path $script:ExplicitConfigPath
        return $(if($config){Apply-Config -Config $config}else{$false})
    }
    $source = Show-ResetConfigDialog
    
    if (-not $source) { return $false }
    
    $configLoaded = $false
    
    if ($source -eq "Network") {
        $networkConfig = Load-ConfigFromFile -Path $script:SharedConfigPath -TimeoutMs $script:SharedConfigTimeoutMs
        if ($networkConfig) {
            $configLoaded = Apply-Config -Config $networkConfig
            if ($configLoaded) {
                Save-LocalConfig | Out-Null
            }
        }
    }
    else {
        $localConfig = Load-ConfigFromFile -Path $script:LocalConfigCachePath
        if ($localConfig) {
            $configLoaded = Apply-Config -Config $localConfig
        }
    }
    
    if ($configLoaded) {
        [System.Windows.Forms.MessageBox]::Show(
            "Configuration reset from $source successfully!",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return $true
    }
    else {
        Write-Log "Failed to reset configuration from $source." -Level "ERROR"
        return $false
    }
}

function Refresh-SettingsForm {
    try {
        # SCCM Settings
        if ($txtSettingsSCCM) { $txtSettingsSCCM.Text = $script:DefaultSettings.SCCMServer }
        if ($txtSettingsSiteCode) { $txtSettingsSiteCode.Text = $script:DefaultSettings.SCCMSiteCode }
        
        # DNS Settings
        if ($txtSettingsDNS) { $txtSettingsDNS.Text = $script:DefaultSettings.DNSServer }
        if ($txtSettingsDNSZone) { $txtSettingsDNSZone.Text = $script:DefaultSettings.DNSZone }
        
        # AD Settings
        if ($txtSettingsDC) { $txtSettingsDC.Text = $script:DefaultSettings.DomainController }
        if ($txtSettingsEmail) { $txtSettingsEmail.Text = $script:DefaultSettings.EmailNotification }
        
        # General Options
        if ($chkEnableLogging) { $chkEnableLogging.Checked = $script:DefaultSettings.EnableLogging }
        if ($chkConfirmDeletion) { $chkConfirmDeletion.Checked = $script:DefaultSettings.ConfirmDeletion }
        if ($chkDeleteFromDisk) { $chkDeleteFromDisk.Checked = $script:DefaultSettings.DeleteFromDisk }
        if ($numTimeout) { $numTimeout.Value = $script:DefaultSettings.PowerOffTimeout }
        if ($txtLogPath) { $txtLogPath.Text = $script:DefaultSettings.LogPath }
        
        # Email Settings
        if ($chkEnableEmail) { $chkEnableEmail.Checked = $script:DefaultSettings.EnableEmailNotification }
        if ($txtSMTPServer) { $txtSMTPServer.Text = $script:DefaultSettings.SMTPServer }
        if ($numSMTPPort) { $numSMTPPort.Value = $script:DefaultSettings.SMTPPort }
        if ($chkSMTPSSL) { $chkSMTPSSL.Checked = $script:DefaultSettings.SMTPUseSSL }
        if ($txtSMTPFrom) { $txtSMTPFrom.Text = $script:DefaultSettings.SMTPFrom }
        
        # Email Groups - Common
        if ($script:lstGroupsCommon) {
            $script:lstGroupsCommon.Items.Clear()
            if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsCommon)) {
                $groups = $script:DefaultSettings.GroupsCommon.Split(',')
                foreach ($group in $groups) {
                    $g = $group.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($g)) {
                        [void]$script:lstGroupsCommon.Items.Add($g)
                    }
                }
            }
        }
        
        # Email Groups - Windows
        if ($script:lstGroupsWindows) {
            $script:lstGroupsWindows.Items.Clear()
            if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsWindows)) {
                $groups = $script:DefaultSettings.GroupsWindows.Split(',')
                foreach ($group in $groups) {
                    $g = $group.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($g)) {
                        [void]$script:lstGroupsWindows.Items.Add($g)
                    }
                }
            }
        }
        
        # Email Groups - Linux
        if ($script:lstGroupsLinux) {
            $script:lstGroupsLinux.Items.Clear()
            if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsLinux)) {
                $groups = $script:DefaultSettings.GroupsLinux.Split(',')
                foreach ($group in $groups) {
                    $g = $group.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($g)) {
                        [void]$script:lstGroupsLinux.Items.Add($g)
                    }
                }
            }
        }
        
        if($script:PersistentStepControls){
            foreach($definition in (Get-DeletionStepDefinitions)){$script:PersistentStepControls[$definition.Name].Checked=Test-DeletionStepEnabled -Step $definition.Name}
            Sync-DeletionStepControls -ResetSelection
        }
        # Status
        if ($lblSettingsStatus) {
            $lblSettingsStatus.Text = "Settings refreshed from $($script:ConfigSource)"
            $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        }
        
        #Write-Host "Settings form refreshed successfully" -ForegroundColor Green
    }
    catch {
        #Write-Host "Error refreshing form: $($_.Exception.Message)" -ForegroundColor Red
        if ($lblSettingsStatus) {
            $lblSettingsStatus.Text = "Error refreshing settings"
            $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
        }
    }
}


# ============================================
# Initialize Configuration on Startup
# ============================================
if(-not $Help){Initialize-Configuration | Out-Null}

#endregion

#region Email Functions
function New-DeletionSqlConnection {
    param([string]$ConnectionString)
    return (New-Object System.Data.SqlClient.SqlConnection($ConnectionString))
}

function Resolve-DeletionMailRecipients {
    param([string[]]$Values,[switch]$UsersOnly,[switch]$UsersFirst)
    $found=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach($value in $Values){
        foreach($entry in ([string]$value -split '[,;\r\n]+')){
            $item=$entry.Trim();if(-not $item){continue}
            if($item -match '@'){
                try{[void]$found.Add(([Net.Mail.MailAddress]::new($item)).Address)}catch{Write-Log "Invalid recipient address: $item" -Level 'WARNING'}
                continue
            }
            $identity=($item -split '\\')[-1]
            $emails=@()
            if($UsersOnly -or $UsersFirst){
                $user=Get-ADUserEmail -SamAccountName $identity -DomainController $script:DefaultSettings.DomainController
                if($user.Success -and $user.Email){$emails=@($user.Email)}
            }
            if(-not $emails.Count -and -not $UsersOnly){$emails=@(Get-ADGroupMemberEmails -GroupName $identity -DomainController $script:DefaultSettings.DomainController)}
            if(-not $emails.Count -and -not $UsersOnly -and -not $UsersFirst){
                $user=Get-ADUserEmail -SamAccountName $identity -DomainController $script:DefaultSettings.DomainController
                if($user.Success -and $user.Email){$emails=@($user.Email)}
            }
            foreach($email in $emails){try{[void]$found.Add(([Net.Mail.MailAddress]::new([string]$email)).Address)}catch{}}
            if(-not $emails.Count){Write-Log "No email address resolved for '$item'." -Level 'WARNING'}
        }
    }
    return @($found)
}

function Get-DeletionSummaryRecipients {
    param([array]$Records)
    $values=@()
    foreach($record in $Records){
        $metadata=$script:allVMsData|Where-Object{$_.Name -eq $record.VMName}|Select-Object -First 1
        $template=([bool]$record.IsTemplate -or ($metadata -and [bool]$metadata.IsTemplate))
        if($template){$values+=$script:DefaultSettings.GroupsWindows;continue}
        $values+=$script:DefaultSettings.GroupsCommon
        if($record.OSType -eq 'Linux'){
            # Preserve the existing special-cluster recipient rules from the JSON-backed workflow.
            if([string]$record.ClusterName -match 'LINUXGPU|LINUXHANA|LINUXNOORA|LINUXORASYN'){
                $values+=$script:DefaultSettings.GroupsWindows;$values+=$script:DefaultSettings.GroupsLinux
            }
        }
        else{$values+=$script:DefaultSettings.GroupsWindows}
    }
    return @(Resolve-DeletionMailRecipients -Values @($values|Where-Object{$_}|Select-Object -Unique))
}

function Get-DeletionAnnotationValues {
    param($Entity)
    $values=@{}
    if(-not $Entity){return $values}
    try{
        foreach($annotation in @(Get-Annotation -Entity $Entity -ErrorAction Stop)){
            $name=[string]$annotation.Name
            if(-not $name -and $annotation.CustomAttribute){$name=[string]$annotation.CustomAttribute.Name}
            if($name -and -not [string]::IsNullOrWhiteSpace([string]$annotation.Value)){$values[$name]=[string]$annotation.Value}
        }
    }
    catch{Write-Log "Could not read owner annotations for '$($Entity.Name)': $($_.Exception.Message)" -Level 'WARNING'}
    return $values
}

function Get-DeletionOwnerMetadata {
    param($VM)
    $annotations=Get-DeletionAnnotationValues -Entity $VM
    $serverOwner='';$folderOwner=''
    foreach($name in @('SRVOWNER','srvOWN1','ServerOwner')){if($annotations[$name]){$serverOwner=$annotations[$name];break}}
    foreach($name in @('FOLDEROWNER','FolderOwner')){if($annotations[$name]){$folderOwner=$annotations[$name];break}}
    if(-not $serverOwner -and [string]$VM.Notes -match '(?im)^\s*(?:srvOWN1|SRVOWNER|ServerOwner)[:=\s]+([^\r\n]+)'){$serverOwner=$Matches[1].Trim()}
    $folder=$VM.Folder
    $visited=New-Object 'Collections.Generic.HashSet[string]'
    for($depth=0;$folder -and -not $folderOwner -and $depth -lt 32;$depth++){
        $key=$(if($folder.Id){[string]$folder.Id}else{[string]$folder.Name})
        if(-not $visited.Add($key)){break}
        $folderAnnotations=Get-DeletionAnnotationValues -Entity $folder
        foreach($name in @('FOLDEROWNER','FolderOwner','srvOWN1','SRVOWNER')){if($folderAnnotations[$name]){$folderOwner=$folderAnnotations[$name];break}}
        $folder=$folder.Parent
    }
    return @{ServerOwner=$serverOwner;FolderOwner=$folderOwner}
}

function Show-WizardDeletionConfirmation {
    param([object[]]$SelectedVMs,[string]$Reason)
    $templateCount=@($SelectedVMs|Where-Object{$name=$_.Name;@($script:allVMsData|Where-Object{$_.Name -eq $name -and $_.IsTemplate}).Count -gt 0}).Count
    $vmCount=$SelectedVMs.Count-$templateCount
    $xaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Confirm deletion" Width="560" Height="510" MinWidth="440" MinHeight="380"
        WindowStartupLocation="CenterOwner" Background="#0C1730" Foreground="#E8F0FF" FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True">
  <Window.Resources>
    <Style x:Key="ConfirmButton" TargetType="Button">
      <Setter Property="Foreground" Value="#E8F0FF"/><Setter Property="Background" Value="#12234A"/>
      <Setter Property="BorderBrush" Value="#2B4577"/><Setter Property="Cursor" Value="Hand"/><Setter Property="MinWidth" Value="112"/>
      <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
        <Border x:Name="edge" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="10" Padding="18,11">
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="edge" Property="Opacity" Value="0.85"/></Trigger>
          <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="edge" Property="BorderBrush" Value="#62E4F4"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate></Setter.Value></Setter>
    </Style>
  </Window.Resources>
  <Grid Margin="24">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <StackPanel Margin="0,0,0,20">
      <TextBlock Text="Confirm deletion" FontFamily="Segoe UI Light" FontSize="28"/>
      <TextBlock Text="Review the selected machines before continuing." Foreground="#A7B9D9" Margin="0,7,0,0" TextWrapping="Wrap"/>
    </StackPanel>
    <Border Grid.Row="1" Background="#12234A" BorderBrush="#2B4577" BorderThickness="1" CornerRadius="16" Padding="18">
      <Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        <TextBlock x:Name="ItemTypeText" Foreground="#62E4F4" FontSize="15" FontWeight="SemiBold" Margin="0,0,0,12"/>
        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
          <TextBlock x:Name="VMNamesText" FontFamily="Consolas" Foreground="#E8F0FF" FontSize="13"/>
        </ScrollViewer>
      </Grid>
    </Border>
    <StackPanel Grid.Row="2" Margin="0,16,0,18">
      <TextBlock x:Name="confirmReason" Foreground="#A7B9D9" TextWrapping="Wrap"/>
      <TextBlock Text="Deletion cannot be undone. Confirm only after checking the list." Foreground="#F0647A" Margin="0,8,0,0" TextWrapping="Wrap"/>
    </StackPanel>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="CancelButton" Content="Cancel" IsCancel="True" IsDefault="True" Style="{StaticResource ConfirmButton}" Margin="0,0,10,0"/>
      <Button x:Name="DeleteButton" Content="Delete selected" Style="{StaticResource ConfirmButton}" Background="#F0647A" BorderBrush="#F0647A" Foreground="#0C1730"/>
    </StackPanel>
  </Grid>
</Window>
'@
    $dialog=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader ([xml]$xaml)))
    if($script:WizardWindow -and $script:WizardWindow.IsVisible){$dialog.Owner=$script:WizardWindow}
    $dialog.MaxHeight=[Windows.SystemParameters]::WorkArea.Height
    $dialog.Height=[Math]::Min(510,$dialog.MaxHeight)
    $dialog.FindName('ItemTypeText').Text="$vmCount VM(s) · $templateCount template(s)"
    $dialog.FindName('VMNamesText').Text=(@($SelectedVMs|ForEach-Object{$_.Name}) -join "`r`n")
    $dialog.FindName('confirmReason').Text="Reason: $Reason"
    $dialog.FindName('DeleteButton').Add_Click({$script:WizardConfirmDialog.DialogResult=$true})
    $dialog.FindName('CancelButton').Add_Click({$script:WizardConfirmDialog.DialogResult=$false})
    $dialog.Add_SourceInitialized({param($sender,$eventArgs) Set-WizardAdvancedCaption -Window $sender})
    $script:WizardConfirmDialog=$dialog
    try{return ($dialog.ShowDialog() -eq $true)}
    finally{$script:WizardConfirmDialog=$null}
}

function Get-WizardSqlSearchFilter {
    param([System.Data.DataTable]$Data,[string]$Text)
    if([string]::IsNullOrWhiteSpace($Text)){return ''}
    $literal=New-Object Text.StringBuilder
    foreach($character in $Text.ToCharArray()){
        $part=switch([string]$character){"'"{"''"};'['{'[[]'};']'{'[]]'};'*'{'[*]'};'%'{'[%]'};default{[string]$character}}
        [void]$literal.Append($part)
    }
    $clauses=@(foreach($column in $Data.Columns){
        if($column.DataType -eq [byte[]]){continue}
        $name=$column.ColumnName.Replace('\','\\').Replace(']','\]')
        "ISNULL(CONVERT([$name], 'System.String'), '') LIKE '%$($literal.ToString())%'"
    })
    if(-not $clauses.Count){return '1=0'}
    return '('+($clauses -join ' OR ')+')'
}

function Update-WizardSqlSearch {
    if(-not $script:AdvancedUi -or -not $script:AdvancedUi.aSqlResults.Tag -or $null -eq $script:SqlTableView){return}
    $snapshot=$script:AdvancedUi.aSqlResults.Tag
    try{
        $script:SqlTableView.RowFilter=Get-WizardSqlSearchFilter -Data $snapshot.Data -Text $script:AdvancedUi.tableSearch.Text
        $script:AdvancedUi.tableSearch.ToolTip='Search the loaded table without querying SQL again.'
        $script:AdvancedUi.tableCount.Text="$($script:SqlTableView.Count) of $($snapshot.Data.Rows.Count) records · $($snapshot.Data.Columns.Count) columns · loaded $($snapshot.LoadedAt.ToString('HH:mm:ss'))"
    }
    catch{
        $script:SqlTableView.RowFilter=''
        $script:AdvancedUi.tableSearch.ToolTip="Search could not be applied: $($_.Exception.Message)"
        $script:AdvancedUi.tableCount.Text="Search could not be applied. Showing all $($snapshot.Data.Rows.Count) records."
    }
    $script:AdvancedUi.tableEmpty.Text=$(if($script:AdvancedUi.tableSearch.Text){'No records match your search.'}else{'The SQL table is empty.'})
    $script:AdvancedUi.tableEmpty.Visibility=$(if($script:SqlTableView.Count){'Collapsed'}else{'Visible'})
    $script:AdvancedUi.tableClearSearch.IsEnabled=($script:AdvancedUi.tableSearch.IsEnabled -and $script:AdvancedUi.tableSearch.Text.Length -gt 0)
}

function Get-WizardSqlVisibleSnapshot {
    $snapshot=$script:AdvancedUi.aSqlResults.Tag
    if(-not $snapshot){return $null}
    if($null -eq $script:SqlTableView -or -not $script:SqlTableView.RowFilter){return $snapshot}
    return [pscustomobject]@{Data=$script:SqlTableView.ToTable();Server=$snapshot.Server;Database=$snapshot.Database;Table=$snapshot.Table;LoadedAt=$snapshot.LoadedAt}
}

function New-VmDeletionReportHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [string]$InitiatedBy = $env:USERNAME,
        [string]$VCenter = '',
        [string]$Reason = '',
        [datetime]$Date = (Get-Date)
    )

    $e = { param($t) [System.Net.WebUtility]::HtmlEncode([string]$t) }
    $font = "'Segoe UI',Arial,Helvetica,sans-serif"
    $steps = @(
        @{ Key = 'SCCM';     Label = 'SCCM' }
        @{ Key = 'AD';       Label = 'AD' }
        @{ Key = 'DNS';      Label = 'DNS' }
        @{ Key = 'VMware';   Label = 'vCenter' }
        @{ Key = 'Archived'; Label = 'Archive' }
    )

    # failed = has an error text, or any step explicitly $false
    $failed = @($Results | Where-Object {
        $r = $_
        [bool]$r.Error -or (@($steps | Where-Object { $r.($_.Key) -eq $false }).Count -gt 0)
    })
    $skipped = @($Results | Where-Object {
        $r = $_
        $failed -notcontains $r -and ($r.Skipped -or @($steps | Where-Object { $null -ne $r.($_.Key) }).Count -eq 0)
    })
    $ok = @($Results | Where-Object { $failed -notcontains $_ -and $skipped -notcontains $_ })

    $statusText  = if ($failed.Count) { 'COMPLETED WITH ERRORS' } elseif ($skipped.Count -eq $Results.Count) { 'NO ACTIONS PERFORMED' } else { 'COMPLETED' }
    $statusColor = if ($failed.Count) { '#F0647A' } elseif ($skipped.Count -eq $Results.Count) { '#8CA2C9' } else { '#4FD8A0' }
    $preheader   = "{0} succeeded, {1} failed, {2} skipped" -f $ok.Count, $failed.Count, $skipped.Count

    function Get-StepCell($value, $bg, $description) {
        if     ($value -eq $true)  { $t = '&#10003;'; $c = '#4FD8A0'; $w = 'bold' }
        elseif ($value -eq $false) { $t = '&#10005;'; $c = '#F0647A'; $w = 'bold' }
        else                       { $t = '&#8211;';  $c = '#5B7099'; $w = 'normal' }
        "<td align=`"center`" bgcolor=`"$bg`" title=`"$(& $e $description)`" style=`"padding:13px 4px;border-bottom:1px solid #1E3463;color:$c;font-size:16px;font-weight:$w;`">$t</td>"
    }

    function Get-VmTable($items, $showErrors) {
        $head = ($steps | ForEach-Object {
            "<td align=`"center`" width=`"76`" style=`"padding:10px 4px;color:#8CA2C9;font-size:11px;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid #2B4577;`">$($_.Label)</td>"
        }) -join ''
        $rows = ''
        $i = 0
        foreach ($r in $items) {
            $i++
            $bg = if ($i % 2) { '#12234A' } else { '#0F1E40' }
            $cells = ($steps | ForEach-Object { Get-StepCell $r.($_.Key) $bg $(if ($r.StepDescriptions) { $r.StepDescriptions[$_.Key] } else { $null }) }) -join ''
            $rows += "<tr><td bgcolor=`"$bg`" style=`"padding:13px 14px;border-bottom:1px solid #1E3463;color:#E8F0FF;font-size:14px;font-weight:600;`">$(& $e $r.Name)</td>" +
                     "<td bgcolor=`"$bg`" style=`"padding:13px 4px;border-bottom:1px solid #1E3463;color:#B9C9E8;font-size:13px;`">$(& $e $r.OS)</td>$cells</tr>"
            if ($showErrors -and $r.Error) {
                $rows += "<tr><td colspan=`"7`" bgcolor=`"#2A1B33`" style=`"padding:10px 14px;border-bottom:1px solid #1E3463;border-left:3px solid #F0647A;color:#FF9AA8;font-size:12px;`">$(& $e $r.Error)</td></tr>"
            }
        }
        @"
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid #2B4577;border-collapse:separate;border-radius:10px;">
<tr>
<td width="150" style="padding:10px 14px;color:#8CA2C9;font-size:11px;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid #2B4577;">VM name</td>
<td width="86" style="padding:10px 4px;color:#8CA2C9;font-size:11px;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid #2B4577;">System</td>
$head
</tr>
$rows
</table>
"@
    }

    function Get-Section($title, $color, $items, $showErrors) {
        if ($items.Count -eq 0) { return '' }
        @"
<tr><td style="padding:30px 32px 12px 32px;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>
<td width="10" height="10" bgcolor="$color" style="width:10px;height:10px;font-size:0;line-height:0;border-radius:5px;">&nbsp;</td>
<td style="padding-left:10px;color:$color;font-size:17px;font-weight:600;font-family:$font;">$title <span style="color:#8CA2C9;font-weight:normal;">($($items.Count))</span></td>
</tr></table>
</td></tr>
<tr><td style="padding:0 32px;font-family:$font;">$(Get-VmTable $items $showErrors)</td></tr>
"@
    }

    $successSection = Get-Section 'Successful VMs' '#4FD8A0' $ok $false
    $failedSection  = Get-Section 'Failed VMs'     '#F0647A' $failed $true
    $skippedSection = Get-Section 'Skipped VMs' '#8CA2C9' $skipped $false

    function Get-MetaCell($label, $value) {
        "<td width=`"50%`" valign=`"top`" style=`"padding:0 0 18px 0;`"><div style=`"color:#8CA2C9;font-size:11px;letter-spacing:1px;text-transform:uppercase;`">$label</div><div style=`"color:#E8F0FF;font-size:15px;padding-top:4px;`">$(& $e $value)</div></td>"
    }

    $dateText = $Date.ToString('dd/MM/yyyy HH:mm')
    $vcText   = if ($VCenter) { $VCenter } else { '-' }
    $rsText   = if ($Reason)  { $Reason }  else { '-' }

    @"
<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="color-scheme" content="dark">
<meta name="supported-color-schemes" content="dark">
<title>VM Deletion Report</title>
</head>
<body style="margin:0;padding:0;background:#070F20;" bgcolor="#070F20">
<div style="display:none;max-height:0;overflow:hidden;opacity:0;color:#070F20;">$preheader</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#070F20">
<tr><td align="center" style="padding:28px 12px;">

<table role="presentation" width="680" cellpadding="0" cellspacing="0" border="0" bgcolor="#0C1730" style="width:680px;max-width:680px;font-family:$font;border:1px solid #2B4577;border-collapse:separate;border-radius:14px;">

<tr><td height="4" bgcolor="#62E4F4" style="height:4px;font-size:0;line-height:0;border-radius:14px 14px 0 0;">&nbsp;</td></tr>

<tr><td style="padding:30px 32px 6px 32px;">
<div style="color:#62E4F4;font-size:12px;letter-spacing:2px;text-transform:uppercase;">VM lifecycle &middot; automation</div>
<div style="color:#E8F0FF;font-size:30px;font-weight:300;padding-top:8px;font-family:'Segoe UI Light','Segoe UI',Arial,sans-serif;">VM Deletion Report</div>
</td></tr>
<tr><td style="padding:10px 32px 24px 32px;">
<span style="display:inline-block;padding:5px 14px;border:1px solid $statusColor;border-radius:999px;color:$statusColor;font-size:11px;letter-spacing:1.5px;font-weight:600;">$statusText</span>
</td></tr>

<tr><td style="padding:0 32px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#12234A" style="border:1px solid #2B4577;border-collapse:separate;border-radius:10px;">
<tr><td style="padding:20px 22px 2px 22px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>$(Get-MetaCell 'Initiated by' $InitiatedBy)$(Get-MetaCell 'Date' $dateText)</tr>
<tr>$(Get-MetaCell 'vCenter' $vcText)$(Get-MetaCell 'Reason' $rsText)</tr>
</table>
</td></tr></table>
</td></tr>

<tr><td style="padding:16px 32px 0 32px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#12234A" style="border:1px solid #2B4577;border-collapse:separate;border-radius:10px;">
<tr>
<td width="33%" align="center" style="padding:18px 0;border-right:1px solid #2B4577;"><div style="color:#E8F0FF;font-size:34px;font-weight:300;">$($Results.Count)</div><div style="color:#8CA2C9;font-size:11px;letter-spacing:1px;text-transform:uppercase;">Total VMs</div></td>
<td width="34%" align="center" style="padding:18px 0;border-right:1px solid #2B4577;"><div style="color:#4FD8A0;font-size:34px;font-weight:300;">$($ok.Count)</div><div style="color:#8CA2C9;font-size:11px;letter-spacing:1px;text-transform:uppercase;">Succeeded</div></td>
<td width="33%" align="center" style="padding:18px 0;"><div style="color:#F0647A;font-size:34px;font-weight:300;">$($failed.Count)</div><div style="color:#8CA2C9;font-size:11px;letter-spacing:1px;text-transform:uppercase;">Failed</div></td>
</tr></table>
</td></tr>

$failedSection
$successSection
$skippedSection

<tr><td style="padding:30px 32px 26px 32px;">
<div style="border-top:1px solid #2B4577;padding-top:16px;color:#6F85AD;font-size:11px;line-height:1.6;">
&#10003; done / already absent &nbsp;&middot;&nbsp; &#10005; failed &nbsp;&middot;&nbsp; &#8211; skipped<br>
Generated automatically by the VM Deletion Tool on $(& $e $env:COMPUTERNAME). Please do not reply to this message.
</div>
</td></tr>

</table>
</td></tr></table>
</body></html>
"@
}

function New-VmDeletionReportStepStates {
    # Reporting snapshot only; operational checkboxes and deletion behavior are unchanged.
    $selected = @{SCCM=$chkSCCM.Checked; AD=$chkAD.Checked; DNS=$chkDNS.Checked; VMware=$chkVM.Checked; Archived=$chkArchive.Checked}
    $steps = @{}
    foreach ($key in $selected.Keys) {
        $steps[$key] = @{Requested=[bool]$selected[$key]; Status=$(if ($selected[$key]) {'Not run'} else {'Skipped'}); Detail=''}
    }
    return $steps
}

function Update-VmDeletionReportStep {
    param([int]$TaskIndex, [string]$Status, [string]$Message)
    $record = $script:CurrentVmDeletionReportRecord
    if (-not $record -or -not $record.ReportSteps -or $TaskIndex -lt 1 -or $TaskIndex -gt 5 -or $Status -eq 'Pending') { return }
    $key = @('Connection','SCCM','AD','DNS','VMware','Archived')[$TaskIndex]
    $state = $record.ReportSteps[$key]
    $state.Status = $Status
    $state.Detail = $Message
    if ($Status -eq 'Success') {
        $flag = @{SCCM='SCCMDeleted'; AD='ADDeleted'; DNS='DNSDeleted'}[$key]
        if ($flag -and -not $record[$flag]) { $state.Status = 'Not found' }
    }

}

function Get-VmDeletionReportVCenter {
    param([array]$Records=@())
    $names = @($Records | ForEach-Object { $_.VCenter } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if (-not $names.Count) {
        $names = @($global:DefaultVIServer | ForEach-Object { $_.Name } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    }
    if (-not $names.Count -and $script:DefaultSettings.vCenterServer) { $names = @($script:DefaultSettings.vCenterServer) }
    if (-not $names.Count) { return 'Not recorded' }
    return ($names -join ', ')
}

function ConvertTo-VmDeletionReportRows {
    param([array]$SuccessList=@(), [array]$FailedList=@(), [array]$SkippedList=@())
    $steps = @{SCCM='SCCMDeleted'; AD='ADDeleted'; DNS='DNSDeleted'; VMware='VMDeleted'; Archived='Archived'}
    foreach ($record in @($SuccessList) + @($FailedList) + @($SkippedList)) {
        if (-not $record -or [string]::IsNullOrWhiteSpace([string]$record.VMName)) { continue }
        $errors = @($record.Errors | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        $row = @{Name=[string]$record.VMName; OS=[string]$record.OSType; Error=''; Skipped=$false; StepDescriptions=@{}}
        foreach ($key in $steps.Keys) {
            $state = $null
            if ($record.ReportSteps) { $state = $record.ReportSteps[$key] }
            if ($state) {
                switch ($state.Status) {
                    'Success' { $row[$key] = $true; $row.StepDescriptions[$key] = 'Completed' }
                    'Not found' { $row[$key] = $true; $row.StepDescriptions[$key] = 'No matching record existed (already absent)' }
                    'Skipped' { $row[$key] = $null; $row.StepDescriptions[$key] = $(if($state.Detail){$state.Detail}else{'Not selected for this run'}) }
                    default {
                        $row[$key] = $false
                        $detail = [string]$state.Detail
                        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = 'Requested action did not complete.' }
                        $row.StepDescriptions[$key] = $detail
                        $errors += "${key}: $detail"
                    }
                }
            }
            elseif ($record[$steps[$key]]) {
                $row[$key] = $true
                $row.StepDescriptions[$key] = 'Completed'
            }
            else {
                # Older callers do not distinguish absent records from skipped actions.
                $prefix = switch ($key) {'Archived' {'Archive'}; 'VMware' {'VM(?:ware|/Template Delete)?'}; default {[regex]::Escape($key)}}
                if (@($errors | Where-Object { $_ -match ('^' + $prefix + ':') }).Count) {
                    $row[$key] = $false
                    $row.StepDescriptions[$key] = 'Failed; see details below'
                }
                else { $row[$key] = $null; $row.StepDescriptions[$key] = 'Skipped or not recorded by the older caller' }
            }
        }
        if ($FailedList -contains $record -and -not $errors.Count) { $errors += 'The run recorded this machine as failed.' }
        $row.Error = @($errors | Select-Object -Unique) -join ' | '
        $row.Skipped = (-not $row.Error -and @($steps.Keys | Where-Object { $null -ne $row[$_] }).Count -eq 0)
        [pscustomobject]$row
    }
}

function Send-DeletionSummaryReport {
    param([array]$SuccessList, [array]$FailedList, [array]$SkippedList, [string]$Reason)
    if(-not $script:DefaultSettings.EnableEmailNotification){return}
    try {
        $results = @(ConvertTo-VmDeletionReportRows -SuccessList $SuccessList -FailedList $FailedList -SkippedList $SkippedList)
        if (-not $results.Count) { Write-Log 'No VMs to report - skipping email'; return }
        $records = @($SuccessList) + @($FailedList) + @($SkippedList)
        $recipients=@(Get-DeletionSummaryRecipients -Records $records)
        if(-not $recipients.Count){Write-Log 'No JSON team recipients resolved for the summary report.' -Level 'WARNING';return}
        $html = New-VmDeletionReportHtml -Results $results -InitiatedBy "$env:USERDOMAIN\$env:USERNAME" -VCenter (Get-VmDeletionReportVCenter -Records $records) -Reason $Reason
        $failed = @($results | Where-Object { $_.Error -or $_.SCCM -eq $false -or $_.AD -eq $false -or $_.DNS -eq $false -or $_.VMware -eq $false -or $_.Archived -eq $false })
        $skipped = @($results | Where-Object { $_.Skipped })
        $succeeded = $results.Count - $failed.Count - $skipped.Count
        $parts = @()
        if ($succeeded) { $parts += "$succeeded Success" }
        if ($failed.Count) { $parts += "$($failed.Count) Failed" }
        if ($skipped.Count) { $parts += "$($skipped.Count) Skipped" }
        $emailParams = @{
            From=$script:DefaultSettings.SMTPFrom; To=$recipients
            Subject="$($script:DefaultSettings.EmailSubjectPrefix) Summary: " + ($parts -join ', ')
            Body=$html; BodyAsHtml=$true; SmtpServer=$script:DefaultSettings.SMTPServer
            Port=$script:DefaultSettings.SMTPPort; Encoding=[Text.Encoding]::UTF8
        }
        if ($script:DefaultSettings.SMTPUseSSL) { $emailParams.UseSsl = $true }
        Send-MailMessage @emailParams -ErrorAction Stop
        Write-Log "Deletion summary report sent to the configured team recipients: $($recipients -join ', ')"
    }
    catch { Write-Log "Failed to send summary report: $($_.Exception.Message)" -Level 'ERROR' }
}


function Get-ADUserEmail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,
        [string]$DomainController
    )
    
    $result = @{
        Success = $false
        Email   = ""
        Message = ""
    }
    
    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            $result.Message = "ActiveDirectory module not installed"
            return $result
        }
        
        Import-Module ActiveDirectory -ErrorAction Stop
        
        $adParams = @{
            Identity    = $SamAccountName
            Properties  = @("mail", "EmailAddress")
            ErrorAction = 'Stop'
        }
        
        if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
            $adParams.Server = $DomainController
        }
        
        $user = Get-ADUser @adParams
        
        if ($user) {
            $email = if ($user.mail) { $user.mail } else { $user.EmailAddress }
            if ($email) {
                $result.Success = $true
                $result.Email = $email
                $result.Message = "Email found"
            }
            else {
                $result.Message = "User found but no email address"
            }
        }
    }
    catch {
        $result.Message = "Error: $($_.Exception.Message)"
    }
    
    return $result
}

function Get-ADGroupMemberEmails {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        [string]$DomainController
    )
    
    $emails = @()
    
    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            Write-Log "ActiveDirectory module not installed" -Level "WARNING"
            return $emails
        }
        
        Import-Module ActiveDirectory -ErrorAction Stop
        
        $adParams = @{
            Identity    = $GroupName
            Recursive   = $true
            ErrorAction = 'Stop'
        }
        
        if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
            $adParams.Server = $DomainController
        }
        
        $members = Get-ADGroupMember @adParams
        
        foreach ($member in $members) {
            if ($member.objectClass -eq "user") {
                $userParams = @{
                    Identity    = $member.SamAccountName
                    Properties  = @("mail", "EmailAddress")
                    ErrorAction = 'SilentlyContinue'
                }
                
                if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
                    $userParams.Server = $DomainController
                }
                
                $user = Get-ADUser @userParams
                
                if ($user) {
                    $email = if ($user.mail) { $user.mail } else { $user.EmailAddress }
                    if ($email -and $emails -notcontains $email) {
                        $emails += $email
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Error getting group members for $GroupName : $($_.Exception.Message)" -Level "WARNING"
    }
    
    return $emails
}

function Send-VMDeletionNotification {
    param([Parameter(Mandatory=$true)][hashtable]$VMDetails,[string]$Reason,[string]$OSType,[bool]$DeletionSuccess,$ReportRecord=$null)
    $result=@{Success=$false;Message='';SentTo=@()}
    if(-not $script:DefaultSettings.EnableEmailNotification){$result.Message='Email notifications disabled';return $result}
    try {
        $actor=@(Resolve-DeletionMailRecipients -Values @($env:USERNAME) -UsersOnly)
        $owner=@(Resolve-DeletionMailRecipients -Values @($VMDetails.ServerOwner) -UsersFirst)
        if(-not $owner.Count){$owner=@(Resolve-DeletionMailRecipients -Values @($VMDetails.FolderOwner) -UsersFirst)}
        $recipients=@(@($actor)+@($owner)|Where-Object{$_}|Sort-Object -Unique)
        if(-not $recipients.Count){throw 'No email address could be resolved for the deletion initiator or server/folder owner.'}
        if(-not $actor.Count){Write-Log 'The deletion initiator email could not be resolved; no team address was substituted.' -Level 'WARNING'}
        if(-not $owner.Count){Write-Log "No server/folder owner email was resolved for '$($VMDetails.VMName)'." -Level 'WARNING'}
        if($ReportRecord){$rows=@(ConvertTo-VmDeletionReportRows -SuccessList @($ReportRecord))}
        else{
            $rows=@([pscustomobject]@{Name=$VMDetails.VMName;OS=$OSType;SCCM=$null;AD=$null;DNS=$null;VMware=$DeletionSuccess;Archived=$null;Error=$(if($DeletionSuccess){''}else{'VM deletion did not complete.'})})
        }
        $vc=$VMDetails.VCenterName;if(-not $vc){$vc=$VMDetails.vCenter}
        $body=New-VmDeletionReportHtml -Results $rows -InitiatedBy "$env:USERDOMAIN\$env:USERNAME" -VCenter $vc -Reason $Reason
        $state=$(if($rows[0].Error){'Needs attention'}elseif($DeletionSuccess){'Deleted'}else{'Cleanup completed'})
        $mail=@{From=$script:DefaultSettings.SMTPFrom;To=$recipients;Subject="$($script:DefaultSettings.EmailSubjectPrefix) $($VMDetails.VMName) - $state";Body=$body;BodyAsHtml=$true;SmtpServer=$script:DefaultSettings.SMTPServer;Port=$script:DefaultSettings.SMTPPort;Encoding=[Text.Encoding]::UTF8;Priority='High'}
        if($script:DefaultSettings.SMTPUseSSL){$mail.UseSsl=$true}
        Send-MailMessage @mail -ErrorAction Stop
        $result.Success=$true;$result.SentTo=$recipients;$result.Message="Personal deletion report sent to $($recipients.Count) recipient(s)."
        Write-Log "Personal deletion report sent for '$($VMDetails.VMName)'."
    }
    catch{$result.Message="Email error: $($_.Exception.Message)";Write-Log $result.Message -Level 'ERROR'}
    return $result
}

function Send-FailureEmail {
    param(
        [string]$Subject,
        [string]$Body
    )

    try {
        $smtpServer = $script:DefaultSettings.SmtpServer
        $from       = $script:DefaultSettings.EmailNotification
        $to         = $script:DefaultSettings.AdminAlertEmail.Split(";")

        Send-MailMessage `
            -From $from `
            -To $to `
            -Subject $Subject `
            -Body $Body `
            -SmtpServer $smtpServer `
            -ErrorAction Stop

        Write-Log "Failure email sent successfully"
    }
    catch {
        Write-Log "FAILED sending mail: $($_.Exception.Message)" -Level "ERROR"
    }
}

#endregion Email Functions

# Main Tab Events

function Update-TaskStatus {
    param([int]$TaskIndex, [string]$Status, [string]$Message)
    Update-VmDeletionReportStep -TaskIndex $TaskIndex -Status $Status -Message $Message
    $sl = $statusLabels[$TaskIndex]
    $rl = $resultLabels[$TaskIndex]
    switch ($Status) {
        "Running" {
            $sl.Text = "RUN"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(20,184,166)
            $rl.Text = "In Progress..."
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(20,184,166)
        }
        "Success" {
            $sl.Text = "OK"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(34,197,94)
            $rl.Text = $Message
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(34,197,94)
        }
        "Failed" {
            $sl.Text = "X"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(225,29,72)
            $rl.Text = $Message
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(225,29,72)
        }
        "Skipped" {
            $sl.Text = "SKIP"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(234,179,8)
            $rl.Text = $Message
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(234,179,8)
        }
        "Pending" {
            $sl.Text = "WAIT"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(214, 218, 225)
            $rl.Text = "Pending"
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(214, 218, 225)
        }
    }
    if ($form) { $form.Refresh() }
    [System.Windows.Forms.Application]::DoEvents()
    if ($script:WizardWindow) {
        Sync-DeletionWizard
        $script:WizardWindow.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
    }
}

function Reset-AllStatus { 
    for ($i = 0; $i -lt $tasks.Count; $i++) { 
        Update-TaskStatus -TaskIndex $i -Status "Pending" -Message "Pending" 
    }
    $progressBar.Value = 0 
}

function Get-VMList {
    $script:LastAppliedVMFilterKey = ""
    $script:IsUpdatingList = $true
    
    # שמירת בחירות קיימות
    $selectedHash = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($item in $lstVMs.CheckedItems) {
        if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace($item.Text)) {
            [void]$selectedHash.Add($item.Text)
        }
    }
    # הוסף גם את הבחירות הקודמות מה-ArrayList
    if ($null -ne $script:SelectedVMNames) {
        foreach ($vmName in $script:SelectedVMNames) {
            if (-not [string]::IsNullOrWhiteSpace($vmName)) {
                [void]$selectedHash.Add($vmName)
            }
        }
    }
    
    $lstVMs.BeginUpdate()
    $lstVMs.Items.Clear()
    $script:allVMs = @()
    $script:allVMsData = @()
    $script:allTemplatesData = @()
    
    try {
        # שליפה מהירה של VMs – כולל PowerState וסטטוס Template
        $views = Get-View -ViewType VirtualMachine -Property Name,Config.GuestFullName,Config.GuestId,Runtime.PowerState,Config.Template -ErrorAction Stop
        
        if ($null -eq $views) {
            $lblVMCount.Text = "0 VMs"
            $lstVMs.EndUpdate()
            $script:IsUpdatingList = $false
            return
        }
        
        foreach ($vm in $views | Sort-Object Name) {
            # בדיקת null בסיסית
            if ($null -eq $vm -or [string]::IsNullOrWhiteSpace($vm.Name)) { 
                continue 
            }
            
            # בדוק אם זה Template (עם בדיקת null)
            $isTemplate = $false
            if ($null -ne $vm.Config) {
                $isTemplate = [bool]$vm.Config.Template
            }
            
            # דלג על Templates אם לא נבחרה האפשרות להציג אותם
            if ($isTemplate -and -not $script:IncludeTemplates) {
                continue
            }
            
            # קבל ערכים עם בדיקות null
            $osFullName = $null
            $guestId = $null
            if ($null -ne $vm.Config) {
                $osFullName = $vm.Config.GuestFullName
                $guestId = $vm.Config.GuestId
            }
            
            # קבע PowerState עם בדיקת null
            $powerState = "Unknown"
            if ($isTemplate) {
                $powerState = "Template"
            }
            elseif ($null -ne $vm.Runtime -and $null -ne $vm.Runtime.PowerState) {
                $powerState = $vm.Runtime.PowerState.ToString()
            }
            
            # קבע סוג OS
            $osType = "Other"
            $osFullNameStr = if ($osFullName) { $osFullName.ToString() } else { "" }
            $guestIdStr = if ($guestId) { $guestId.ToString() } else { "" }
            
            if ($osFullNameStr -match "Windows" -or $guestIdStr -match "windows") {
                $osType = "Windows"
            }
            elseif ($osFullNameStr -match "Linux|Ubuntu|CentOS|Red Hat|RHEL|Debian|SUSE|Fedora" -or
                    $guestIdStr -match "linux|ubuntu|centos|rhel|debian|suse") {
                $osType = "Linux"
            }
            
            # קבע OSFull
            $osFull = "Unknown"
            if (-not [string]::IsNullOrWhiteSpace($osFullNameStr)) {
                $osFull = $osFullNameStr
            }
            elseif (-not [string]::IsNullOrWhiteSpace($guestIdStr)) {
                $osFull = $guestIdStr
            }
            
            $vmData = @{
                Name       = $vm.Name
                OSType     = $osType
                OSFull     = $osFull
                PowerState = $powerState
                IsTemplate = $isTemplate
            }
            
            $script:allVMs += $vm.Name
            $script:allVMsData += $vmData
            
            if ($isTemplate) {
                $script:allTemplatesData += $vmData
            }
            
            $item = New-Object System.Windows.Forms.ListViewItem($vm.Name)
            [void]$item.SubItems.Add($powerState)
            [void]$item.SubItems.Add($osType)
            [void]$item.SubItems.Add($osFull)
            
            # צבע טקסט לפי OS או Template
            if ($isTemplate) {
                $item.ForeColor = [System.Drawing.Color]::FromArgb(196, 181, 253)  # Soft purple for templates
                $item.BackColor = [System.Drawing.Color]::FromArgb(96, 86, 132)  # Dark purple row
            }
            else {
                switch ($osType) {
                    "Windows" { $item.ForeColor = [System.Drawing.Color]::FromArgb(147, 197, 253) }
                    "Linux"   { $item.ForeColor = [System.Drawing.Color]::FromArgb(251, 191, 36) }
                    default   { $item.ForeColor = [System.Drawing.Color]::FromArgb(214, 218, 225) }
                }
                
                # צבע רקע לפי PowerState
                switch ($powerState) {
                    "poweredOff" { $item.BackColor = [System.Drawing.Color]::FromArgb(104, 66, 80) }
                    "suspended"  { $item.BackColor = [System.Drawing.Color]::FromArgb(110, 94, 60) }
                    default      { $item.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 78) }
                }
            }
            
            # שחזר בחירה
            if ($selectedHash.Contains($vm.Name)) {
                $item.Checked = $true
            }
            
            [void]$lstVMs.Items.Add($item)
        }
        
        # עדכן את ה-ArrayList
        $script:SelectedVMNames.Clear()
        foreach ($name in $selectedHash) {
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                [void]$script:SelectedVMNames.Add($name)
            }
        }
        
        $templateCount = @($script:allVMsData | Where-Object { $_.IsTemplate }).Count
        $vmCount = $lstVMs.Items.Count - $templateCount
        if ($script:IncludeTemplates -and $templateCount -gt 0) {
            $lblVMCount.Text = "$vmCount VMs + $templateCount Templates"
        } else {
            $lblVMCount.Text = "$($lstVMs.Items.Count) VMs"
        }
        $lblSelectedCount.Text = "$($script:SelectedVMNames.Count) selected"
    }
    catch {
        Write-Log "Failed to retrieve VM list: $($_.Exception.Message); Stack: $($_.ScriptStackTrace)" -Level "ERROR"
    }
    
    $lstVMs.EndUpdate()
    $script:IsUpdatingList = $false
}

function New-VMListViewItem {
    param(
        [hashtable]$VMData,
        [System.Collections.Generic.HashSet[string]]$SelectedNames
    )

    if ($null -eq $VMData -or [string]::IsNullOrWhiteSpace($VMData.Name)) { return $null }

    $powerState = if ($VMData.PowerState) { $VMData.PowerState } else { "Unknown" }
    $osType = if ($VMData.OSType) { $VMData.OSType } else { "Other" }
    $osFull = if ($VMData.OSFull) { $VMData.OSFull } else { "Unknown" }

    $item = New-Object System.Windows.Forms.ListViewItem($VMData.Name)
    [void]$item.SubItems.Add($powerState)
    [void]$item.SubItems.Add($osType)
    [void]$item.SubItems.Add($osFull)

    if ($VMData.IsTemplate) {
        $item.ForeColor = [System.Drawing.Color]::FromArgb(196, 181, 253)
        $item.BackColor = [System.Drawing.Color]::FromArgb(96, 86, 132)
    }
    else {
        switch ($osType) {
            "Windows" { $item.ForeColor = [System.Drawing.Color]::FromArgb(147, 197, 253) }
            "Linux"   { $item.ForeColor = [System.Drawing.Color]::FromArgb(251, 191, 36) }
            default   { $item.ForeColor = [System.Drawing.Color]::FromArgb(214, 218, 225) }
        }

        switch ($powerState) {
            "poweredOff" { $item.BackColor = [System.Drawing.Color]::FromArgb(104, 66, 80) }
            "suspended"  { $item.BackColor = [System.Drawing.Color]::FromArgb(110, 94, 60) }
            default      { $item.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 78) }
        }
    }

    if ($SelectedNames -and $SelectedNames.Contains($VMData.Name)) {
        $item.Checked = $true
    }

    return $item
}
function Filter-VMList {
    param(
        [string]$FilterText,
        [string]$OSFilter = "All"
    )

    if ($script:IsUpdatingList) { return }

    $filterKey = "{0}|{1}|{2}" -f $FilterText, $OSFilter, $script:IncludeTemplates
    if ($script:LastAppliedVMFilterKey -eq $filterKey) { return }

    $filterStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $script:LastAppliedVMFilterKey = $filterKey
    $script:IsUpdatingList = $true

    $selectedHash = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $lstVMs.CheckedItems) {
        if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace($item.Text)) {
            [void]$selectedHash.Add($item.Text)
        }
    }
    if ($null -ne $script:SelectedVMNames) {
        foreach ($vmName in $script:SelectedVMNames) {
            if (-not [string]::IsNullOrWhiteSpace($vmName)) {
                [void]$selectedHash.Add([string]$vmName)
            }
        }
    }

    $items = New-Object 'System.Collections.Generic.List[System.Windows.Forms.ListViewItem]'
    $filterValue = if ($null -eq $FilterText) { "" } else { $FilterText.Trim() }
    $useTextFilter = -not [string]::IsNullOrWhiteSpace($filterValue)
    $useOSFilter = -not [string]::IsNullOrWhiteSpace($OSFilter) -and $OSFilter -ne "All"

    foreach ($vmData in $script:allVMsData) {
        if ($null -eq $vmData -or [string]::IsNullOrWhiteSpace($vmData.Name)) { continue }
        if ($useTextFilter -and $vmData.Name.IndexOf($filterValue, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        if ($useOSFilter -and $vmData.OSType -ne $OSFilter) { continue }

        $item = New-VMListViewItem -VMData $vmData -SelectedNames $selectedHash
        if ($null -ne $item) { $items.Add($item) }
    }

    $lstVMs.BeginUpdate()
    try {
        $lstVMs.Items.Clear()
        if ($items.Count -gt 0) {
            $lstVMs.Items.AddRange($items.ToArray())
        }
    }
    finally {
        $lstVMs.EndUpdate()
        $script:IsUpdatingList = $false
    }

    $script:SelectedVMNames.Clear()
    foreach ($name in $selectedHash) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            [void]$script:SelectedVMNames.Add($name)
        }
    }

    $lblVMCount.Text = "$($lstVMs.Items.Count) of $(@($script:allVMsData).Count)"
    $lblSelectedCount.Text = "$($script:SelectedVMNames.Count) selected"

    $filterStopwatch.Stop()
    if ($filterStopwatch.ElapsedMilliseconds -gt 300) {
        Write-Log "VM filter completed in $($filterStopwatch.ElapsedMilliseconds) ms. Results=$($lstVMs.Items.Count), Filter='$filterValue', OS='$OSFilter'"
    }
}

function Schedule-VMFilter {
    $osFilter = "All"
    if ($null -ne $cboOSFilter.SelectedItem) {
        $osFilter = $cboOSFilter.SelectedItem.ToString()
    }

    if ($null -eq $script:VMFilterTimer) {
        Filter-VMList -FilterText $txtFilter.Text -OSFilter $osFilter
        return
    }

    $script:PendingVMFilterText = $txtFilter.Text
    $script:PendingVMOSFilter = $osFilter
    $script:VMFilterTimer.Stop()
    $script:VMFilterTimer.Start()
}

function Set-ControlsEnabled { 
    param([bool]$Enabled)
    $lstVMs.Enabled = $true
    $lstVMs.BackColor = $uiInput
    $lstVMs.ForeColor = $uiText
    $txtFilter.Enabled = $Enabled
    $cboOSFilter.Enabled = $Enabled
    $btnRefresh.Enabled = $Enabled
    $btnClearSelection.Enabled = $Enabled
    $btnLoadCSV.Enabled = $Enabled
    $btnDelete.Enabled = $Enabled
    $chkSCCM.Enabled = $Enabled
    $chkAD.Enabled = $Enabled
    $chkDNS.Enabled = $Enabled
    $chkVM.Enabled = $Enabled
    $chkArchive.Enabled = $Enabled
    $chkEmail.Enabled = $Enabled
    Sync-DeletionStepControls -Available $Enabled
    $txtReason.Enabled = $Enabled
}

function Get-ADSearchBaseForEnvironment {
    param([string]$EnvironmentName)

    $envKey = if ([string]::IsNullOrWhiteSpace($EnvironmentName)) {
        "PROD"
    }
    else {
        $EnvironmentName.Trim().ToUpper()
    }

    try {
        $map = $script:DefaultSettings.ADSearchBaseByEnvironment
        if ($map) {
            if ($map -is [hashtable] -and $map.ContainsKey($envKey)) {
                return $map[$envKey]
            }

            if ($map.PSObject.Properties.Name -contains $envKey) {
                return $map.$envKey
            }
        }
    }
    catch { }

    return $script:DefaultSettings.ADSearchBase
}

function Resolve-VCenterEnvironment {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $script:RunEnvironment }

    foreach ($vc in $script:vCenterServers) {
        if ($vc.Name -eq $Value -or $vc.Server -eq $Value) {
            return $vc.Name
        }
    }

    return $Value
}





function Save-Settings {
    # One serializer and destination policy for both configuration buttons.
    return (Save-LocalConfig -PersistSource)
}

function Load-Settings {
    $path=if($script:ExplicitConfigPath){$script:ExplicitConfigPath}else{$script:LocalConfigCachePath}
    $config=Load-ConfigFromFile -Path $path
    if(-not $config){return $false}
    return (Apply-Config -Config $config)
}

function Add-GuiLogEntry {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Component = "VMDeletionTool"
    )

    if ($null -eq $script:txtLogs) { return }

    try {
        $entry = "{0} [{1}] [{2}] {3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level.ToUpper(), $Component, $Message
        if ($script:txtLogs.InvokeRequired) {
            [void]$script:txtLogs.BeginInvoke([Action[string]]{
                param($line)
                $script:txtLogs.AppendText($line + [Environment]::NewLine + [Environment]::NewLine)
                $script:txtLogs.SelectionStart = $script:txtLogs.TextLength
                $script:txtLogs.ScrollToCaret()
            }, $entry)
        }
        else {
            $script:txtLogs.AppendText($entry + [Environment]::NewLine + [Environment]::NewLine)
            $script:txtLogs.SelectionStart = $script:txtLogs.TextLength
            $script:txtLogs.ScrollToCaret()
        }
    }
    catch { }
}

function Show-GuiErrorNotice {
    if ($script:IsShowingLogErrorNotice) { return }
    if ($null -eq $form -or -not $form.Visible) { return }

    try {
        $script:IsShowingLogErrorNotice = $true
        [System.Windows.Forms.MessageBox]::Show(
            "ERROR: Go to Logs to see it.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    catch { }
    finally {
        $script:IsShowingLogErrorNotice = $false
    }
}
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    Write-ExtendedLog -Message $Message -Level $Level
}

function Write-ExtendedLog {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Component = "VMDeletionTool",
        [switch]$SendEmail
    )
    
    # Log Directory Management
    $logDir = $script:DefaultSettings.LogPath
    $maxLogFiles = 10
    
    # Ensure log directory exists
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory | Out-Null
    }
    
    # Log Rotation
    $logFiles = Get-ChildItem "$logDir\VMDeletion_*.log" -ErrorAction SilentlyContinue |
    Sort-Object CreationTime -Descending
    
    if ($logFiles.Count -ge $maxLogFiles) {
        $oldestLog = $logFiles[-1].FullName
        Remove-Item $oldestLog -ErrorAction SilentlyContinue
    }
    
    # Create Log Entry
    $logEntry = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Level = $Level.ToUpper()
        Message = $Message
        Component = $Component
        User = "$env:USERDOMAIN\$env:USERNAME"
        Machine = $env:COMPUTERNAME
    }
    
    # Write to log file
    $logFile = Join-Path $logDir "VMDeletion_$(Get-Date -Format 'yyyyMMdd').log"
    
    # Serialize to JSON
    $jsonLogEntry = ($logEntry | ConvertTo-Json -Compress) + "`r`n"
    $jsonLogEntry | Out-File -FilePath $logFile -Append -Encoding UTF8
    
    # Console Output
    $formattedMessage = "{0} [{1}] [{2}] {3} - {4}" -f 
        $logEntry.Timestamp, 
        $logEntry.Level, 
        $logEntry.Component, 
        $logEntry.User, 
        $logEntry.Message
    if ($NoGui -or $Help) {
        $consoleColor = switch ($logEntry.Level) {
            "ERROR"   { "Red" }
            "WARNING" { "Yellow" }
            "INFO"    { "Gray" }
            default   { "White" }
        }
        Write-Host $formattedMessage -ForegroundColor $consoleColor
    }
    
    if ($Level -eq "ERROR") {
        Add-GuiLogEntry -Message $Message -Level $Level -Component $Component
        Show-GuiErrorNotice
    }

    # Email Notification for Critical Errors
    if ($SendEmail -and $Level -eq "ERROR") {
        Send-ErrorNotificationEmail -LogEntry $logEntry
    }
}

function Send-ErrorNotificationEmail {
    param($LogEntry)
    
    try {
        $emailParams = @{
            From = $script:DefaultSettings.EmailNotification
            To = $script:DefaultSettings.EmailNotification
            Subject = "VM Deletion Tool - Critical Error"
            Body = @"
<html>
<body>
<h2>Critical Error Notification</h2>
<table border='1' cellpadding='5'>
<tr><th>Timestamp</th><td>$($LogEntry.Timestamp)</td></tr>
<tr><th>Level</th><td>$($LogEntry.Level)</td></tr>
<tr><th>Message</th><td>$($LogEntry.Message)</td></tr>
<tr><th>User</th><td>$($LogEntry.User)</td></tr>
<tr><th>Machine</th><td>$($LogEntry.Machine)</td></tr>
</table>
</body>
</html>
"@
            BodyAsHtml = $true
            SmtpServer = $script:DefaultSettings.SMTPServer
        }
        
        Send-MailMessage @emailParams
    }
    catch {
        # Fallback logging if email fails
        $fallbackLogPath = Join-Path $script:DefaultSettings.LogPath "email_errors.log"
        "Failed to send email: $($_.Exception.Message)" | 
            Out-File -FilePath $fallbackLogPath -Append
    }
}

function Save-EmailGroups {
    # שמירת קבוצות מה-ListBoxes להגדרות
    if ($script:lstGroupsCommon) {
        $groups = @()
        for ($i = 0; $i -lt $script:lstGroupsCommon.Items.Count; $i++) {
            $groups += $script:lstGroupsCommon.Items[$i].ToString()
        }
        $script:DefaultSettings.GroupsCommon = $groups -join ","
    }
    
    if ($script:lstGroupsWindows) {
        $groups = @()
        for ($i = 0; $i -lt $script:lstGroupsWindows.Items.Count; $i++) {
            $groups += $script:lstGroupsWindows.Items[$i].ToString()
        }
        $script:DefaultSettings.GroupsWindows = $groups -join ","
    }
    
    if ($script:lstGroupsLinux) {
        $groups = @()
        for ($i = 0; $i -lt $script:lstGroupsLinux.Items.Count; $i++) {
            $groups += $script:lstGroupsLinux.Items[$i].ToString()
        }
        $script:DefaultSettings.GroupsLinux = $groups -join ","
    }
}


#endregion

#region SQL Functions

function ConvertTo-WizardSqlIdentifier {
    param([Parameter(Mandatory=$true)][string]$TableName)
    $partPattern='(?:[A-Za-z_][A-Za-z0-9_@$#]*|\[[^\[\]\r\n]+\])'
    if($TableName.Trim() -notmatch ('^'+$partPattern+'(?:\.'+$partPattern+')?$')) {
        throw 'Use a table name or schema.table, for example dbo.TBL_SERVERS_ACTIVITIES.'
    }
    $builder=New-Object System.Data.SqlClient.SqlCommandBuilder
    try {
        $quoted=@(foreach($identifierMatch in [regex]::Matches($TableName.Trim(),$partPattern)) {
            $part=$identifierMatch.Value
            if($part.StartsWith('[')){$part=$part.Substring(1,$part.Length-2)}
            $builder.QuoteIdentifier($part)
        })
        return ($quoted -join '.')
    }
    finally {$builder.Dispose()}
}

function ConvertTo-DeletionSqlCredential {
    param([string]$UserName,[AllowEmptyString()][string]$Secret)
    $secure=New-Object Security.SecureString
    try {
        foreach($character in $Secret.ToCharArray()){$secure.AppendChar($character)}
        $secure.MakeReadOnly()
        return (New-Object Management.Automation.PSCredential($UserName,$secure))
    }
    catch {$secure.Dispose();throw}
}

function New-WizardSqlConnectionString {
    param([string]$Server,[string]$Database,[bool]$IntegratedSecurity,[Management.Automation.PSCredential]$Credential)
    if([string]::IsNullOrWhiteSpace($Server) -or [string]::IsNullOrWhiteSpace($Database)){throw 'Enter the SQL server and database.'}
    $builder=New-Object Data.SqlClient.SqlConnectionStringBuilder
    try {
        $builder['Data Source']=$Server.Trim();$builder['Initial Catalog']=$Database.Trim()
        $builder['Integrated Security']=$IntegratedSecurity;$builder['Connect Timeout']=8
        $builder['Application Name']='VM Deletion Tool'
        if(-not $IntegratedSecurity){
            if(-not $Credential){throw 'Enter the SQL authentication username and password.'}
            $builder['User ID']=$Credential.UserName
            $builder['Password']=$Credential.GetNetworkCredential().Password
        }
        return $builder.get_ConnectionString()
    }
    finally {$builder.Clear()}
}

function Save-VMDeletionToArchive {
    param([Parameter(Mandatory=$true)][hashtable]$VMDetails,[string]$Reason,[bool]$SCCMDeleted,[bool]$ADDeleted,[bool]$DNSDeleted,[bool]$VMDeleted,[string]$Status)
    try {
        $connectionString=Get-SQLConnectionString
        if([string]::IsNullOrWhiteSpace($connectionString)){throw 'SQL connection settings are incomplete.'}
        $serverName=$VMDetails.ServerName
        if([string]::IsNullOrWhiteSpace($serverName)){$serverName=$VMDetails.VMName}
        if([string]::IsNullOrWhiteSpace($serverName)){throw 'Cannot archive a deletion without a server name.'}
        $vc=$VMDetails.vCenter
        if([string]::IsNullOrWhiteSpace($vc)){$vc=$VMDetails.VCenterName}
        if([string]::IsNullOrWhiteSpace($vc)){$vc=(@($global:DefaultVIServer|ForEach-Object{$_.Name}) -join ', ')}
        $trackingData=@{
            Timestamp=(Get-Date);ServerName=$serverName;Action='Delete';Initiator="$env:USERDOMAIN\$env:USERNAME"
            Reason=$(if([string]::IsNullOrWhiteSpace($Reason)){'No reason provided'}else{$Reason})
            Platform=$(if($VMDetails.Platform){$VMDetails.Platform}else{'Virtual Server'})
            Network=$(if($VMDetails.Network){$VMDetails.Network}else{'Unknown'})
            vCenter=$(if($vc){$vc}else{'Unknown'});Owner=$VMDetails.Owner;Cluster=$VMDetails.Cluster
            OperatingSystem=$VMDetails.OperatingSystem;PowerState=$VMDetails.PowerState;PowerStateDate=$VMDetails.PowerStateDate
            CPUCores=$VMDetails.CPUCores;MemoryGB=$VMDetails.MemoryGB;TotalDisksGB=$VMDetails.TotalDisksGB;NBLastBackup=$VMDetails.NBLastBackup
        }
        $script:LastSqlArchiveError=''
        if(-not (Write-VMTrackingRecord -ConnectionString $connectionString -Data $trackingData)){
            throw $(if($script:LastSqlArchiveError){$script:LastSqlArchiveError}else{'SQL did not acknowledge the archive insert.'})
        }
        $script:sqlTableExists=$true
        return @{Success=$true;Message="Archived VM deletion info for $serverName"}
    }
    catch {
        Write-Log "Archive failed for '$($VMDetails.VMName)': $($_.Exception.Message)" -Level 'ERROR'
        return @{Success=$false;Message="Failed to archive: $($_.Exception.Message)"}
    }
    finally {$connectionString=$null}
}

function New-SQLArchiveTable {
    param([string]$ConnectionString,[string]$TableName)
    $connection=$null; $command=$null
    try {
        $quoted=ConvertTo-WizardSqlIdentifier -TableName $TableName
        $connection=New-DeletionSqlConnection -ConnectionString $ConnectionString
        $connection.Open()
        $command=$connection.CreateCommand()
        $command.CommandTimeout=60
        $command.CommandText=@"
IF OBJECT_ID(@TableName, N'U') IS NULL
BEGIN
    CREATE TABLE $quoted (
        [Timestamp] DATETIME2 NULL, [ServerName] NVARCHAR(128) NULL,
        [Action] NVARCHAR(50) NULL, [Owner] NVARCHAR(128) NULL,
        [Initiator] NVARCHAR(128) NULL, [Reason] NVARCHAR(512) NULL,
        [Platform] NVARCHAR(50) NULL, [Network] NVARCHAR(128) NULL,
        [vCenter] NVARCHAR(128) NULL, [Cluster] NVARCHAR(128) NULL,
        [OperatingSystem] NVARCHAR(128) NULL, [PowerState] NVARCHAR(50) NULL,
        [PowerStateDate] DATETIME2 NULL, [CPUCores] INT NULL,
        [MemoryGB] FLOAT NULL, [TotalDisksGB] FLOAT NULL, [NBLastBackup] NVARCHAR(128) NULL
    );
END;
"@
        [void]$command.Parameters.AddWithValue('@TableName',$TableName.Trim())
        [void]$command.ExecuteNonQuery()
        return @{Success=$true;Message="Table '$TableName' is ready."}
    }
    catch{return @{Success=$false;Message="Table creation failed: $($_.Exception.Message)"}}
    finally{if($command){$command.Dispose()};if($connection){$connection.Dispose()}}
}

function New-VMTrackingTable {
    param([string]$ConnectionString)
    $result=New-SQLArchiveTable -ConnectionString $ConnectionString -TableName $script:DefaultSettings.SQLTable
    Write-Log $result.Message -Level $(if($result.Success){'INFO'}else{'ERROR'})
}

function Write-VMTrackingRecord {
    param (
        [string]$ConnectionString,
        [hashtable]$Data
    )

    # ✅ Validate and ensure NOT NULL fields have values
    $notNullFields = @{
        'Timestamp'  = { Get-Date }
        'ServerName' = { "UNKNOWN" }
        'Action'     = { "Delete" }
        'Initiator'  = { "$env:USERDOMAIN\$env:USERNAME" }
        'Reason'     = { "No reason provided" }
        'Platform'   = { "Unknown" }
        'Network'    = { "Unknown" }
        'vCenter'    = { "Unknown" }
    }

    foreach ($field in $notNullFields.Keys) {
        if ([string]::IsNullOrWhiteSpace($Data[$field])) {
            $Data[$field] = & $notNullFields[$field]
            Write-Log "Field '$field' was empty, set to default: '$($Data[$field])'" -Level "WARNING"
        }
    }

    if ([string]::IsNullOrWhiteSpace($Data.ServerName)) {
        Write-Log "CRITICAL: Cannot insert record - ServerName is still empty after defaults" -Level "ERROR"
        return $false
    }

    Write-Log "=== Attempting SQL Insert ===" -Level "INFO"
    Write-Log "ServerName: $($Data.ServerName)" -Level "INFO"

    # Sanitize string fields
    $stringFields = @('ServerName', 'Action', 'Owner', 'Initiator', 'Reason', 
                      'Platform', 'Network', 'vCenter', 'Cluster', 'OperatingSystem', 
                      'PowerState', 'NBLastBackup')

    foreach ($field in $stringFields) {
        if ($Data.ContainsKey($field)) {
            if ([string]::IsNullOrWhiteSpace($Data[$field])) {
                if (-not $notNullFields.ContainsKey($field)) {
                    $Data[$field] = $null
                }
            } else {
                $Data[$field] = $Data[$field].ToString().Trim()
            }
        }
    }

    # Sanitize DateTime fields
    if ($Data.ContainsKey('Timestamp')) {
        if ($Data.Timestamp -isnot [datetime]) {
            try {
                $Data.Timestamp = [datetime]::Parse($Data.Timestamp)
            } catch {
                $Data.Timestamp = Get-Date
                Write-Log "Invalid Timestamp, using current time" -Level "WARNING"
            }
        }
    }

    if ($Data.ContainsKey('PowerStateDate') -and $Data.PowerStateDate) {
        if ($Data.PowerStateDate -isnot [datetime]) {
            try {
                $Data.PowerStateDate = [datetime]::Parse($Data.PowerStateDate)
            } catch {
                $Data.PowerStateDate = $null
                Write-Log "Invalid PowerStateDate, setting to NULL" -Level "WARNING"
            }
        }
    } else {
        $Data.PowerStateDate = $null
    }

    # Sanitize numeric fields
    $numericFields = @('CPUCores', 'MemoryGB', 'TotalDisksGB')

    foreach ($field in $numericFields) {
        if ($Data.ContainsKey($field)) {
            if ($null -eq $Data[$field] -or [string]::IsNullOrWhiteSpace($Data[$field].ToString())) {
                $Data[$field] = $null
            }
            else {
                try {
                    if ($field -eq 'CPUCores') {
                        $Data[$field] = [int]$Data[$field]
                    } else {
                        $Data[$field] = [decimal]$Data[$field]
                    }
                } 
                catch {
                    $Data[$field] = $null
                    Write-Log "Invalid $field format, setting to NULL" -Level "WARNING"
                }
            }
        }
    }

    # ✅ Transaction מפורש עם COMMIT
        $connection=$null;$command=$null;$transaction=$null;$committed=$false
    $script:LastSqlArchiveError=''
    try {
        $tableName=ConvertTo-WizardSqlIdentifier -TableName $script:DefaultSettings.SQLTable
        $connection=New-DeletionSqlConnection -ConnectionString $ConnectionString
        $connection.Open()
        $transaction=$connection.BeginTransaction()
        $command=$connection.CreateCommand()
        $command.Transaction=$transaction
        $command.CommandTimeout=60
        $command.CommandText=@"
SET XACT_ABORT ON;
SET NOCOUNT ON;
INSERT INTO $tableName (
    [Timestamp], [ServerName], [Action], [Owner], [Initiator], [Reason],
    [Platform], [Network], [vCenter], [Cluster], [OperatingSystem],
    [PowerState], [PowerStateDate], [CPUCores], [MemoryGB], [TotalDisksGB], [NBLastBackup]
) VALUES (
    @Timestamp, @ServerName, @Action, @Owner, @Initiator, @Reason,
    @Platform, @Network, @vCenter, @Cluster, @OperatingSystem,
    @PowerState, @PowerStateDate, @CPUCores, @MemoryGB, @TotalDisksGB, @NBLastBackup
);
SET @InsertedRows = @@ROWCOUNT;
"@
        foreach($key in $Data.Keys){
            $value=$(if($null -eq $Data[$key]){[DBNull]::Value}else{$Data[$key]})
            [void]$command.Parameters.AddWithValue("@$key",$value)
        }
        $acknowledged=$command.Parameters.Add('@InsertedRows',[System.Data.SqlDbType]::Int)
        $acknowledged.Direction=[System.Data.ParameterDirection]::Output
        [void]$command.ExecuteNonQuery()
        if($acknowledged.Value -is [DBNull] -or $null -eq $acknowledged.Value -or [int]$acknowledged.Value -ne 1){throw 'SQL did not confirm exactly one inserted archive record.'}
        $transaction.Commit();$committed=$true
        Write-Log "SQL archive insert committed for '$($Data.ServerName)'."
        return $true
    }
    catch {
        $script:LastSqlArchiveError=$_.Exception.Message
        if($transaction -and -not $committed){try{$transaction.Rollback()}catch{}}
        Write-Log "SQL archive insert failed for '$($Data.ServerName)': $script:LastSqlArchiveError" -Level 'ERROR'
        return $false
    }
    finally {if($command){$command.Dispose()};if($transaction){$transaction.Dispose()};if($connection){$connection.Dispose()}}
}


function Get-SQLConnectionString {
    $credential=$null
    try {
        if(-not $script:DefaultSettings.SQLIntegratedAuth){$credential=ConvertTo-DeletionSqlCredential -UserName $script:DefaultSettings.SQLUsername -Secret $script:DefaultSettings.SQLPassword}
        return (New-WizardSqlConnectionString -Server $script:DefaultSettings.SQLServer -Database $script:DefaultSettings.SQLDatabase -IntegratedSecurity $script:DefaultSettings.SQLIntegratedAuth -Credential $credential)
    }
    catch {Write-Log "Invalid SQL configuration: $($_.Exception.Message)" -Level 'ERROR';return $null}
    finally {if($credential){$credential.Password.Dispose()}}
}

function Test-SQLConnection {
    param([string]$ConnectionString)
    
    $result = @{
        Success = $false
        Message = ""
    }
    
    if ([string]::IsNullOrWhiteSpace($ConnectionString)) {
        $result.Message = "Connection string is empty"
        return $result
    }
    
    $connection=$null; $command=$null; $adapter=$null
    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $ConnectionString
        $connection.Open()
        
        # בדיקה שאפשר לקרוא מה-DB
        $command = New-Object System.Data.SqlClient.SqlCommand("SELECT 1", $connection)
        $command.ExecuteScalar() | Out-Null
        
        $connection.Close()
        $result.Success = $true
        $result.Message = "Connected successfully to $($script:DefaultSettings.SQLServer)"
    }
    catch [System.Data.SqlClient.SqlException] {
        $sqlError = $_.Exception
        switch ($sqlError.Number) {
            18456 { $result.Message = "Login failed - check username/password" }
            4060  { $result.Message = "Database '$($script:DefaultSettings.SQLDatabase)' not found" }
            53    { $result.Message = "Server not found - check server name" }
            -1    { $result.Message = "Connection timeout - server unreachable" }
            default { $result.Message = "SQL Error $($sqlError.Number): $($sqlError.Message)" }
        }
    }
    catch {
        $result.Message = "Connection failed: $($_.Exception.Message)"
    }
    
    finally {if($adapter){$adapter.Dispose()};if($command){$command.Dispose()};if($connection){$connection.Dispose()}}
    return $result
}

function Test-SQLTableExists {
    param([string]$ConnectionString,[string]$TableName)
    $result=@{Success=$false;Exists=$false;Message='';Columns=@()}
    $connection=$null;$command=$null;$adapter=$null
    try {
        [void](ConvertTo-WizardSqlIdentifier -TableName $TableName)
        $connection=New-DeletionSqlConnection -ConnectionString $ConnectionString
        $connection.Open()
        $command=$connection.CreateCommand()
        $command.CommandText='SELECT name AS COLUMN_NAME FROM sys.columns WHERE object_id=OBJECT_ID(@TableName) ORDER BY column_id'
        [void]$command.Parameters.AddWithValue('@TableName',$TableName.Trim())
        $adapter=New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $table=New-Object System.Data.DataTable
        [void]$adapter.Fill($table)
        $result.Success=$true;$result.Exists=($table.Rows.Count -gt 0)
        $result.Columns=@($table.Rows|ForEach-Object{$_.COLUMN_NAME})
        $result.Message=$(if($result.Exists){"Table exists with $($table.Rows.Count) columns"}else{'Table does not exist or its metadata is unavailable.'})
    }
    catch {$result.Message="Error checking table: $($_.Exception.Message)"}
    finally {if($adapter){$adapter.Dispose()};if($command){$command.Dispose()};if($connection){$connection.Dispose()}}
    return $result
}

function Get-VMDetailsForArchive {
    param([string]$VMName)
    
    $details = @{
        # === שדות למייל (Send-VMDeletionNotification) ===
        VMName          = $VMName
        VCenterName     = $null
        ClusterName     = $null
        FolderName      = $null
        ServerOwner     = $null
        FolderOwner     = $null
        IsTemplate      = $false
        TotalCPUs       = $null
        TotalMemoryGB   = $null
        TotalDisksGB    = $null
        OperatingSystem = $null
        PowerState      = $null
        
        # === שדות ל-Archive (Save-VMDeletionToArchive) ===
        Timestamp       = Get-Date
        ServerName      = $VMName
        Action          = "Delete"
        Owner           = $null
        Initiator       = "$env:USERDOMAIN\$env:USERNAME"
        Reason          = $null
        Platform        = "Virtual Server"
        Network         = "External"
        vCenter         = $null
        Cluster         = $null
        PowerStateDate  = $null
        CPUCores        = $null
        MemoryGB        = $null
        NBLastBackup    = $null
    }
    
    try {
        try{$vm=Get-VM -Name $VMName -ErrorAction Stop}
        catch{$vm=Get-Template -Name $VMName -ErrorAction Stop;$details.IsTemplate=$true}
       
        # vCenter Name
        if ($global:DefaultVIServer -and $global:DefaultVIServer.Name) {
            $details.VCenterName = $global:DefaultVIServer.Name
            $details.vCenter = $global:DefaultVIServer.Name  # לArchive
        }
        
        # Cluster Name
        $cluster = Get-Cluster -VM $vm -ErrorAction SilentlyContinue
        if ($cluster) { 
            $details.ClusterName = $cluster.Name
            $details.Cluster = $cluster.Name  # לArchive
        }

        # Folder Name
        $folder = $vm.Folder
        if ($folder) {
            $details.FolderName = $folder.Name
        }

        $owners=Get-DeletionOwnerMetadata -VM $vm
        $details.ServerOwner=$owners.ServerOwner
        $details.FolderOwner=$owners.FolderOwner
        $details.Owner=$(if($owners.ServerOwner){$owners.ServerOwner}else{$owners.FolderOwner})
        # NB_LAST_BACKUP
        try {
            $nbAttr = Get-Annotation -Entity $vm -CustomAttribute "NB_LAST_BACKUP" -ErrorAction SilentlyContinue
            if ($nbAttr -and $nbAttr.Value) {
                $details.NBLastBackup = $nbAttr.Value
            }
        }
        catch { 
            $details.NBLastBackup = $null 
        }

        # srvPSDate
        try {
            $psDate = Get-Annotation -Entity $vm -CustomAttribute "srvPSDate" -ErrorAction SilentlyContinue
            if ($psDate -and $psDate.Value) {
                try {
                    $details.PowerStateDate = [datetime]::Parse($psDate.Value)
                }
                catch {
                    $details.PowerStateDate = $null
                }
            }
        } 
        catch { 
            $details.PowerStateDate = $null 
        }

        # Power State
        if ($vm.PowerState) {
            $details.PowerState = $vm.PowerState.ToString()
        }
        
        # CPUs
        if ($vm.NumCpu) {
            $details.TotalCPUs = [int]$vm.NumCpu
            $details.CPUCores = [int]$vm.NumCpu  # לArchive
        }
        
        # Memory (GB)
        if ($vm.MemoryGB) {
            $details.TotalMemoryGB = [math]::Round($vm.MemoryGB, 2)
            $details.MemoryGB = [math]::Round($vm.MemoryGB, 2)  # לArchive
        }
        
        # Total Disk Size (GB)
        $disks = Get-HardDisk -VM $vm -ErrorAction SilentlyContinue
        if ($disks) {
            $totalDiskGB = ($disks | Measure-Object -Property CapacityGB -Sum).Sum
            if ($totalDiskGB) {
                $details.TotalDisksGB = [math]::Round($totalDiskGB, 2)
            }
        }
        
        # Operating System
        $osFullName = $vm.Guest.OSFullName
        $guestId = $vm.GuestId
        
        if ($osFullName) { 
            $details.OperatingSystem = $osFullName 
        } 
        elseif ($guestId) { 
            $details.OperatingSystem = $guestId 
        }
    }
    catch {
        Write-Log "Error getting VM details for $VMName : $($_.Exception.Message)" -Level "WARNING"
    }
    
    return $details
}

function Get-ArchiveData {

    param([string]$ConnectionString, [string]$TableName)
    
    $result = @{
        Success = $false
        Message = ""
        Data    = $null
    }
    
    $connection=$null; $command=$null; $adapter=$null
    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $ConnectionString
        $connection.Open()
        
        $query = "SELECT * FROM " + (ConvertTo-WizardSqlIdentifier -TableName $TableName) + " ORDER BY [TIMESTAMP] DESC"
        $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $dataTable = New-Object System.Data.DataTable
        $adapter.Fill($dataTable) | Out-Null
        
        $connection.Close()
        
        $result.Success = $true
        $result.Data = $dataTable
        $result.Message = "Retrieved $($dataTable.Rows.Count) records"
    }
    catch {
        $result.Message = "Error retrieving data: $($_.Exception.Message)"
    }
    
    finally {if($adapter){$adapter.Dispose()};if($command){$command.Dispose()};if($connection){$connection.Dispose()}}
    return $result
}

function Export-ArchiveToFile {
    param(
        [System.Data.DataTable]$Data,
        [string]$FilePath,
        [string]$Format  # "CSV" or "XLSX"
    )
    
    $result = @{
        Success = $false
        Message = ""
    }
    
    try {
        if ($Format -eq "CSV") {
            $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
            $result.Success = $true
            $result.Message = "Exported to CSV: $FilePath"
        }
        elseif ($Format -eq "XLSX") {
            # Check if ImportExcel module is available
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel
                $Data | Export-Excel -Path $FilePath -AutoSize -TableName "VMDeletionArchive" -FreezeTopRow
                $result.Success = $true
                $result.Message = "Exported to Excel: $FilePath"
            }
            else {
                # Fallback to CSV if ImportExcel not available
                $csvPath = $FilePath -replace '\.xlsx$', '.csv'
                $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                $result.Success = $true
                $result.Message = "ImportExcel module not found. Exported to CSV: $csvPath"
            }
        }
    }
    catch {
        $result.Message = "Export error: $($_.Exception.Message)"
    }
    
    return $result
}

function Export-ArchiveToExcel {
    param (
        [Parameter(Mandatory = $true)]
        $Data,
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        # יצירת אובייקט Excel
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # יצירת Workbook חדש
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Archive Data"
        
        # קבלת העמודות
        $columns = $Data | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
        
        # כתיבת כותרות
        $col = 1
        foreach ($column in $columns) {
            $worksheet.Cells.Item(1, $col) = $column
            $worksheet.Cells.Item(1, $col).Font.Bold = $true
            $worksheet.Cells.Item(1, $col).Interior.ColorIndex = 15  # אפור
            $col++
        }
        
        # כתיבת הנתונים
        $row = 2
        foreach ($item in $Data) {
            $col = 1
            foreach ($column in $columns) {
                $worksheet.Cells.Item($row, $col) = $item.$column
                $col++
            }
            $row++
        }
        
        # התאמת רוחב עמודות
        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        
        # שמירה
        $workbook.SaveAs($FilePath)
        $workbook.Close($true)
        $excel.Quit()
        
        # שחרור זיכרון
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        return @{ Success = $true; Message = "Exported successfully to: $FilePath" }
    }
    catch {
        return @{ Success = $false; Message = "Export failed: $($_.Exception.Message)" }
    }
}

#endregion SQL Functions

#region Deletion Functions

function Remove-SCCMDevice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$SCCMServer,
        
        [Parameter(Mandatory = $true)]
        [string]$SiteCode
    )
    
    $result = @{
        Success = $false
        Message = ""
        Found   = $false
    }
    
    try {
        $namespace = "root\SMS\site_$SiteCode"
        
        $query = "SELECT * FROM SMS_R_System WHERE Name = '$ComputerName'"
        $device = Get-WmiObject -ComputerName $SCCMServer -Namespace $namespace -Query $query -ErrorAction Stop
        
        if ($null -eq $device) {
            $result.Message = "Device not found in SCCM"
            $result.Found = $false
            $result.Success = $true
            return $result
        }
        
        $result.Found = $true
        $deviceList = @($device)
        $deletedCount = 0
        
        foreach ($dev in $deviceList) {
            $resourceId = $dev.ResourceId
            $deviceToDelete = Get-WmiObject -ComputerName $SCCMServer -Namespace $namespace `
                -Query "SELECT * FROM SMS_R_System WHERE ResourceId = $resourceId" -ErrorAction Stop
            
            if ($null -ne $deviceToDelete) {
                $deviceToDelete.Delete()
                $deletedCount++
            }
        }
        
        $result.Success = $true
        $result.Message = "Deleted $deletedCount record(s) from SCCM"
    } 
    catch [System.UnauthorizedAccessException] {
        $result.Message = "Access denied to SCCM server"
    }
    catch [System.Runtime.InteropServices.COMException] {
        $result.Message = "Cannot connect to SCCM server"
    }
    catch {
        $result.Message = "SCCM Error: $($_.Exception.Message)"
    }
    
    return $result
}

function Remove-ADDevice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [string]$DomainController,
        
        [Parameter(Mandatory = $false)]
        [string]$SearchBase 
    )

    $result = @{
        Success  = $false
        Message  = ""
        Found    = $false
        Location = ""
    }

    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            $result.Message = "ActiveDirectory module not installed"
            return $result
        }

        Import-Module ActiveDirectory -ErrorAction Stop

        # קבע את ה-SearchBase - פרמטר > הגדרות > ברירת מחדל
        $effectiveSearchBase = if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
            $SearchBase
        } elseif (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.ADSearchBase)) {
            $script:DefaultSettings.ADSearchBase
        } else {
            $null  # יחפש בכל ה-AD
        }

        # קבע את ה-DomainController
        $effectiveDC = if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
            $DomainController
        } elseif (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.DomainController)) {
            $script:DefaultSettings.DomainController
        } else {
            $null
        }

        # בנה את הפרמטרים לחיפוש
        $searchParams = @{
            Filter      = "Name -eq '$ComputerName'"
            ErrorAction = 'Stop'
        }

        # הוסף SearchBase אם הוגדר
        if (-not [string]::IsNullOrWhiteSpace($effectiveSearchBase)) {
            $searchParams.SearchBase = $effectiveSearchBase
            $searchParams.SearchScope = 'Subtree'  # חפש גם בתתי-OU
            Write-Log "Searching for '$ComputerName' in OU: $effectiveSearchBase" -Level "INFO"
        } else {
            Write-Log "Searching for '$ComputerName' in entire AD" -Level "WARNING"
        }

        # הוסף DC אם הוגדר
        if (-not [string]::IsNullOrWhiteSpace($effectiveDC)) {
            $searchParams.Server = $effectiveDC
        }

        # חפש את המחשב
        $computer = $null
        try {
            $computer = Get-ADComputer @searchParams
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $result.Message = "Server not found in AD"
            $result.Found = $false
            $result.Success = $true
            return $result
        }

        if ($null -eq $computer) {
            $result.Message = "Server not found "
            $result.Found = $false
            $result.Success = $true
            return $result
        }

        # נמצא!
        $result.Found = $true
        $result.Location = $computer.DistinguishedName
        Write-Log "Found computer at: $($computer.DistinguishedName)" -Level "INFO"

        # מחק את האובייקט
        $removeParams = @{
            Identity    = $computer.DistinguishedName
            Confirm     = $false
            ErrorAction = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($effectiveDC)) {
            $removeParams.Server = $effectiveDC
        }

        Remove-ADComputer @removeParams

        $result.Success = $true
        $result.Message = "Server object deleted successfully"
        Write-Log "✅ Deleted AD object: $ComputerName" -Level "INFO"

    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        $result.Message = "Server not found in specified OU"
        $result.Found = $false
        $result.Success = $true
    }
    catch [System.UnauthorizedAccessException] {
        $result.Message = "Access denied - insufficient AD permissions"
        Write-Log "❌ AD Access denied for: $ComputerName" -Level "ERROR"
    }
    catch {
        $result.Message = "AD Error: $($_.Exception.Message)"
        Write-Log "❌ AD Error: $($_.Exception.Message)" -Level "ERROR"
    }

    return $result
}

function Remove-DNSRecords {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$DNSServer,
        
        [Parameter(Mandatory = $true)]
        [string]$DNSZone
    )
    
    $result = @{
        Success      = $false
        Message      = ""
        Found        = $false
        DeletedCount = 0
    }
    
    try {
        if (-not (Get-Module -ListAvailable -Name DnsServer)) {
            $result.Message = "DnsServer module not installed"
            return $result
        }
        
        Import-Module DnsServer -ErrorAction Stop
        
        $deletedRecords = 0
        $foundRecords = 0
        
        # A Records
        try {
            $aRecords = Get-DnsServerResourceRecord -ComputerName $DNSServer `
                -ZoneName $DNSZone `
                -Name $ComputerName `
                -RRType A `
                -ErrorAction SilentlyContinue
            
            if ($null -ne $aRecords) {
                $foundRecords += @($aRecords).Count
                
                foreach ($record in $aRecords) {
                    Remove-DnsServerResourceRecord -ComputerName $DNSServer `
                        -ZoneName $DNSZone `
                        -Name $ComputerName `
                        -RRType A `
                        -RecordData $record.RecordData.IPv4Address.IPAddressToString `
                        -Force `
                        -ErrorAction Stop
                    $deletedRecords++
                }
            }
        }
        catch [Microsoft.Management.Infrastructure.CimException] { }
        
        # PTR Records
        try {
            $aRecordsForPtr = Get-DnsServerResourceRecord -ComputerName $DNSServer `
                -ZoneName $DNSZone `
                -Name $ComputerName `
                -RRType A `
                -ErrorAction SilentlyContinue
            
            if ($null -ne $aRecordsForPtr) {
                foreach ($aRecord in $aRecordsForPtr) {
                    $ipAddress = $aRecord.RecordData.IPv4Address.IPAddressToString
                    $ipParts = $ipAddress.Split('.')
                    $reverseZone = "$($ipParts[2]).$($ipParts[1]).$($ipParts[0]).in-addr.arpa"
                    $ptrName = $ipParts[3]
                    
                    try {
                        $ptrRecord = Get-DnsServerResourceRecord -ComputerName $DNSServer `
                            -ZoneName $reverseZone `
                            -Name $ptrName `
                            -RRType PTR `
                            -ErrorAction SilentlyContinue
                        
                        if ($null -ne $ptrRecord) {
                            $foundRecords++
                            Remove-DnsServerResourceRecord -ComputerName $DNSServer `
                                -ZoneName $reverseZone `
                                -Name $ptrName `
                                -RRType PTR `
                                -Force `
                                -ErrorAction SilentlyContinue
                            $deletedRecords++
                        }
                    }
                    catch { }
                }
            }
        }
        catch { }
        
        # AAAA Records
        try {
            $aaaaRecords = Get-DnsServerResourceRecord -ComputerName $DNSServer `
                -ZoneName $DNSZone `
                -Name $ComputerName `
                -RRType AAAA `
                -ErrorAction SilentlyContinue
            
            if ($null -ne $aaaaRecords) {
                $foundRecords += @($aaaaRecords).Count
                
                Remove-DnsServerResourceRecord -ComputerName $DNSServer `
                    -ZoneName $DNSZone `
                    -Name $ComputerName `
                    -RRType AAAA `
                    -Force `
                    -ErrorAction SilentlyContinue
                $deletedRecords += @($aaaaRecords).Count
            }
        }
        catch { }
        
        $result.Found = ($foundRecords -gt 0)
        $result.DeletedCount = $deletedRecords
        
        if ($foundRecords -eq 0) {
            $result.Success = $true
            $result.Message = "No DNS records found"
        }
        elseif ($deletedRecords -gt 0) {
            $result.Success = $true
            $result.Message = "Deleted $deletedRecords DNS record(s)"
        }
        else {
            $result.Message = "Found records but failed to delete"
        }
    }
    catch [System.UnauthorizedAccessException] {
        $result.Message = "Access denied to DNS server"
    }
    catch {
        $result.Message = "DNS Error: $($_.Exception.Message)"
    }
    
    return $result
}

#endregion Deletion Functions

function Get-VCenterServerForEnvironment {
    param([string]$EnvironmentName)

    $envName = Resolve-VCenterEnvironment -Value $EnvironmentName
    foreach ($vc in $script:vCenterServers) {
        if ($vc.Name -eq $envName -or $vc.Server -eq $envName) {
            return @{ Name = $vc.Name; Server = $vc.Server }
        }
    }

    return @{ Name = $envName; Server = $envName }
}

function Import-VMDeletionCsv {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "Missing required -CsvPath.`nNext step: run with -NoGui -CsvPath 'C:\Input\vms.csv', or use -Help for examples."
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "CSV file was not found: $Path`nNext step: verify the path exists and that the Task Scheduler account can read it."
    }

    try {
        $csvData = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    }
    catch {
        throw "CSV could not be read: $Path`nReason: $($_.Exception.Message)`nNext step: check file encoding, delimiter, and read permissions."
    }

    if ($csvData.Count -eq 0) {
        throw "CSV file is empty: $Path`nNext step: add at least one row with a VMName value."
    }

    $hasVMName = $csvData | Get-Member -Name "VMName" -MemberType NoteProperty
    if (-not $hasVMName) {
        throw "CSV is missing required column: VMName.`nExpected minimum format:`nVMName`nSERVER01`nSERVER02`nNext step: rename the server-name column to VMName."
    }

    $rows = @()
    $skippedBlankRows = 0
    foreach ($row in $csvData) {
        $vmName = [string]$row.VMName
        if ([string]::IsNullOrWhiteSpace($vmName)) {
            $skippedBlankRows++
            continue
        }

        $rowEnvironment = $script:RunEnvironment
        if ($row.PSObject.Properties.Name -contains "Environment" -and -not [string]::IsNullOrWhiteSpace($row.Environment)) {
            $rowEnvironment = $row.Environment
        }
        elseif ($row.PSObject.Properties.Name -contains "vCenter" -and -not [string]::IsNullOrWhiteSpace($row.vCenter)) {
            $rowEnvironment = Resolve-VCenterEnvironment -Value $row.vCenter
        }

        $rows += [pscustomobject]@{
            VMName      = $vmName.Trim()
            Environment = $rowEnvironment
            OSType = [string]$row.OSType
        }
    }

    if ($rows.Count -eq 0) {
        throw "CSV does not contain valid VMName values.`nNext step: fill the VMName column with at least one VM or Template name."
    }

    if ($skippedBlankRows -gt 0) {
        Write-Log "CSV validation skipped $skippedBlankRows blank VMName row(s)." -Level "WARNING"
    }

    Write-Log "CSV validation completed. Rows=$($rows.Count), Source='$Path'"
    return $rows
}

function Show-NoGuiUsage {
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    $usage = @"
VM Deletion Tool - NoGui usage

Configuration:
  -ConfigPath <path-to-config.json>       Use and save this explicit JSON file.
  -SharedConfigPath <path-to-shared.json> Override the optional shared source.
  Without -ConfigPath: shared JSON, then %TEMP%\VMDeletionTool\config.json, then defaults.
  Persistent workflow flags also apply here. CLI -Skip switches only disable more steps.
  Linux skips SCCM only. AD follows the saved preference and -SkipAD.

Required:
  -NoGui -CsvPath <path-to-csv>

Common example:
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$scriptPath" -NoGui -CsvPath "C:\Input\vms.csv" -Environment PROD

Safe validation example, no cleanup actions enabled:
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$scriptPath" -NoGui -CsvPath "C:\Input\vms.csv" -Environment PROD -SkipSCCM -SkipAD -SkipDNS -SkipVM -SkipArchive -SkipEmail

CSV minimum format:
  VMName
  SERVER01
  SERVER02

CSV recommended format:
  VMName,Environment
  SERVER01,PROD
  SERVER02,MGMT
  SERVER03,VDI

Exit codes:
  0  Completed without VM-level failures
  1  Startup validation failed or one/more VM failures occurred
"@
    Write-Host $usage -ForegroundColor White
}

function Get-NoGuiActionSummary {
    $enabled = @()
    $skipped = @()
    foreach ($item in @(
        @{ Name = "SCCM";    Skipped = ([bool]$SkipSCCM -or -not (Test-DeletionStepEnabled -Step SCCM)) },
        @{ Name = "AD";      Skipped = ([bool]$SkipAD -or -not (Test-DeletionStepEnabled -Step AD)) },
        @{ Name = "DNS";     Skipped = ([bool]$SkipDNS -or -not (Test-DeletionStepEnabled -Step DNS)) },
        @{ Name = "VM";      Skipped = ([bool]$SkipVM -or -not (Test-DeletionStepEnabled -Step VM)) },
        @{ Name = "Archive"; Skipped = ([bool]$SkipArchive -or -not (Test-DeletionStepEnabled -Step Archive)) },
        @{ Name = "Email";   Skipped = ([bool]$SkipEmail -or -not (Test-DeletionStepEnabled -Step Email)) }
    )) {
        if ($item.Skipped) { $skipped += $item.Name } else { $enabled += $item.Name }
    }

    return [pscustomobject]@{
        Enabled = if ($enabled.Count -gt 0) { $enabled -join ", " } else { "None" }
        Skipped = if ($skipped.Count -gt 0) { $skipped -join ", " } else { "None" }
    }
}
function Invoke-HeadlessCsvDeletion {
    param([string]$Path)

    $actions = Get-NoGuiActionSummary
    Write-Log "NoGui run started. CSV='$Path' Environment='$script:RunEnvironment' Enabled=[$($actions.Enabled)] Skipped=[$($actions.Skipped)]"
    $rows = Import-VMDeletionCsv -Path $Path
    $rowCount = @($rows).Count
    $environmentGroupCount = @($rows | Group-Object Environment).Count
    Write-Log "NoGui will process $rowCount VM(s) across $environmentGroupCount environment group(s)."
    $successCount = 0
    $failedCount = 0

    if ((-not $SkipVM -and (Test-DeletionStepEnabled -Step VM))) {
        if (-not (Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) {
            Import-Module VMware.VimAutomation.Core -ErrorAction Stop
        }
        Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }

    foreach ($group in ($rows | Group-Object Environment)) {
        $vcInfo = Get-VCenterServerForEnvironment -EnvironmentName $group.Name
        $connectedForGroup = $false

        if ((-not $SkipVM -and (Test-DeletionStepEnabled -Step VM))) {
            try {
                Write-Log "Connecting to vCenter '$($vcInfo.Name)' ($($vcInfo.Server))"
                Connect-VIServer -Server $vcInfo.Server -ErrorAction Stop | Out-Null
                $connectedForGroup = $true
            }
            catch {
                Write-Log "Failed to connect to vCenter '$($vcInfo.Server)': $($_.Exception.Message)" -Level "ERROR"
                $failedCount += $group.Group.Count
                continue
            }
        }

        foreach ($row in $group.Group) {
            $vmName = $row.VMName
            $vmEnvironment = $row.Environment
            $vmErrors = @()
            Write-Log "NoGui deletion started for '$vmName' in environment '$vmEnvironment'"

            if ($SkipVM -or -not (Test-DeletionStepEnabled -Step VM)) {
                $vmDetails = @{ ServerName = $vmName; VMName = $vmName }
            }
            else {
                try { $vmDetails = Get-VMDetailsForArchive -VMName $vmName }
                catch { $vmDetails = @{ ServerName = $vmName; VMName = $vmName } }
            }

            $osType=Get-DeletionOSType -Preferred $row.OSType -OperatingSystem $vmDetails.OperatingSystem -GuestId $vmDetails.GuestId
            $sccmSkip=Get-DeletionStepSkipReason -Step SCCM -OSType $osType -Selected (-not $SkipSCCM)
            $sccmDeleted = $false
            $adDeleted = $false
            $dnsDeleted = $false
            $vmDeleted = $false

            if (-not $sccmSkip) {
                $sccmResult = Remove-SCCMDevice -ComputerName $vmName -SCCMServer $script:DefaultSettings.SCCMServer -SiteCode $script:DefaultSettings.SCCMSiteCode
                $sccmDeleted = $sccmResult.Found
                if (-not $sccmResult.Success) { $vmErrors += "SCCM: $($sccmResult.Message)" }
            }

            if($sccmSkip){Write-Log "$vmName - SCCM: $sccmSkip"}
            if ((-not $SkipAD -and (Test-DeletionStepEnabled -Step AD))) {
                $adSearchBase = Get-ADSearchBaseForEnvironment -EnvironmentName $vmEnvironment
                $adResult = Remove-ADDevice -ComputerName $vmName -DomainController $script:DefaultSettings.DomainController -SearchBase $adSearchBase
                $adDeleted = $adResult.Found
                if (-not $adResult.Success) { $vmErrors += "AD: $($adResult.Message)" }
            }

            if ((-not $SkipDNS -and (Test-DeletionStepEnabled -Step DNS))) {
                $dnsResult = Remove-DNSRecords -ComputerName $vmName -DNSServer $script:DefaultSettings.DNSServer -DNSZone $script:DefaultSettings.DNSZone
                $dnsDeleted = $dnsResult.Found
                if (-not $dnsResult.Success) { $vmErrors += "DNS: $($dnsResult.Message)" }
            }

            if ((-not $SkipVM -and (Test-DeletionStepEnabled -Step VM))) {
                try {
                    $template = Get-Template -Name $vmName -ErrorAction SilentlyContinue
                    if ($template) {
                        if ($script:DefaultSettings.DeleteFromDisk) {
                            Remove-Template -Template $template -DeletePermanently -Confirm:$false -ErrorAction Stop
                        }
                        else {
                            Remove-Template -Template $template -Confirm:$false -ErrorAction Stop
                        }
                    }
                    else {
                        $vm = Get-VM -Name $vmName -ErrorAction Stop
                        if ($vm.PowerState -eq "PoweredOn") {
                            Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null
                            $timeout = $script:DefaultSettings.PowerOffTimeout
                            $elapsed = 0
                            while ((Get-VM -Name $vmName -ErrorAction SilentlyContinue).PowerState -eq "PoweredOn" -and $elapsed -lt $timeout) {
                                Start-Sleep -Seconds 2
                                $elapsed += 2
                            }
                        }

                        if ($script:DefaultSettings.DeleteFromDisk) {
                            Remove-VM -VM $vm -DeletePermanently -Confirm:$false -ErrorAction Stop
                        }
                        else {
                            Remove-VM -VM $vm -Confirm:$false -ErrorAction Stop
                        }
                    }
                    $vmDeleted = $true
                }
                catch {
                    $vmErrors += "VM: $($_.Exception.Message)"
                }
            }

            $status = if ($vmErrors.Count -eq 0) { "Success" } else { "Failed" }

            if ((-not $SkipArchive -and (Test-DeletionStepEnabled -Step Archive))) {
                $archiveResult = Save-VMDeletionToArchive -VMDetails $vmDetails -Reason $Reason -SCCMDeleted $sccmDeleted -ADDeleted $adDeleted -DNSDeleted $dnsDeleted -VMDeleted $vmDeleted -Status $status
                if (-not $archiveResult.Success) {
                    $vmErrors += "Archive: $($archiveResult.Message)"
                }
            }

            if ((-not $SkipEmail -and (Test-DeletionStepEnabled -Step Email)) -and $script:DefaultSettings.EnableEmailNotification) {
                Send-VMDeletionNotification -VMDetails $vmDetails -Reason $Reason -OSType $osType -DeletionSuccess $vmDeleted | Out-Null
            }

            if ($vmErrors.Count -eq 0) {
                $successCount++
                Write-Log "NoGui deletion completed successfully for '$vmName'"
            }
            else {
                $failedCount++
                Write-Log "NoGui deletion failed for '$vmName': $($vmErrors -join '; ')" -Level "ERROR"
            }
        }

        if ($connectedForGroup) {
            Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue
        }
    }

    Write-Log "NoGui CSV deletion completed. Success=$successCount Failed=$failedCount"
    if ($failedCount -gt 0) { return 1 }
    return 0
}

# Settings are loaded by Initialize-Configuration above. Do not reload from TEMP here.
if ($Help) {
    Show-NoGuiUsage
    exit 0
}

if ($NoGui) {
    try {
        if ([string]::IsNullOrWhiteSpace($CsvPath)) {
            Show-NoGuiUsage
            throw "Missing required -CsvPath."
        }

        $exitCode = Invoke-HeadlessCsvDeletion -Path $CsvPath
        if ($exitCode -eq 0) {
            Write-Host "Completed successfully. ExitCode=0" -ForegroundColor Green
        }
        else {
            Write-Host "Completed with failures. ExitCode=$exitCode. Next step: review the VMDeletion log and failed VM entries." -ForegroundColor Yellow
        }
        exit $exitCode
    }
    catch {
        Write-Log "NoGui run failed: $($_.Exception.Message)" -Level "ERROR"
        Write-Host "Startup failed. ExitCode=1. Next step: fix the error above and rerun with -Help if needed." -ForegroundColor Red
        exit 1
    }
}
#endregion

#region Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "VM Deletion Tool"
$form.Size = New-Object System.Drawing.Size(820, 950)
$form.Height = [Math]::Min($form.Height, [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height - 24)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(22, 34, 53)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.TopMost = $false

# Modern frontend palette: calm charcoal surfaces, clear accent, no pure white glare.
$uiAccent = [System.Drawing.Color]::FromArgb(120, 190, 211)
$uiAccent2 = [System.Drawing.Color]::FromArgb(45, 105, 137)
$uiDanger = [System.Drawing.Color]::FromArgb(164, 66, 78)
$uiWarning = [System.Drawing.Color]::FromArgb(139, 111, 62)
$uiSuccess = [System.Drawing.Color]::FromArgb(51, 116, 101)
$uiText = [System.Drawing.Color]::FromArgb(229, 236, 245)
$uiMuted = [System.Drawing.Color]::FromArgb(177, 192, 211)
$uiButtonText = [System.Drawing.Color]::FromArgb(248, 249, 252)
$uiPanel = [System.Drawing.Color]::FromArgb(29, 46, 69)
$uiPanel2 = [System.Drawing.Color]::FromArgb(38, 58, 83)
$uiInput = [System.Drawing.Color]::FromArgb(20, 34, 53)
$uiInput2 = [System.Drawing.Color]::FromArgb(25, 41, 62)
$uiBorder = [System.Drawing.Color]::FromArgb(83, 109, 139)
$uiBorderSoft = [System.Drawing.Color]::FromArgb(55, 77, 104)
$uiBlack = [System.Drawing.Color]::FromArgb(22, 34, 53)

function Set-ModernButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor,
        [int]$BorderSize = 1,
        [int]$Radius = 6
    )
    $hoverColor = [System.Windows.Forms.ControlPaint]::Light($BackColor, 0.10)
    $borderColor = [System.Windows.Forms.ControlPaint]::Light($BackColor, 0.18)
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = $BorderSize
    $Button.FlatAppearance.BorderColor = $borderColor
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.UseVisualStyleBackColor = $false
    $Button.Padding = New-Object System.Windows.Forms.Padding(6, 2, 6, 2)
    $Button.FlatAppearance.MouseDownBackColor = [System.Windows.Forms.ControlPaint]::Dark($BackColor, 0.12)
    $Button.Region = New-RoundedRegion -Width $Button.Width -Height $Button.Height -Radius $Radius
    $Button.Tag = @{ Radius = $Radius; BackColor = $BackColor; HoverColor = $hoverColor; BorderColor = $borderColor }
    $Button.Add_MouseEnter({
        param($sender, $eventArgs)
        $sender.BackColor = $sender.Tag.HoverColor
    })
    $Button.Add_MouseLeave({
        param($sender, $eventArgs)
        $sender.BackColor = $sender.Tag.BackColor
    })
    $Button.Add_Resize({
        param($sender, $eventArgs)
        $sender.Region = New-RoundedRegion -Width $sender.Width -Height $sender.Height -Radius $sender.Tag.Radius
    })
    $Button.Add_Paint({
        param($sender, $eventArgs)
        if ($sender.Width -le 0 -or $sender.Height -le 0) { return }
        $eventArgs.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        if (-not $sender.Enabled) {
            # WinForms' default disabled caption is too dark on navy surfaces.
            $disabledBrush = New-Object System.Drawing.SolidBrush($uiPanel2)
            $disabledPen = New-Object System.Drawing.Pen($uiBorderSoft, 1)
            $eventArgs.Graphics.FillRectangle($disabledBrush, $sender.ClientRectangle)
            $eventArgs.Graphics.DrawRectangle($disabledPen, 0, 0, ($sender.Width - 1), ($sender.Height - 1))
            $captionFlags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::SingleLine
            [System.Windows.Forms.TextRenderer]::DrawText($eventArgs.Graphics, $sender.Text, $sender.Font, $sender.ClientRectangle, $uiMuted, $captionFlags)
            $disabledBrush.Dispose()
            $disabledPen.Dispose()
        }
        $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(38, 255, 255, 255), 1)
        $eventArgs.Graphics.DrawLine($linePen, 8, 1, ($sender.Width - 9), 1)
        $linePen.Dispose()
    })
}

function New-RoundedRegion {
    param([int]$Width, [int]$Height, [int]$Radius)
    $diameter = [Math]::Max(2, $Radius * 2)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $diameter, $diameter, 180, 90)
    $path.AddArc(($Width - $diameter - 1), 0, $diameter, $diameter, 270, 90)
    $path.AddArc(($Width - $diameter - 1), ($Height - $diameter - 1), $diameter, $diameter, 0, 90)
    $path.AddArc(0, ($Height - $diameter - 1), $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return (New-Object System.Drawing.Region -ArgumentList $path)
}

function Set-ModernGroup {
    param([System.Windows.Forms.GroupBox]$Group)
    $Group.BackColor = $uiPanel
    $Group.ForeColor = $uiAccent
    $Group.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
}

function Set-ModernInput {
    param([System.Windows.Forms.Control]$Control)
    $Control.BackColor = $uiInput
    $Control.ForeColor = $uiText
    if ($Control -is [System.Windows.Forms.TextBox] -or $Control -is [System.Windows.Forms.ListBox]) {
        $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
    if ($Control -is [System.Windows.Forms.ComboBox]) {
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    }
}

function Set-ModernButtonIfPresent {
    param(
        [string]$Name,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor = $uiButtonText,
        [int]$BorderSize = 1,
        [int]$Radius = 10
    )

    $buttonVar = Get-Variable -Name $Name -ErrorAction SilentlyContinue
    if ($buttonVar -and $buttonVar.Value -is [System.Windows.Forms.Button]) {
        Set-ModernButton -Button $buttonVar.Value -BackColor $BackColor -ForeColor $ForeColor -BorderSize $BorderSize -Radius $Radius
    }
}

function Set-ButtonPaddingTree {
    param([System.Windows.Forms.Control]$Root)

    foreach ($control in $Root.Controls) {
        if ($control -is [System.Windows.Forms.Button]) {
            $control.Padding = New-Object System.Windows.Forms.Padding(6, 2, 6, 2)
        }

        if ($control.Controls.Count -gt 0) {
            Set-ButtonPaddingTree -Root $control
        }
    }
}

function Set-AccessibleControlIfPresent {
    param(
        [string]$Name,
        [string]$AccessibleName,
        [string]$AccessibleDescription = "",
        [Nullable[int]]$TabIndex = $null
    )

    $controlVar = Get-Variable -Name $Name -ErrorAction SilentlyContinue
    if ($controlVar -and $controlVar.Value -is [System.Windows.Forms.Control]) {
        $control = $controlVar.Value
        $control.AccessibleName = $AccessibleName
        if (-not [string]::IsNullOrWhiteSpace($AccessibleDescription)) {
            $control.AccessibleDescription = $AccessibleDescription
        }
        if ($null -ne $TabIndex) {
            $control.TabIndex = [int]$TabIndex
        }
    }
}

function Set-AccessibilityTree {
    param([System.Windows.Forms.Control]$Root)

    foreach ($control in $Root.Controls) {
        if ($control -is [System.Windows.Forms.Label]) {
            $control.TabStop = $false
        }
        elseif ($control -is [System.Windows.Forms.Button]) {
            $control.AccessibleRole = [System.Windows.Forms.AccessibleRole]::PushButton
            if ([string]::IsNullOrWhiteSpace($control.AccessibleName)) {
                $control.AccessibleName = $control.Text
            }
            if ($control.FlatAppearance.BorderSize -lt 1) {
                $control.FlatAppearance.BorderSize = 1
            }
        }
        elseif ($control -is [System.Windows.Forms.CheckBox]) {
            $control.AccessibleRole = [System.Windows.Forms.AccessibleRole]::CheckButton
            if ([string]::IsNullOrWhiteSpace($control.AccessibleName)) {
                $control.AccessibleName = $control.Text
            }
        }
        elseif ($control -is [System.Windows.Forms.TextBox]) {
            $control.AccessibleRole = [System.Windows.Forms.AccessibleRole]::Text
        }
        elseif ($control -is [System.Windows.Forms.ComboBox]) {
            $control.AccessibleRole = [System.Windows.Forms.AccessibleRole]::DropList
        }
        elseif ($control -is [System.Windows.Forms.ListView] -or $control -is [System.Windows.Forms.DataGridView]) {
            if ([string]::IsNullOrWhiteSpace($control.AccessibleName)) {
                $control.AccessibleName = $control.Name
            }
        }

        if ($control.Controls.Count -gt 0) {
            Set-AccessibilityTree -Root $control
        }
    }
}

function Set-DarkControlTree {
    param([System.Windows.Forms.Control]$Root)

    foreach ($control in $Root.Controls) {
        if ($control -is [System.Windows.Forms.TabPage]) {
            $control.BackColor = $uiPanel
            $control.ForeColor = $uiText
        }
        elseif ($control -is [System.Windows.Forms.GroupBox]) {
            Set-ModernGroup -Group $control
        }
        elseif ($control -is [System.Windows.Forms.Label]) {
            $control.BackColor = [System.Drawing.Color]::Transparent
            $control.ForeColor = $uiText
        }
        elseif ($control -is [System.Windows.Forms.CheckBox]) {
            $control.BackColor = $uiPanel
            $control.ForeColor = $uiText
        }
        elseif ($control -is [System.Windows.Forms.TextBox] -or
                $control -is [System.Windows.Forms.ComboBox] -or
                $control -is [System.Windows.Forms.NumericUpDown] -or
                $control -is [System.Windows.Forms.ListBox]) {
            Set-ModernInput -Control $control
        }
        elseif ($control -is [System.Windows.Forms.ListView]) {
            $control.BackColor = $uiInput
            $control.ForeColor = $uiText
            $control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        }
        elseif ($control -is [System.Windows.Forms.DataGridView]) {
            $control.BackgroundColor = $uiInput
            $control.GridColor = $uiBorder
            $control.ForeColor = $uiText
            $control.DefaultCellStyle.BackColor = $uiInput
            $control.DefaultCellStyle.ForeColor = $uiText
            $control.ColumnHeadersDefaultCellStyle.BackColor = $uiPanel2
            $control.ColumnHeadersDefaultCellStyle.ForeColor = $uiText
            $control.DefaultCellStyle.SelectionBackColor = $uiAccent2
            $control.DefaultCellStyle.SelectionForeColor = $uiButtonText
            $control.AlternatingRowsDefaultCellStyle.BackColor = $uiInput2
            $control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $control.EnableHeadersVisualStyles = $false
        }
        elseif ($control -is [System.Windows.Forms.Button]) {
            Set-ModernButton -Button $control -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1
        }
        elseif ($control -is [System.Windows.Forms.ProgressBar]) {
            $control.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
            $control.BackColor = $uiInput
            $control.ForeColor = $uiAccent2
        }

        if ($control.Controls.Count -gt 0) {
            Set-DarkControlTree -Root $control
        }
    }
}

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(24, 13)
$titleLabel.Size = New-Object System.Drawing.Size(400, 30)
$titleLabel.Text = "VM Deletion Console"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = $uiText
$form.Controls.Add($titleLabel)

$btnUnlockSettings = New-Object System.Windows.Forms.Button
$btnUnlockSettings.Location = New-Object System.Drawing.Point(606, 16)
$btnUnlockSettings.Size = New-Object System.Drawing.Size(170, 34)
$btnUnlockSettings.Text = "Unlock settings"
$btnUnlockSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnUnlockSettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnUnlockSettings.FlatAppearance.BorderSize = 0
$btnUnlockSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnUnlockSettings)

$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Location = New-Object System.Drawing.Point(500, 23)
$versionLabel.Size = New-Object System.Drawing.Size(60, 20)
$versionLabel.Text = "v1.0"
$versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$versionLabel.ForeColor = $uiMuted
$versionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($versionLabel)
#endregion

#region Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(20, 82)
$tabControl.Size = New-Object System.Drawing.Size(764, ($form.ClientSize.Height - 100))
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabControl.ItemSize = New-Object System.Drawing.Size(187, 42)
$tabControl.Add_DrawItem({
    param($sender, $eventArgs)

    $tabPage = $sender.TabPages[$eventArgs.Index]
    $isSelected = ($eventArgs.Index -eq $sender.SelectedIndex)
    $bounds = $sender.GetTabRect($eventArgs.Index)
    $tabRect = New-Object System.Drawing.Rectangle(($bounds.X + 2), ($bounds.Y + 3), ($bounds.Width - 4), ($bounds.Height - 4))

    if ($isSelected) {
        $backColor = $uiAccent2
        $textColor = $uiButtonText
    }
    else {
        $backColor = $uiPanel2
        $textColor = $uiMuted
    }

    $eventArgs.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $tabBackRect = New-Object System.Drawing.Rectangle($bounds.X, $bounds.Y, $bounds.Width, $bounds.Height)
    $tabBackBrush = New-Object System.Drawing.SolidBrush($uiPanel)
    $backBrush = New-Object System.Drawing.SolidBrush($backColor)
    $borderPen = New-Object System.Drawing.Pen($(if ($isSelected) { $uiAccent } else { $uiBorderSoft }), 1)
    $textFlags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor
                 [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
                 [System.Windows.Forms.TextFormatFlags]::EndEllipsis

    $eventArgs.Graphics.FillRectangle($tabBackBrush, $tabBackRect)
    $eventArgs.Graphics.FillRectangle($backBrush, $tabRect)
    $eventArgs.Graphics.DrawRectangle($borderPen, $tabRect)
    [System.Windows.Forms.TextRenderer]::DrawText($eventArgs.Graphics, $tabPage.Text, $sender.Font, $tabRect, $textColor, $textFlags)

    $tabBackBrush.Dispose()
    $backBrush.Dispose()
    $borderPen.Dispose()
})
$form.Controls.Add($tabControl)

$tabMain = New-Object System.Windows.Forms.TabPage
$tabMain.Text = "1  VM Deletion"
$tabMain.BackColor = $uiPanel
$tabMain.AutoScroll = $true
$tabControl.Controls.Add($tabMain)

$tabArchive = New-Object System.Windows.Forms.TabPage
$tabArchive.Enabled = $false
$tabArchive.Text = "2  Archive"
$tabArchive.BackColor = $uiPanel
$tabArchive.AutoScroll = $true
$tabControl.Controls.Add($tabArchive)

$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Enabled = $false
$tabSettings.Text = "3  Settings"
$tabSettings.BackColor = $uiPanel
$tabSettings.AutoScroll = $true
$tabControl.Controls.Add($tabSettings)

$tabLogs = New-Object System.Windows.Forms.TabPage
$tabLogs.Text = "4  Activity logs"
$tabLogs.BackColor = $uiPanel
$tabLogs.AutoScroll = $false
$tabControl.Controls.Add($tabLogs)
$script:logsTab = $tabLogs
#endregion

#region Logs Tab Controls
$logsGroup = New-Object System.Windows.Forms.GroupBox
$logsGroup.Location = New-Object System.Drawing.Point(15, 15)
$logsGroup.Size = New-Object System.Drawing.Size(600, 740)
$logsGroup.Text = "Error Details"
$tabLogs.Controls.Add($logsGroup)

$txtLogs = New-Object System.Windows.Forms.TextBox
$txtLogs.Location = New-Object System.Drawing.Point(15, 28)
$txtLogs.Size = New-Object System.Drawing.Size(570, 655)
$txtLogs.Multiline = $true
$txtLogs.ReadOnly = $true
$txtLogs.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$txtLogs.WordWrap = $false
$txtLogs.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLogs.BackColor = $uiInput
$txtLogs.ForeColor = $uiText
$txtLogs.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logsGroup.Controls.Add($txtLogs)
$script:txtLogs = $txtLogs

$btnClearLogs = New-Object System.Windows.Forms.Button
$btnClearLogs.Location = New-Object System.Drawing.Point(485, 695)
$btnClearLogs.Size = New-Object System.Drawing.Size(100, 30)
$btnClearLogs.Text = "Clear"
$btnClearLogs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnClearLogs.FlatStyle = "Flat"
$btnClearLogs.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClearLogs.Add_Click({ $script:txtLogs.Clear() })
$logsGroup.Controls.Add($btnClearLogs)
#endregion
#region Archive Tab Controls

# SQL Connection Settings
$sqlGroup = New-Object System.Windows.Forms.GroupBox
$sqlGroup.Location = New-Object System.Drawing.Point(15, 15)
$sqlGroup.Size = New-Object System.Drawing.Size(600, 130)
$sqlGroup.Text = "SQL Database Connection"
$tabArchive.Controls.Add($sqlGroup)

$lblSQLServer = New-Object System.Windows.Forms.Label
$lblSQLServer.Location = New-Object System.Drawing.Point(15, 28)
$lblSQLServer.Size = New-Object System.Drawing.Size(90, 20)
$lblSQLServer.Text = "SQL Server:"
$sqlGroup.Controls.Add($lblSQLServer)

$txtSQLServer = New-Object System.Windows.Forms.TextBox
$txtSQLServer.Location = New-Object System.Drawing.Point(110, 25)
$txtSQLServer.Size = New-Object System.Drawing.Size(200, 23)
$txtSQLServer.Text = $script:DefaultSettings.SQLServer
$sqlGroup.Controls.Add($txtSQLServer)

$lblSQLDatabase = New-Object System.Windows.Forms.Label
$lblSQLDatabase.Location = New-Object System.Drawing.Point(325, 28)
$lblSQLDatabase.Size = New-Object System.Drawing.Size(80, 20)
$lblSQLDatabase.Text = "Database:"
$sqlGroup.Controls.Add($lblSQLDatabase)

$txtSQLDatabase = New-Object System.Windows.Forms.TextBox
$txtSQLDatabase.Location = New-Object System.Drawing.Point(410, 25)
$txtSQLDatabase.Size = New-Object System.Drawing.Size(160, 23)
$txtSQLDatabase.Text = $script:DefaultSettings.SQLDatabase
$sqlGroup.Controls.Add($txtSQLDatabase)

$lblSQLTable = New-Object System.Windows.Forms.Label
$lblSQLTable.Location = New-Object System.Drawing.Point(15, 58)
$lblSQLTable.Size = New-Object System.Drawing.Size(90, 20)
$lblSQLTable.Text = "Table Name:"
$sqlGroup.Controls.Add($lblSQLTable)

$txtSQLTable = New-Object System.Windows.Forms.TextBox
$txtSQLTable.Location = New-Object System.Drawing.Point(110, 55)
$txtSQLTable.Size = New-Object System.Drawing.Size(200, 23)
$txtSQLTable.Text = $($script:DefaultSettings.SQLTable)
$sqlGroup.Controls.Add($txtSQLTable)

$chkSQLIntegrated = New-Object System.Windows.Forms.CheckBox
$chkSQLIntegrated.Location = New-Object System.Drawing.Point(325, 57)
$chkSQLIntegrated.Size = New-Object System.Drawing.Size(225, 20)
$chkSQLIntegrated.Text = "Windows Authentication"
$chkSQLIntegrated.Checked = $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($chkSQLIntegrated)

$lblSQLUser = New-Object System.Windows.Forms.Label
$lblSQLUser.Location = New-Object System.Drawing.Point(15, 88)
$lblSQLUser.Size = New-Object System.Drawing.Size(85, 20)
$lblSQLUser.Text = "Username:"
$lblSQLUser.Enabled = -not $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($lblSQLUser)

$txtSQLUser = New-Object System.Windows.Forms.TextBox
$txtSQLUser.Location = New-Object System.Drawing.Point(110, 85)
$txtSQLUser.Size = New-Object System.Drawing.Size(145, 23)
$txtSQLUser.Text = $script:DefaultSettings.SQLUsername
$txtSQLUser.Enabled = -not $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($txtSQLUser)

$lblSQLPass = New-Object System.Windows.Forms.Label
$lblSQLPass.Location = New-Object System.Drawing.Point(270, 88)
$lblSQLPass.Size = New-Object System.Drawing.Size(80, 20)
$lblSQLPass.Text = "Password:"
$lblSQLPass.Enabled = -not $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($lblSQLPass)

$txtSQLPass = New-Object System.Windows.Forms.TextBox
$txtSQLPass.Location = New-Object System.Drawing.Point(360, 85)
$txtSQLPass.Size = New-Object System.Drawing.Size(95, 23)
$txtSQLPass.UseSystemPasswordChar = $true
$txtSQLPass.Text = $script:DefaultSettings.SQLPassword
$txtSQLPass.Enabled = -not $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($txtSQLPass)

$btnCheckTable = New-Object System.Windows.Forms.Button
$btnCheckTable.Location = New-Object System.Drawing.Point(470, 83)
$btnCheckTable.Size = New-Object System.Drawing.Size(110, 30)
$btnCheckTable.Text = "Check Table"
$btnCheckTable.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
$btnCheckTable.ForeColor = [System.Drawing.Color]::White
$btnCheckTable.FlatStyle = "Flat"
$btnCheckTable.Cursor = [System.Windows.Forms.Cursors]::Hand
$sqlGroup.Controls.Add($btnCheckTable)

# Progress Area
$archiveProgressGroup = New-Object System.Windows.Forms.GroupBox
$archiveProgressGroup.Location = New-Object System.Drawing.Point(15, 155)
$archiveProgressGroup.Size = New-Object System.Drawing.Size(600, 82)
$archiveProgressGroup.Text = "Status"
$tabArchive.Controls.Add($archiveProgressGroup)

$lblArchiveStatus = New-Object System.Windows.Forms.TextBox
$lblArchiveStatus.Location = New-Object System.Drawing.Point(15, 25)
$lblArchiveStatus.Size = New-Object System.Drawing.Size(570, 46)
$lblArchiveStatus.Text = "Click 'Check Table' to verify database connection and table status"
$lblArchiveStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblArchiveStatus.Multiline = $true
$lblArchiveStatus.ReadOnly = $true
$lblArchiveStatus.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$lblArchiveStatus.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$archiveProgressGroup.Controls.Add($lblArchiveStatus)

# Data Grid
$archiveDataGroup = New-Object System.Windows.Forms.GroupBox
$archiveDataGroup.Location = New-Object System.Drawing.Point(15, 247)
$archiveDataGroup.Size = New-Object System.Drawing.Size(600, 453)
$archiveDataGroup.Text = "Deletion Archive Data"
$tabArchive.Controls.Add($archiveDataGroup)

$dgvArchive = New-Object System.Windows.Forms.DataGridView
$dgvArchive.Location = New-Object System.Drawing.Point(10, 20)
$dgvArchive.Size = New-Object System.Drawing.Size(580, 388)
$dgvArchive.AllowUserToAddRows = $false
$dgvArchive.AllowUserToDeleteRows = $false
$dgvArchive.ReadOnly = $true

# ========== שינויים חדשים - גלילה אופקית ==========
$dgvArchive.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dgvArchive.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$dgvArchive.AllowUserToResizeColumns = $true
# =================================================

$dgvArchive.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgvArchive.BackgroundColor = $uiInput
$dgvArchive.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$dgvArchive.RowHeadersVisible = $false
$archiveDataGroup.Controls.Add($dgvArchive)

$lblRecordCount = New-Object System.Windows.Forms.Label
$lblRecordCount.Location = New-Object System.Drawing.Point(10, 418)
$lblRecordCount.Size = New-Object System.Drawing.Size(230, 20)
$lblRecordCount.Text = "0 records"
$lblRecordCount.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblRecordCount.ForeColor = $uiMuted
$archiveDataGroup.Controls.Add($lblRecordCount)

$btnRefreshArchive = New-Object System.Windows.Forms.Button
$btnRefreshArchive.Location = New-Object System.Drawing.Point(284, 413)
$btnRefreshArchive.Size = New-Object System.Drawing.Size(90, 30)
$btnRefreshArchive.Text = "Refresh"
$btnRefreshArchive.FlatStyle = "Flat"
$btnRefreshArchive.Enabled = $false
$btnRefreshArchive.Cursor = [System.Windows.Forms.Cursors]::Hand
$archiveDataGroup.Controls.Add($btnRefreshArchive)

$btnExportCSV = New-Object System.Windows.Forms.Button
$btnExportCSV.Location = New-Object System.Drawing.Point(384, 413)
$btnExportCSV.Size = New-Object System.Drawing.Size(92, 30)
$btnExportCSV.Text = "Export CSV"
$btnExportCSV.FlatStyle = "Flat"
$btnExportCSV.Enabled = $false
$btnExportCSV.Cursor = [System.Windows.Forms.Cursors]::Hand
$archiveDataGroup.Controls.Add($btnExportCSV)

$btnExportExcel = New-Object System.Windows.Forms.Button
$btnExportExcel.Location = New-Object System.Drawing.Point(486, 413)
$btnExportExcel.Size = New-Object System.Drawing.Size(104, 30)
$btnExportExcel.Text = "Export Excel"
$btnExportExcel.FlatStyle = "Flat"
$btnExportExcel.Enabled = $false
$btnExportExcel.Cursor = [System.Windows.Forms.Cursors]::Hand
$archiveDataGroup.Controls.Add($btnExportExcel)

$btnSaveSQLSettings = New-Object System.Windows.Forms.Button
$btnSaveSQLSettings.Location = New-Object System.Drawing.Point(510, 715)
$btnSaveSQLSettings.Size = New-Object System.Drawing.Size(105, 30)
$btnSaveSQLSettings.Text = "Save"
$btnSaveSQLSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSaveSQLSettings.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnSaveSQLSettings.ForeColor = [System.Drawing.Color]::White
$btnSaveSQLSettings.FlatStyle = "Flat"
$btnSaveSQLSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$tabArchive.Controls.Add($btnSaveSQLSettings)

#endregion

#region Settings Tab Controls
# SCCM Settings
$sccmSettingsGroup = New-Object System.Windows.Forms.GroupBox
$sccmSettingsGroup.Location = New-Object System.Drawing.Point(15, 15)
$sccmSettingsGroup.Size = New-Object System.Drawing.Size(600, 70)
$sccmSettingsGroup.Text = "SCCM Settings"
$tabSettings.Controls.Add($sccmSettingsGroup)

$lblSCCMServer = New-Object System.Windows.Forms.Label
$lblSCCMServer.Location = New-Object System.Drawing.Point(15, 28)
$lblSCCMServer.Size = New-Object System.Drawing.Size(100, 20)
$lblSCCMServer.Text = "SCCM Server:"
$sccmSettingsGroup.Controls.Add($lblSCCMServer)

$txtSettingsSCCM = New-Object System.Windows.Forms.TextBox
$txtSettingsSCCM.Location = New-Object System.Drawing.Point(120, 25)
$txtSettingsSCCM.Size = New-Object System.Drawing.Size(220, 23)
$txtSettingsSCCM.Text = $script:DefaultSettings.SCCMServer
$sccmSettingsGroup.Controls.Add($txtSettingsSCCM)

$lblSiteCode = New-Object System.Windows.Forms.Label
$lblSiteCode.Location = New-Object System.Drawing.Point(350, 28)
$lblSiteCode.Size = New-Object System.Drawing.Size(45, 20)
$lblSiteCode.Text = "Site:"
$sccmSettingsGroup.Controls.Add($lblSiteCode)

$txtSettingsSiteCode = New-Object System.Windows.Forms.TextBox
$txtSettingsSiteCode.Location = New-Object System.Drawing.Point(405, 25)
$txtSettingsSiteCode.Size = New-Object System.Drawing.Size(55, 23)
$txtSettingsSiteCode.MaxLength = 3
$txtSettingsSiteCode.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper
$txtSettingsSiteCode.Text = $script:DefaultSettings.SCCMSiteCode
$sccmSettingsGroup.Controls.Add($txtSettingsSiteCode)

# DNS Settings
$dnsSettingsGroup = New-Object System.Windows.Forms.GroupBox
$dnsSettingsGroup.Location = New-Object System.Drawing.Point(15, 95)
$dnsSettingsGroup.Size = New-Object System.Drawing.Size(600, 70)
$dnsSettingsGroup.Text = "DNS Settings"
$tabSettings.Controls.Add($dnsSettingsGroup)

$lblDNSServer = New-Object System.Windows.Forms.Label
$lblDNSServer.Location = New-Object System.Drawing.Point(15, 28)
$lblDNSServer.Size = New-Object System.Drawing.Size(90, 20)
$lblDNSServer.Text = "DNS Server:"
$dnsSettingsGroup.Controls.Add($lblDNSServer)

$txtSettingsDNS = New-Object System.Windows.Forms.TextBox
$txtSettingsDNS.Location = New-Object System.Drawing.Point(115, 25)
$txtSettingsDNS.Size = New-Object System.Drawing.Size(160, 23)
$txtSettingsDNS.Text = $script:DefaultSettings.DNSServer
$dnsSettingsGroup.Controls.Add($txtSettingsDNS)

$lblDNSZone = New-Object System.Windows.Forms.Label
$lblDNSZone.Location = New-Object System.Drawing.Point(285, 28)
$lblDNSZone.Size = New-Object System.Drawing.Size(45, 20)
$lblDNSZone.Text = "DNS:"
$dnsSettingsGroup.Controls.Add($lblDNSZone)

$txtSettingsDNSZone = New-Object System.Windows.Forms.TextBox
$txtSettingsDNSZone.Location = New-Object System.Drawing.Point(335, 25)
$txtSettingsDNSZone.Size = New-Object System.Drawing.Size(135, 23)
$txtSettingsDNSZone.Text = $script:DefaultSettings.DNSZone
$dnsSettingsGroup.Controls.Add($txtSettingsDNSZone)

# AD Settings
$adSettingsGroup = New-Object System.Windows.Forms.GroupBox
$adSettingsGroup.Location = New-Object System.Drawing.Point(15, 175)
$adSettingsGroup.Size = New-Object System.Drawing.Size(600, 100)
$adSettingsGroup.Text = "Active Directory Settings"
$tabSettings.Controls.Add($adSettingsGroup)

$lblDC = New-Object System.Windows.Forms.Label
$lblDC.Location = New-Object System.Drawing.Point(15, 28)
$lblDC.Size = New-Object System.Drawing.Size(135, 20)
$lblDC.Text = "Domain Controller:"
$adSettingsGroup.Controls.Add($lblDC)

$txtSettingsDC = New-Object System.Windows.Forms.TextBox
$txtSettingsDC.Location = New-Object System.Drawing.Point(165, 25)
$txtSettingsDC.Size = New-Object System.Drawing.Size(300, 23)
$txtSettingsDC.Text = $script:DefaultSettings.DomainController
$adSettingsGroup.Controls.Add($txtSettingsDC)

$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Location = New-Object System.Drawing.Point(15, 58)
$lblEmail.Size = New-Object System.Drawing.Size(150, 20)
$lblEmail.Text = "Notification Email:"
$adSettingsGroup.Controls.Add($lblEmail)

$txtSettingsEmail = New-Object System.Windows.Forms.TextBox
$txtSettingsEmail.Location = New-Object System.Drawing.Point(175, 55)
$txtSettingsEmail.Size = New-Object System.Drawing.Size(285, 23)
$txtSettingsEmail.Text = $script:DefaultSettings.EmailNotification
$adSettingsGroup.Controls.Add($txtSettingsEmail)


# General Options
$optionsGroup = New-Object System.Windows.Forms.GroupBox
$optionsGroup.Location = New-Object System.Drawing.Point(15, 285)
$optionsGroup.Size = New-Object System.Drawing.Size(600, 130)
$optionsGroup.Text = "General Options"
$tabSettings.Controls.Add($optionsGroup)

$chkEnableLogging = New-Object System.Windows.Forms.CheckBox
$chkEnableLogging.Location = New-Object System.Drawing.Point(15, 25)
$chkEnableLogging.Size = New-Object System.Drawing.Size(150, 20)
$chkEnableLogging.Text = "Enable Logging"
$chkEnableLogging.Checked = $script:DefaultSettings.EnableLogging
$optionsGroup.Controls.Add($chkEnableLogging)

$chkConfirmDeletion = New-Object System.Windows.Forms.CheckBox
$chkConfirmDeletion.Location = New-Object System.Drawing.Point(200, 25)
$chkConfirmDeletion.Size = New-Object System.Drawing.Size(230, 20)
$chkConfirmDeletion.Text = "Require Deletion Confirmation"
$chkConfirmDeletion.Checked = $script:DefaultSettings.ConfirmDeletion
$optionsGroup.Controls.Add($chkConfirmDeletion)

$chkDeleteFromDisk = New-Object System.Windows.Forms.CheckBox
$chkDeleteFromDisk.Location = New-Object System.Drawing.Point(15, 50)
$chkDeleteFromDisk.Size = New-Object System.Drawing.Size(280, 20)
$chkDeleteFromDisk.Text = "Delete VM files from disk permanently"
$chkDeleteFromDisk.Checked = $script:DefaultSettings.DeleteFromDisk
$optionsGroup.Controls.Add($chkDeleteFromDisk)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Location = New-Object System.Drawing.Point(320, 50)
$lblTimeout.Size = New-Object System.Drawing.Size(135, 20)
$lblTimeout.Text = "Power Off Timeout:"
$optionsGroup.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(465, 48)
$numTimeout.Size = New-Object System.Drawing.Size(60, 23)
$numTimeout.Minimum = 10
$numTimeout.Maximum = 300
$numTimeout.Value = $script:DefaultSettings.PowerOffTimeout
$optionsGroup.Controls.Add($numTimeout)

$lblLogPath = New-Object System.Windows.Forms.Label
$lblLogPath.Location = New-Object System.Drawing.Point(15, 85)
$lblLogPath.Size = New-Object System.Drawing.Size(80, 20)
$lblLogPath.Text = "Log Path:"
$optionsGroup.Controls.Add($lblLogPath)

$txtLogPath = New-Object System.Windows.Forms.TextBox
$txtLogPath.Location = New-Object System.Drawing.Point(105, 82)
$txtLogPath.Size = New-Object System.Drawing.Size(370, 23)
$txtLogPath.Text = $script:DefaultSettings.LogPath
$optionsGroup.Controls.Add($txtLogPath)

$btnBrowseLog = New-Object System.Windows.Forms.Button
$btnBrowseLog.Location = New-Object System.Drawing.Point(485, 79)
$btnBrowseLog.Size = New-Object System.Drawing.Size(85, 30)
$btnBrowseLog.Text = "Browse..."
$btnBrowseLog.FlatStyle = "Flat"
$optionsGroup.Controls.Add($btnBrowseLog)

# Email Settings Group
$emailSettingsGroup = New-Object System.Windows.Forms.GroupBox
$emailSettingsGroup.Location = New-Object System.Drawing.Point(15, 425)
$emailSettingsGroup.Size = New-Object System.Drawing.Size(600, 275)
$emailSettingsGroup.Text = "Email Notification Settings"
$tabSettings.Controls.Add($emailSettingsGroup)

# Enable Email
$chkEnableEmail = New-Object System.Windows.Forms.CheckBox
$chkEnableEmail.Location = New-Object System.Drawing.Point(15, 22)
$chkEnableEmail.Size = New-Object System.Drawing.Size(205, 20)
$chkEnableEmail.Text = "Enable Email Notifications"
$chkEnableEmail.Checked = $script:DefaultSettings.EnableEmailNotification
$emailSettingsGroup.Controls.Add($chkEnableEmail)

# SMTP Row
$lblSMTPServer = New-Object System.Windows.Forms.Label
$lblSMTPServer.Location = New-Object System.Drawing.Point(15, 50)
$lblSMTPServer.Size = New-Object System.Drawing.Size(45, 20)
$lblSMTPServer.Text = "SMTP:"
$emailSettingsGroup.Controls.Add($lblSMTPServer)

$txtSMTPServer = New-Object System.Windows.Forms.TextBox
$txtSMTPServer.Location = New-Object System.Drawing.Point(65, 47)
$txtSMTPServer.Size = New-Object System.Drawing.Size(130, 23)
$txtSMTPServer.Text = $script:DefaultSettings.SMTPServer
$emailSettingsGroup.Controls.Add($txtSMTPServer)

$lblSMTPPort = New-Object System.Windows.Forms.Label
$lblSMTPPort.Location = New-Object System.Drawing.Point(200, 50)
$lblSMTPPort.Size = New-Object System.Drawing.Size(45, 20)
$lblSMTPPort.Text = "Port:"
$emailSettingsGroup.Controls.Add($lblSMTPPort)

$numSMTPPort = New-Object System.Windows.Forms.NumericUpDown
$numSMTPPort.Location = New-Object System.Drawing.Point(250, 47)
$numSMTPPort.Size = New-Object System.Drawing.Size(55, 23)
$numSMTPPort.Minimum = 1
$numSMTPPort.Maximum = 65535
$numSMTPPort.Value = $script:DefaultSettings.SMTPPort
$emailSettingsGroup.Controls.Add($numSMTPPort)

$chkSMTPSSL = New-Object System.Windows.Forms.CheckBox
$chkSMTPSSL.Location = New-Object System.Drawing.Point(315, 48)
$chkSMTPSSL.Size = New-Object System.Drawing.Size(55, 20)
$chkSMTPSSL.Text = "SSL"
$chkSMTPSSL.Checked = $script:DefaultSettings.SMTPUseSSL
$emailSettingsGroup.Controls.Add($chkSMTPSSL)

$lblSMTPFrom = New-Object System.Windows.Forms.Label
$lblSMTPFrom.Location = New-Object System.Drawing.Point(380, 50)
$lblSMTPFrom.Size = New-Object System.Drawing.Size(45, 20)
$lblSMTPFrom.Text = "From:"
$emailSettingsGroup.Controls.Add($lblSMTPFrom)

$txtSMTPFrom = New-Object System.Windows.Forms.TextBox
$txtSMTPFrom.Location = New-Object System.Drawing.Point(430, 47)
$txtSMTPFrom.Size = New-Object System.Drawing.Size(140, 23)
$txtSMTPFrom.Text = $script:DefaultSettings.SMTPFrom
$emailSettingsGroup.Controls.Add($txtSMTPFrom)

# Groups Section - Headers
$lblGroupsCommon = New-Object System.Windows.Forms.Label
$lblGroupsCommon.Location = New-Object System.Drawing.Point(15, 80)
$lblGroupsCommon.Size = New-Object System.Drawing.Size(180, 15)
$lblGroupsCommon.Text = "Other Groups (All VMs):"
$lblGroupsCommon.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$emailSettingsGroup.Controls.Add($lblGroupsCommon)

$lblGroupsWindows = New-Object System.Windows.Forms.Label
$lblGroupsWindows.Location = New-Object System.Drawing.Point(210, 80)
$lblGroupsWindows.Size = New-Object System.Drawing.Size(180, 15)
$lblGroupsWindows.Text = "Windows Only:"
$lblGroupsWindows.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblGroupsWindows.ForeColor = $uiAccent
$emailSettingsGroup.Controls.Add($lblGroupsWindows)

$lblGroupsLinux = New-Object System.Windows.Forms.Label
$lblGroupsLinux.Location = New-Object System.Drawing.Point(405, 80)
$lblGroupsLinux.Size = New-Object System.Drawing.Size(180, 15)
$lblGroupsLinux.Text = "Linux Only:"
$lblGroupsLinux.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblGroupsLinux.ForeColor = $uiWarning
$emailSettingsGroup.Controls.Add($lblGroupsLinux)

# ========== ListBoxes for Groups ==========

# Other Groups ListBox
$script:lstGroupsCommon = New-Object System.Windows.Forms.ListBox
$script:lstGroupsCommon.Name = "lstGroupsCommon"
$script:lstGroupsCommon.Location = New-Object System.Drawing.Point(15, 98)
$script:lstGroupsCommon.Size = New-Object System.Drawing.Size(180, 60)
$script:lstGroupsCommon.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lstGroupsCommon.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$script:lstGroupsCommon.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$emailSettingsGroup.Controls.Add($script:lstGroupsCommon)

# Windows Groups ListBox
$script:lstGroupsWindows = New-Object System.Windows.Forms.ListBox
$script:lstGroupsWindows.Name = "lstGroupsWindows"
$script:lstGroupsWindows.Location = New-Object System.Drawing.Point(210, 98)
$script:lstGroupsWindows.Size = New-Object System.Drawing.Size(180, 60)
$script:lstGroupsWindows.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lstGroupsWindows.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$script:lstGroupsWindows.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$emailSettingsGroup.Controls.Add($script:lstGroupsWindows)

# Linux Groups ListBox
$script:lstGroupsLinux = New-Object System.Windows.Forms.ListBox
$script:lstGroupsLinux.Name = "lstGroupsLinux"
$script:lstGroupsLinux.Location = New-Object System.Drawing.Point(405, 98)
$script:lstGroupsLinux.Size = New-Object System.Drawing.Size(180, 60)
$script:lstGroupsLinux.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lstGroupsLinux.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$script:lstGroupsLinux.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$emailSettingsGroup.Controls.Add($script:lstGroupsLinux)

# Add Group Controls
$lblAddGroup = New-Object System.Windows.Forms.Label
$lblAddGroup.Location = New-Object System.Drawing.Point(15, 202)
$lblAddGroup.Size = New-Object System.Drawing.Size(75, 20)
$lblAddGroup.Text = "Add:"
$emailSettingsGroup.Controls.Add($lblAddGroup)

$txtNewGroup = New-Object System.Windows.Forms.TextBox
$txtNewGroup.Location = New-Object System.Drawing.Point(90, 199)
$txtNewGroup.Size = New-Object System.Drawing.Size(245, 23)
$emailSettingsGroup.Controls.Add($txtNewGroup)

$cboGroupType = New-Object System.Windows.Forms.ComboBox
$cboGroupType.Location = New-Object System.Drawing.Point(345, 199)
$cboGroupType.Size = New-Object System.Drawing.Size(110, 23)
$cboGroupType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboGroupType.Items.AddRange(@("Other", "Windows", "Linux"))
$cboGroupType.SelectedIndex = 0
$emailSettingsGroup.Controls.Add($cboGroupType)

$btnAddGroup = New-Object System.Windows.Forms.Button
$btnAddGroup.Location = New-Object System.Drawing.Point(285, 230)
$btnAddGroup.Size = New-Object System.Drawing.Size(80, 30)
$btnAddGroup.Text = "Add"
$btnAddGroup.FlatStyle = "Flat"
$btnAddGroup.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnAddGroup.ForeColor = [System.Drawing.Color]::White
$btnAddGroup.Cursor = [System.Windows.Forms.Cursors]::Hand
$emailSettingsGroup.Controls.Add($btnAddGroup)

$btnRemoveGroup = New-Object System.Windows.Forms.Button
$btnRemoveGroup.Location = New-Object System.Drawing.Point(375, 230)
$btnRemoveGroup.Size = New-Object System.Drawing.Size(95, 30)
$btnRemoveGroup.Text = "Remove"
$btnRemoveGroup.FlatStyle = "Flat"
$btnRemoveGroup.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
$btnRemoveGroup.ForeColor = [System.Drawing.Color]::White
$btnRemoveGroup.Cursor = [System.Windows.Forms.Cursors]::Hand
$emailSettingsGroup.Controls.Add($btnRemoveGroup)

$btnTestEmail = New-Object System.Windows.Forms.Button
$btnTestEmail.Location = New-Object System.Drawing.Point(480, 230)
$btnTestEmail.Size = New-Object System.Drawing.Size(105, 30)
$btnTestEmail.Text = "Test Email"
$btnTestEmail.FlatStyle = "Flat"
$btnTestEmail.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
$btnTestEmail.ForeColor = [System.Drawing.Color]::White
$btnTestEmail.Cursor = [System.Windows.Forms.Cursors]::Hand
$emailSettingsGroup.Controls.Add($btnTestEmail)

# ========== טעינת קבוצות קיימות ==========
# Common Groups
if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsCommon)) {
    $groups = $script:DefaultSettings.GroupsCommon.Split(',')
    foreach ($group in $groups) {
        $g = $group.Trim()
        if (-not [string]::IsNullOrWhiteSpace($g)) {
            [void]$script:lstGroupsCommon.Items.Add($g)
        }
    }
}

# Windows Groups
if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsWindows)) {
    $groups = $script:DefaultSettings.GroupsWindows.Split(',')
    foreach ($group in $groups) {
        $g = $group.Trim()
        if (-not [string]::IsNullOrWhiteSpace($g)) {
            [void]$script:lstGroupsWindows.Items.Add($g)
        }
    }
}

# Linux Groups
if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsLinux)) {
    $groups = $script:DefaultSettings.GroupsLinux.Split(',')
    foreach ($group in $groups) {
        $g = $group.Trim()
        if (-not [string]::IsNullOrWhiteSpace($g)) {
            [void]$script:lstGroupsLinux.Items.Add($g)
        }
    }
}


# Settings Buttons
$btnSaveSettings = New-Object System.Windows.Forms.Button
$btnSaveSettings.Location = New-Object System.Drawing.Point(510, 715)
$btnSaveSettings.Size = New-Object System.Drawing.Size(100, 30)
$btnSaveSettings.Text = "Save"
$btnSaveSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSaveSettings.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnSaveSettings.ForeColor = [System.Drawing.Color]::White
$btnSaveSettings.FlatStyle = "Flat"
$btnSaveSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$tabSettings.Controls.Add($btnSaveSettings)

$btnResetSettings = New-Object System.Windows.Forms.Button
$btnResetSettings.Location = New-Object System.Drawing.Point(380, 715)
$btnResetSettings.Size = New-Object System.Drawing.Size(120, 30)
$btnResetSettings.Text = "Reset Settings"
$btnResetSettings.FlatStyle = "Flat"
$btnResetSettings.BackColor = [System.Drawing.Color]::FromArgb(255, 193, 7)
$btnResetSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$tabSettings.Controls.Add($btnResetSettings)

$lblSettingsStatus = New-Object System.Windows.Forms.Label
$lblSettingsStatus.Location = New-Object System.Drawing.Point(15, 720)
$lblSettingsStatus.Size = New-Object System.Drawing.Size(350, 20)
$lblSettingsStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$tabSettings.Controls.Add($lblSettingsStatus)
#endregion

#region Main Tab Controls

# vCenter Connection
$vcGroup = New-Object System.Windows.Forms.GroupBox
$vcGroup.Location = New-Object System.Drawing.Point(15, 15)
$vcGroup.Size = New-Object System.Drawing.Size(600, 105)
$vcGroup.Text = "vCenter Connection"
$tabMain.Controls.Add($vcGroup)

$lblVC = New-Object System.Windows.Forms.Label
$lblVC.Location = New-Object System.Drawing.Point(15, 28)
$lblVC.Size = New-Object System.Drawing.Size(115, 20)
$lblVC.Text = "vCenter Server:"
$vcGroup.Controls.Add($lblVC)

$cboVCServer = New-Object System.Windows.Forms.ComboBox
$cboVCServer.Location = New-Object System.Drawing.Point(135, 25)
$cboVCServer.Size = New-Object System.Drawing.Size(325, 25)
$cboVCServer.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboVCServer.Font = New-Object System.Drawing.Font("Segoe UI", 9)

foreach ($vc in $script:vCenterServers) {
    $cboVCServer.Items.Add($vc.Name.ToUpper()) | Out-Null
}

$selectedIndex = 0
    for ($i = 0; $i -lt $script:vCenterServers.Count; $i++) {
    if ($script:vCenterServers[$i].Server -eq $script:Settings.vCenterServer) {
        $selectedIndex = $i
        break
    }
}
$cboVCServer.SelectedIndex = $selectedIndex
$vcGroup.Controls.Add($cboVCServer)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(475, 22)
$btnConnect.Size = New-Object System.Drawing.Size(110, 30)
$btnConnect.Text = "Connect"
Set-ModernButton -Button $btnConnect -BackColor $uiAccent -ForeColor ([System.Drawing.Color]::White)
$vcGroup.Controls.Add($btnConnect)

$lblConnectionStatus = New-Object System.Windows.Forms.Label
$lblConnectionStatus.Location = New-Object System.Drawing.Point(15, 55)
$lblConnectionStatus.Size = New-Object System.Drawing.Size(570, 36)
$lblConnectionStatus.Text = "Not connected"
$lblConnectionStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblConnectionStatus.ForeColor = $uiMuted
$vcGroup.Controls.Add($lblConnectionStatus)
Set-ModernGroup -Group $vcGroup

# VM Selection
$vmGroup = New-Object System.Windows.Forms.GroupBox
$vmGroup.Location = New-Object System.Drawing.Point(15, 130)
$vmGroup.Size = New-Object System.Drawing.Size(600, 235)
$vmGroup.Text = "Virtual Machine Selection (Hold Ctrl/Shift for multi-select)"
$tabMain.Controls.Add($vmGroup)

# === שורת הפילטרים ===
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Location = New-Object System.Drawing.Point(15, 25)
$lblFilter.Size = New-Object System.Drawing.Size(60, 20)
$lblFilter.Text = "Filter:"
$vmGroup.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(80, 22)
$txtFilter.Size = New-Object System.Drawing.Size(95, 23)
$txtFilter.Enabled = $false
$vmGroup.Controls.Add($txtFilter)

$lblOSFilter = New-Object System.Windows.Forms.Label
$lblOSFilter.Location = New-Object System.Drawing.Point(185, 25)
$lblOSFilter.Size = New-Object System.Drawing.Size(35, 20)
$lblOSFilter.Text = "OS:"
$vmGroup.Controls.Add($lblOSFilter)

$cboOSFilter = New-Object System.Windows.Forms.ComboBox
$cboOSFilter.Location = New-Object System.Drawing.Point(225, 22)
$cboOSFilter.Size = New-Object System.Drawing.Size(90, 23)
$cboOSFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboOSFilter.Items.AddRange(@("All", "Windows", "Linux", "Other"))
$cboOSFilter.SelectedIndex = 0
$cboOSFilter.Enabled = $false
$vmGroup.Controls.Add($cboOSFilter)

# Include Templates Checkbox
$chkIncludeTemplates = New-Object System.Windows.Forms.CheckBox
$chkIncludeTemplates.Location = New-Object System.Drawing.Point(505, 177)
$chkIncludeTemplates.Size = New-Object System.Drawing.Size(85, 20)
$chkIncludeTemplates.Text = "Templates"
$chkIncludeTemplates.Font = New-Object System.Drawing.Font("Segoe UI", 7)
$chkIncludeTemplates.ForeColor = $uiAccent
$chkIncludeTemplates.Checked = $false
$chkIncludeTemplates.Enabled = $false
$vmGroup.Controls.Add($chkIncludeTemplates)

$lblVMCount = New-Object System.Windows.Forms.Label
$lblVMCount.Location = New-Object System.Drawing.Point(320, 25)
$lblVMCount.Size = New-Object System.Drawing.Size(80, 20)
$lblVMCount.Text = "0 VMs"
$lblVMCount.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblVMCount.ForeColor = $uiMuted
$vmGroup.Controls.Add($lblVMCount)

$lblSelectedCount = New-Object System.Windows.Forms.Label
$lblSelectedCount.Location = New-Object System.Drawing.Point(410, 25)
$lblSelectedCount.Size = New-Object System.Drawing.Size(80, 20)
$lblSelectedCount.Text = "0 VM selected"
$lblSelectedCount.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblSelectedCount.ForeColor = $uiAccent
$vmGroup.Controls.Add($lblSelectedCount)

# === ListView - עם CheckBoxes ===
$lstVMs = New-Object System.Windows.Forms.ListView
$lstVMs.Location = New-Object System.Drawing.Point(15, 50)
$lstVMs.Size = New-Object System.Drawing.Size(495, 120)
$lstVMs.View = [System.Windows.Forms.View]::Details
$lstVMs.FullRowSelect = $true
$lstVMs.GridLines = $true
$lstVMs.CheckBoxes = $true  # הפעלת CheckBoxes
$lstVMs.MultiSelect = $true
$lstVMs.HideSelection = $false
$lstVMs.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lstVMs.Enabled = $true
$lstVMs.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$lstVMs.BackColor = $uiInput
$lstVMs.ForeColor = $uiText

# עמודות
$colName = New-Object System.Windows.Forms.ColumnHeader
$colName.Text = "VM Name"
$colName.Width = 180

$colPowerState = New-Object System.Windows.Forms.ColumnHeader
$colPowerState.Text = "Power State"
$colPowerState.Width = 80

$colOS = New-Object System.Windows.Forms.ColumnHeader
$colOS.Text = "OS Type"
$colOS.Width = 70

$colOSFull = New-Object System.Windows.Forms.ColumnHeader
$colOSFull.Text = "Operating System"
$colOSFull.Width = 165

$lstVMs.Columns.AddRange(@($colName, $colPowerState, $colOS, $colOSFull))
$vmGroup.Controls.Add($lstVMs)

# === כפתורים ליד הטבלה ===

# Refresh Button
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Location = New-Object System.Drawing.Point(520, 50)
$btnRefresh.Size = New-Object System.Drawing.Size(70, 30)
$btnRefresh.Text = "Refresh"
$btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnRefresh.ForeColor = $uiSuccess
$btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefresh.FlatAppearance.BorderSize = 1
$btnRefresh.FlatAppearance.BorderColor = $uiSuccess
$btnRefresh.Enabled = $false
$btnRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
$vmGroup.Controls.Add($btnRefresh)

$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($btnRefresh, "Refresh VM List")

# Clear Selection Button
$btnClearSelection = New-Object System.Windows.Forms.Button
$btnClearSelection.Location = New-Object System.Drawing.Point(520, 88)
$btnClearSelection.Size = New-Object System.Drawing.Size(70, 30)
$btnClearSelection.Text = "Clear"
$btnClearSelection.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnClearSelection.ForeColor = $uiDanger
$btnClearSelection.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClearSelection.FlatAppearance.BorderSize = 1
$btnClearSelection.FlatAppearance.BorderColor = $uiDanger
$btnClearSelection.Enabled = $false
$btnClearSelection.Cursor = [System.Windows.Forms.Cursors]::Hand
$vmGroup.Controls.Add($btnClearSelection)

$tooltip.SetToolTip($btnClearSelection, "Clear Selection")

# Load from CSV Button
$btnLoadCSV = New-Object System.Windows.Forms.Button
$btnLoadCSV.Location = New-Object System.Drawing.Point(520, 126)
$btnLoadCSV.Size = New-Object System.Drawing.Size(70, 30)
$btnLoadCSV.Text = "CSV"
$btnLoadCSV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnLoadCSV.ForeColor = $uiAccent
$btnLoadCSV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnLoadCSV.FlatAppearance.BorderSize = 1
$btnLoadCSV.FlatAppearance.BorderColor = $uiAccent
$btnLoadCSV.Enabled = $false
$btnLoadCSV.Cursor = [System.Windows.Forms.Cursors]::Hand
$vmGroup.Controls.Add($btnLoadCSV)

$tooltip.SetToolTip($btnLoadCSV, "Load VM list from CSV file")

$tooltip.SetToolTip($chkIncludeTemplates, "Include VM Templates in the list")


# === Reason - שורה תחתונה ===
$lblReason = New-Object System.Windows.Forms.Label
$lblReason.Location = New-Object System.Drawing.Point(15, 180)
$lblReason.Size = New-Object System.Drawing.Size(65, 20)
$lblReason.Text = "Reason:"
$vmGroup.Controls.Add($lblReason)

$txtReason = New-Object System.Windows.Forms.TextBox
$txtReason.Location = New-Object System.Drawing.Point(90, 177)
$txtReason.Size = New-Object System.Drawing.Size(400, 23)
$txtReason.Enabled = $false
$txtReason.MaxLength = 500
$txtReason.Text = "Activity Ended"
$vmGroup.Controls.Add($txtReason)

$lblReasonHint = New-Object System.Windows.Forms.Label
$lblReasonHint.Location = New-Object System.Drawing.Point(90, 202)
$lblReasonHint.Size = New-Object System.Drawing.Size(380, 18)
$lblReasonHint.Text = "Enter reason for deletion (saved to archive &email)"
$lblReasonHint.Font = New-Object System.Drawing.Font("Segoe UI", 7)
$lblReasonHint.ForeColor = $uiMuted
$vmGroup.Controls.Add($lblReasonHint)
Set-ModernGroup -Group $vmGroup

# Deletion Options
$deleteOptionsGroup = New-Object System.Windows.Forms.GroupBox
$deleteOptionsGroup.Location = New-Object System.Drawing.Point(15, 375)
$deleteOptionsGroup.Size = New-Object System.Drawing.Size(600, 55)
$deleteOptionsGroup.Text = "Deletion Options"
$tabMain.Controls.Add($deleteOptionsGroup)

$chkSCCM = New-Object System.Windows.Forms.CheckBox
$chkSCCM.Location = New-Object System.Drawing.Point(20, 20)
$chkSCCM.Size = New-Object System.Drawing.Size(80, 22)
$chkSCCM.Text = "SCCM"
$chkSCCM.Checked = $true
$deleteOptionsGroup.Controls.Add($chkSCCM)

$chkAD = New-Object System.Windows.Forms.CheckBox
$chkAD.Location = New-Object System.Drawing.Point(115, 20)
$chkAD.Size = New-Object System.Drawing.Size(80, 22)
$chkAD.Text = "AD"
$chkAD.Checked = $true
$deleteOptionsGroup.Controls.Add($chkAD)

$chkDNS = New-Object System.Windows.Forms.CheckBox
$chkDNS.Location = New-Object System.Drawing.Point(210, 20)
$chkDNS.Size = New-Object System.Drawing.Size(80, 22)
$chkDNS.Text = "DNS"
$chkDNS.Checked = $true
$deleteOptionsGroup.Controls.Add($chkDNS)

$chkVM = New-Object System.Windows.Forms.CheckBox
$chkVM.Location = New-Object System.Drawing.Point(305, 20)
$chkVM.Size = New-Object System.Drawing.Size(80, 22)
$chkVM.Text = "VM"
$chkVM.Checked = $true
$deleteOptionsGroup.Controls.Add($chkVM)

$chkArchive = New-Object System.Windows.Forms.CheckBox
$chkArchive.Location = New-Object System.Drawing.Point(400, 20)
$chkArchive.Size = New-Object System.Drawing.Size(80, 22)
$chkArchive.Text = "Archive"
$chkArchive.Checked = $true
$deleteOptionsGroup.Controls.Add($chkArchive)

$chkEmail = New-Object System.Windows.Forms.CheckBox
$chkEmail.Location = New-Object System.Drawing.Point(495, 20)
$chkEmail.Size = New-Object System.Drawing.Size(80, 22)
$chkEmail.Text = "Email"
$chkEmail.Checked = $true
$deleteOptionsGroup.Controls.Add($chkEmail)
$script:PersistentStepControls=@{}
foreach($definition in (Get-DeletionStepDefinitions)){
    if($definition.Name -eq 'Email'){$preference=$chkEnableEmail}
    else{
        $preference=New-Object System.Windows.Forms.CheckBox
        $preference.Name='persistent'+$definition.Name
        $preference.Visible=$false
        $tabSettings.Controls.Add($preference)
    }
    $preference.Checked=Test-DeletionStepEnabled -Step $definition.Name
    $script:PersistentStepControls[$definition.Name]=$preference
}
Sync-DeletionStepControls -ResetSelection
Set-ModernGroup -Group $deleteOptionsGroup

# Progress
$progressGroup = New-Object System.Windows.Forms.GroupBox
$progressGroup.Location = New-Object System.Drawing.Point(15, 440)
$progressGroup.Size = New-Object System.Drawing.Size(600, 260)
$progressGroup.Text = "Deletion Progress"
$tabMain.Controls.Add($progressGroup)


$lblVMProgress = New-Object System.Windows.Forms.Label
$lblVMProgress.Location = New-Object System.Drawing.Point(10, 25)
$lblVMProgress.Size = New-Object System.Drawing.Size(500, 20)
$lblVMProgress.Text = "VM Progress: Idle"
$lblVMProgress.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblVMProgress.ForeColor = $uiAccent

$progressGroup.Controls.Add($lblVMProgress)


foreach ($task in $tasks) {
    $idx = $tasks.IndexOf($task)
    
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(15, $yPos)
    $statusLabel.Size = New-Object System.Drawing.Size(46, 22)
    $statusLabel.Text = "WAIT"
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $statusLabel.ForeColor = $uiMuted
    $progressGroup.Controls.Add($statusLabel)
    $statusLabels += $statusLabel
    
    $taskLabel = New-Object System.Windows.Forms.Label
    $taskLabel.Location = New-Object System.Drawing.Point(68, $yPos)
    $taskLabel.Size = New-Object System.Drawing.Size(177, 22)
    $taskLabel.Text = $task
    $taskLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $taskLabel.ForeColor = $uiText
    $progressGroup.Controls.Add($taskLabel)
    
    $resultLabel = New-Object System.Windows.Forms.Label
    $resultLabel.Location = New-Object System.Drawing.Point(255, $yPos)
    $resultLabel.Size = New-Object System.Drawing.Size(330, 22)
    $resultLabel.Text = "Pending"
    $resultLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $resultLabel.ForeColor = $uiMuted
    $progressGroup.Controls.Add($resultLabel)
    $resultLabels += $resultLabel
    
    $yPos += 25
}

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 225)
$progressBar.Size = New-Object System.Drawing.Size(570, 20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressGroup.Controls.Add($progressBar)
Set-ModernGroup -Group $progressGroup

# Action Buttons
$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Location = New-Object System.Drawing.Point(510, 715)
$btnDelete.Size = New-Object System.Drawing.Size(100, 30)
$btnDelete.Text = "Delete VM"
$btnDelete.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnDelete.Enabled = $false
Set-ModernButton -Button $btnDelete -BackColor $uiDanger -ForeColor ([System.Drawing.Color]::White)
$tabMain.Controls.Add($btnDelete)

# כפתור Resume
$btnResume = New-Object System.Windows.Forms.Button
$btnResume.Location = New-Object System.Drawing.Point(400, 715)
$btnResume.Size = New-Object System.Drawing.Size(100, 30)
$btnResume.Text = "Resume"
$btnResume.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
Set-ModernButton -Button $btnResume -BackColor $uiWarning -ForeColor $uiText
$tabMain.Controls.Add($btnResume)

# בדוק אם יש קובץ progress והפעל את הכפתור בהתאם
$progressFile = "$env:TEMP\VMDeletionTool\progress.json"
$btnResume.Enabled = [bool](Wait-JsonFileAvailable -Path $progressFile)

$tooltip.SetToolTip($btnResume, "Resume from last failed VM")

# כפתור Close
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(510, 675)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$btnClose.Text = "Close"
$btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnClose.FlatStyle = "Flat"
$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClose.Visible = $false

Set-DarkControlTree -Root $form
Set-ButtonPaddingTree -Root $form
Set-AccessibilityTree -Root $form
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$versionLabel.ForeColor = $uiMuted
$lblVMCount.ForeColor = $uiMuted
$lblSelectedCount.ForeColor = $uiAccent
$lblReasonHint.ForeColor = $uiMuted
$lstVMs.BackColor = $uiInput
$lstVMs.ForeColor = $uiText
$tabControl.BackColor = $uiPanel
$tabControl.ForeColor = $uiText
$progressBar.ForeColor = $uiAccent2
$progressBar.BackColor = $uiInput

Set-ModernButtonIfPresent -Name "btnUnlockSettings" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 12
Set-ModernButtonIfPresent -Name "btnConnect" -BackColor $uiAccent2 -ForeColor $uiButtonText -BorderSize 0 -Radius 12
Set-ModernButtonIfPresent -Name "btnCheckTable" -BackColor $uiAccent2 -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnRefreshArchive" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnExportCSV" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnExportExcel" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnSaveSQLSettings" -BackColor $uiSuccess -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnClearLogs" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnBrowseLog" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnAddGroup" -BackColor $uiSuccess -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnRemoveGroup" -BackColor $uiDanger -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnTestEmail" -BackColor $uiAccent2 -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnSaveSettings" -BackColor $uiSuccess -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnResetSettings" -BackColor $uiWarning -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnRefresh" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnClearSelection" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnLoadCSV" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnDelete" -BackColor $uiDanger -ForeColor $uiButtonText -BorderSize 0 -Radius 12
Set-ModernButtonIfPresent -Name "btnResume" -BackColor $uiAccent -ForeColor $uiButtonText -BorderSize 0 -Radius 12
Set-ModernButtonIfPresent -Name "btnClose" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-AccessibleControlIfPresent -Name "btnUnlockSettings" -AccessibleName "Unlock settings" -AccessibleDescription "Unlocks the Archive and Settings tabs." -TabIndex 0
Set-AccessibleControlIfPresent -Name "cboVCServer" -AccessibleName "vCenter Server" -AccessibleDescription "Select the vCenter environment to connect to." -TabIndex 1
Set-AccessibleControlIfPresent -Name "btnConnect" -AccessibleName "Connect" -AccessibleDescription "Connect to the selected vCenter server." -TabIndex 2
Set-AccessibleControlIfPresent -Name "txtFilter" -AccessibleName "VM filter" -AccessibleDescription "Filter the virtual machine list." -TabIndex 3
Set-AccessibleControlIfPresent -Name "cboOSFilter" -AccessibleName "OS filter" -AccessibleDescription "Filter virtual machines by operating system." -TabIndex 4
Set-AccessibleControlIfPresent -Name "lstVMs" -AccessibleName "Virtual machine selection list" -AccessibleDescription "Select one or more VMs or templates for deletion." -TabIndex 5
Set-AccessibleControlIfPresent -Name "btnRefresh" -AccessibleName "Refresh VM list" -AccessibleDescription "Reload the VM list from vCenter." -TabIndex 6
Set-AccessibleControlIfPresent -Name "btnClearSelection" -AccessibleName "Clear VM selection" -AccessibleDescription "Clear the selected VMs." -TabIndex 7
Set-AccessibleControlIfPresent -Name "btnLoadCSV" -AccessibleName "Load VM list from CSV" -AccessibleDescription "Import VM names from a CSV file." -TabIndex 8
Set-AccessibleControlIfPresent -Name "txtReason" -AccessibleName "Deletion reason" -AccessibleDescription "Reason saved to archive and email notification." -TabIndex 9
Set-AccessibleControlIfPresent -Name "chkIncludeTemplates" -AccessibleName "Include templates" -AccessibleDescription "Include VM templates in the selection list." -TabIndex 10
Set-AccessibleControlIfPresent -Name "chkSCCM" -AccessibleName "Delete SCCM record" -TabIndex 11
Set-AccessibleControlIfPresent -Name "chkAD" -AccessibleName "Delete AD record" -TabIndex 12
Set-AccessibleControlIfPresent -Name "chkDNS" -AccessibleName "Delete DNS record" -TabIndex 13
Set-AccessibleControlIfPresent -Name "chkVM" -AccessibleName "Delete virtual machine" -TabIndex 14
Set-AccessibleControlIfPresent -Name "chkArchive" -AccessibleName "Archive deletion result" -TabIndex 15
Set-AccessibleControlIfPresent -Name "chkEmail" -AccessibleName "Send deletion email" -TabIndex 16
Set-AccessibleControlIfPresent -Name "btnResume" -AccessibleName "Resume deletion" -AccessibleDescription "Resume deletion from the last failed VM when progress data exists." -TabIndex 17
Set-AccessibleControlIfPresent -Name "btnDelete" -AccessibleName "Delete selected VMs" -AccessibleDescription "Start the deletion workflow for selected VMs." -TabIndex 18
Set-AccessibleControlIfPresent -Name "btnClose" -AccessibleName "Close" -AccessibleDescription "Close the tool after the workflow is complete." -TabIndex 19

Set-AccessibleControlIfPresent -Name "txtLogs" -AccessibleName "Error details log" -AccessibleDescription "Full error details collected from the tool."
Set-AccessibleControlIfPresent -Name "btnClearLogs" -AccessibleName "Clear logs" -AccessibleDescription "Clear the on-screen LOGS tab text."
Set-AccessibleControlIfPresent -Name "txtSQLServer" -AccessibleName "SQL server"
Set-AccessibleControlIfPresent -Name "txtSQLDatabase" -AccessibleName "SQL database"
Set-AccessibleControlIfPresent -Name "txtSQLTable" -AccessibleName "SQL table"
Set-AccessibleControlIfPresent -Name "chkSQLIntegrated" -AccessibleName "Use Windows authentication for SQL"
Set-AccessibleControlIfPresent -Name "txtSQLUser" -AccessibleName "SQL username"
Set-AccessibleControlIfPresent -Name "txtSQLPass" -AccessibleName "SQL password"
Set-AccessibleControlIfPresent -Name "btnCheckTable" -AccessibleName "Check SQL table"
Set-AccessibleControlIfPresent -Name "dgvArchive" -AccessibleName "Deletion archive data table"
Set-AccessibleControlIfPresent -Name "btnRefreshArchive" -AccessibleName "Refresh archive"
Set-AccessibleControlIfPresent -Name "btnExportCSV" -AccessibleName "Export archive to CSV"
Set-AccessibleControlIfPresent -Name "btnExportExcel" -AccessibleName "Export archive to Excel"
Set-AccessibleControlIfPresent -Name "btnSaveSQLSettings" -AccessibleName "Save SQL settings"
Set-AccessibleControlIfPresent -Name "txtSettingsSCCM" -AccessibleName "SCCM server"
Set-AccessibleControlIfPresent -Name "txtSettingsSiteCode" -AccessibleName "SCCM site code"
Set-AccessibleControlIfPresent -Name "txtSettingsDNS" -AccessibleName "DNS server"
Set-AccessibleControlIfPresent -Name "txtSettingsDNSZone" -AccessibleName "DNS zone"
Set-AccessibleControlIfPresent -Name "txtSettingsDC" -AccessibleName "Domain controller"
Set-AccessibleControlIfPresent -Name "txtSettingsEmail" -AccessibleName "Notification email"
Set-AccessibleControlIfPresent -Name "chkEnableLogging" -AccessibleName "Enable logging"
Set-AccessibleControlIfPresent -Name "chkConfirmDeletion" -AccessibleName "Require deletion confirmation"
Set-AccessibleControlIfPresent -Name "chkDeleteFromDisk" -AccessibleName "Delete VM files from disk permanently"
Set-AccessibleControlIfPresent -Name "numTimeout" -AccessibleName "Power off timeout seconds"
Set-AccessibleControlIfPresent -Name "txtLogPath" -AccessibleName "Log path"
Set-AccessibleControlIfPresent -Name "btnBrowseLog" -AccessibleName "Browse log path"
Set-AccessibleControlIfPresent -Name "chkEnableEmail" -AccessibleName "Enable email notifications"
Set-AccessibleControlIfPresent -Name "txtSMTPServer" -AccessibleName "SMTP server"
Set-AccessibleControlIfPresent -Name "numSMTPPort" -AccessibleName "SMTP port"
Set-AccessibleControlIfPresent -Name "chkSMTPSSL" -AccessibleName "Use SMTP SSL"
Set-AccessibleControlIfPresent -Name "txtSMTPFrom" -AccessibleName "SMTP from address"
Set-AccessibleControlIfPresent -Name "txtNewGroup" -AccessibleName "New email group or address"
Set-AccessibleControlIfPresent -Name "cboGroupType" -AccessibleName "Email group type"
Set-AccessibleControlIfPresent -Name "btnAddGroup" -AccessibleName "Add email group"
Set-AccessibleControlIfPresent -Name "btnRemoveGroup" -AccessibleName "Remove email group"
Set-AccessibleControlIfPresent -Name "btnTestEmail" -AccessibleName "Send test email"
Set-AccessibleControlIfPresent -Name "btnResetSettings" -AccessibleName "Reset settings"
Set-AccessibleControlIfPresent -Name "btnSaveSettings" -AccessibleName "Save settings"
$lblGroupsWindows.ForeColor = $uiAccent
$lblGroupsLinux.ForeColor = $uiWarning
Set-AccessibilityTree -Root $form
#endregion

# The visible single-screen WPF wizard is initialized after the existing handlers.

$script:VMFilterTimer = New-Object System.Windows.Forms.Timer
$script:VMFilterTimer.Interval = 250
$script:PendingVMFilterText = ""
$script:PendingVMOSFilter = "All"
$script:LastAppliedVMFilterKey = ""
$script:VMFilterTimer.Add_Tick({
    $script:VMFilterTimer.Stop()
    Filter-VMList -FilterText $script:PendingVMFilterText -OSFilter $script:PendingVMOSFilter
})
# ========== Event handler לבחירה ברשימה ==========
$lstVMs.Add_ItemChecked({
    param($sender, $e)
    
    if ($script:IsUpdatingList) { return }
    
    $vmName = $e.Item.Text
    
    if ($e.Item.Checked) {
        if ($script:SelectedVMNames -notcontains $vmName) {
            [void]$script:SelectedVMNames.Add($vmName)
        }
    }
    else {
        $script:SelectedVMNames.Remove($vmName)
    }
    
    $lblSelectedCount.Text = "$($script:SelectedVMNames.Count) selected"
    
})


# ========== Email Groups Event Handlers ==========

# משתנה לעקוב אחרי הרשימה הנבחרת
$script:selectedGroupList = $null

# Click Events - לזהות איזו רשימה נבחרה
$script:lstGroupsCommon.Add_MouseDown({
    $script:selectedGroupList = "Other"
    $script:lstGroupsWindows.ClearSelected()
    $script:lstGroupsLinux.ClearSelected()
})

$script:lstGroupsWindows.Add_MouseDown({
    $script:selectedGroupList = "Windows"
    $script:lstGroupsCommon.ClearSelected()
    $script:lstGroupsLinux.ClearSelected()
})

$script:lstGroupsLinux.Add_MouseDown({
    $script:selectedGroupList = "Linux"
    $script:lstGroupsCommon.ClearSelected()
    $script:lstGroupsWindows.ClearSelected()
})

# Add Group Button
$btnAddGroup.Add_Click({
    $groupName = $txtNewGroup.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($groupName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a group name.",
            "Validation",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $txtNewGroup.Focus()
        return
    }
    
    # בחר את הרשימה לפי ה-ComboBox
    $selectedType = $cboGroupType.SelectedItem.ToString()
    $targetList = $null
    
    switch ($selectedType) {
        "Other"   { $targetList = $script:lstGroupsCommon }
        "Windows" { $targetList = $script:lstGroupsWindows }
        "Linux"   { $targetList = $script:lstGroupsLinux }
    }
    
    if ($null -eq $targetList) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a group type.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    
    # בדיקה אם הקבוצה כבר קיימת
    $exists = $false
    for ($i = 0; $i -lt $targetList.Items.Count; $i++) {
        if ($targetList.Items[$i].ToString() -eq $groupName) {
            $exists = $true
            break
        }
    }
    
    if ($exists) {
        [System.Windows.Forms.MessageBox]::Show(
            "Group '$groupName' already exists in $selectedType list.",
            "Duplicate",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    # הוסף את הקבוצה
    [void]$targetList.Items.Add($groupName)
    $txtNewGroup.Text = ""
    $txtNewGroup.Focus()
    
    Write-Log "Added group '$groupName' to $selectedType list"
})

# Remove Group Button
$btnRemoveGroup.Add_Click({
    # מצא איזו רשימה יש בה פריט נבחר
    $targetList = $null
    $listName = ""
    
    if ($script:lstGroupsCommon.SelectedIndex -ge 0) {
        $targetList = $script:lstGroupsCommon
        $listName = "Other"
    }
    elseif ($script:lstGroupsWindows.SelectedIndex -ge 0) {
        $targetList = $script:lstGroupsWindows
        $listName = "Windows"
    }
    elseif ($script:lstGroupsLinux.SelectedIndex -ge 0) {
        $targetList = $script:lstGroupsLinux
        $listName = "Linux"
    }
    
    if ($null -eq $targetList -or $targetList.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please click on a group in one of the lists to select it, then click Remove.",
            "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $selectedIndex = $targetList.SelectedIndex
    $groupName = $targetList.Items[$selectedIndex].ToString()
    
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Remove group '$groupName' from $listName list?",
        "Confirm Remove",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $targetList.Items.RemoveAt($selectedIndex)
        Write-Log "Removed group '$groupName' from $listName list"
    }
})

# Test Email Button
$btnTestEmail.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSMTPServer.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please configure SMTP server first.",
            "SMTP Not Configured",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($txtSettingsEmail.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please configure Notification Email in AD Settings first.",
            "Email Not Configured",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $lblSettingsStatus.Text = "Sending test email..."
    $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
$testBody = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            padding: 0; 
            margin: 0; 
            background-color: #e9ecef; 
        }
        .container { 
            max-width: 500px; 
            margin: 30px auto; 
            background-color: #ffffff; 
            border-radius: 10px; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.1); 
            overflow: hidden; 
        }
        .header { 
            background: linear-gradient(135deg, #0066cc 0%, #004499 100%); 
            color: white; 
            padding: 25px; 
            text-align: center; 
        }
        .header h2 { 
            margin: 0; 
            font-size: 22px; 
            font-weight: 600; 
        }
        .icon { 
            font-size: 40px; 
            margin-bottom: 10px; 
        }
        .content { 
            padding: 30px; 
            text-align: center; 
            color: #333333; 
        }
        .success-badge { 
            background-color: #d4edda; 
            color: #155724; 
            padding: 12px 20px; 
            border-radius: 25px; 
            display: inline-block; 
            margin-bottom: 15px; 
            font-weight: 500; 
        }
        .message { 
            color: #666666; 
            font-size: 14px; 
            margin-bottom: 25px; 
        }
        .details { 
            background-color: #f8f9fa; 
            border-radius: 8px; 
            padding: 20px; 
            text-align: left; 
        }
        .detail-row { 
            display: flex; 
            justify-content: space-between; 
            padding: 8px 0; 
            border-bottom: 1px solid #e9ecef; 
        }
        .detail-row:last-child { 
            border-bottom: none; 
        }
        .detail-label { 
            color: #888888; 
            font-size: 13px; 
        }
        .detail-value { 
            color: #333333; 
            font-weight: 500; 
            font-size: 13px; 
        }
        .footer { 
            background-color: #f8f9fa; 
            padding: 15px; 
            text-align: center; 
            font-size: 11px; 
            color: #999999; 
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="icon"></div>
            <h2>VM Deletion Tool</h2>
        </div>
        <div class="content">
            <div class="success-badge">Test Email</div>
            <p class="message">This is a test email from <strong>VM Deletion Tool</strong>.<br>If you received this email, your SMTP settings are configured correctly.</p>
            <div class="details">
                <div class="detail-row">
                    <span class="detail-label">Sent at</span>
                    <span class="detail-value">$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Sent by</span>
                    <span class="detail-value">$env:USERDOMAIN\$env:USERNAME</span>
                </div>
            </div>
        </div>
        <div class="footer">
            VM Deletion Tool - IT Infrastructure Management
        </div>
    </div>
</body>
</html>
"@
        
        $mailParams = @{
            From       = $txtSMTPFrom.Text
            To         = $txtSettingsEmail.Text
            Subject    = "[VM Deletion Tool] Test Email - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            Body       = $testBody
            BodyAsHtml = $true
            SmtpServer = $txtSMTPServer.Text
            Port       = [int]$numSMTPPort.Value
            Priority   = "Normal"
        }
        
        if ($chkSMTPSSL.Checked) {
            $mailParams.UseSsl = $true
        }
        
        Send-MailMessage @mailParams -ErrorAction Stop
        
        $lblSettingsStatus.Text = "OK - Test email sent successfully to $($txtSettingsEmail.Text)"
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        
        [System.Windows.Forms.MessageBox]::Show(
            "Test email sent successfully!`n`nCheck inbox: $($txtSettingsEmail.Text)",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        $errorMessage = "Failed to send test email: $($_.Exception.Message)"
        $lblSettingsStatus.Text = "Failed - Email failed. See LOGS."
        $lblSettingsStatus.ForeColor = $uiDanger
        Write-Log $errorMessage -Level "ERROR"
    }
})


$btnUnlockSettings.Add_Click({
    # If tabs are currently enabled, lock them
    if ($tabArchive.Enabled) {
        $tabArchive.Enabled = $false
        $tabSettings.Enabled = $false
        $btnUnlockSettings.Text = "Unlock settings"
        $btnUnlockSettings.ForeColor = [System.Drawing.Color]::Black
    }
    else {
        # If tabs are disabled, show password prompt
        $passwordForm = New-Object System.Windows.Forms.Form
        $passwordForm.Text = "Unlock Settings"
        $passwordForm.Size = New-Object System.Drawing.Size(250, 150)
        $passwordForm.StartPosition = "CenterScreen"
        $passwordForm.FormBorderStyle = "FixedDialog"
        $passwordForm.MaximizeBox = $false
        $passwordForm.TopMost = $true

        $lblPassword = New-Object System.Windows.Forms.Label
        $lblPassword.Location = New-Object System.Drawing.Point(20, 20)
        $lblPassword.Size = New-Object System.Drawing.Size(200, 20)
        $lblPassword.Text = "Enter Settings Password:"
        $passwordForm.Controls.Add($lblPassword)

        $txtPasswordInput = New-Object System.Windows.Forms.TextBox
        $txtPasswordInput.Location = New-Object System.Drawing.Point(20, 50)
        $txtPasswordInput.Size = New-Object System.Drawing.Size(200, 23)
        $txtPasswordInput.UseSystemPasswordChar = $true
        $passwordForm.Controls.Add($txtPasswordInput)

        $btnSubmit = New-Object System.Windows.Forms.Button
        $btnSubmit.Location = New-Object System.Drawing.Point(20, 80)
        $btnSubmit.Size = New-Object System.Drawing.Size(90, 30)
        $btnSubmit.Text = "Unlock"
        $passwordForm.Controls.Add($btnSubmit)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(130, 80)
        $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
        $btnCancel.Text = "Cancel"
        $passwordForm.Controls.Add($btnCancel)

        $btnSubmit.Add_Click({
            if ($txtPasswordInput.Text -eq $script:settingsPassword) {
                $tabArchive.Enabled = $true
                $tabSettings.Enabled = $true
                $btnUnlockSettings.Text = "Lock settings"
                $btnUnlockSettings.ForeColor = [System.Drawing.Color]::Green
                $passwordForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Incorrect password", 
                    "Access Denied", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        })

        $btnCancel.Add_Click({ $passwordForm.Close() })
        $passwordForm.AcceptButton = $btnSubmit
        $passwordForm.CancelButton = $btnCancel

        $passwordForm.ShowDialog()
    }
})

$btnUnlockSettings.Add_DoubleClick({
    $tabArchive.Enabled = $false
    $tabSettings.Enabled = $false
    $btnUnlockSettings.Text = "Unlock settings"
    $btnUnlockSettings.ForeColor = [System.Drawing.Color]::Black

    [System.Windows.Forms.MessageBox]::Show(
        "Settings and Archive tabs have been locked.", 
        "Locked", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

# ========== vCenter Connect Button ==========
$btnConnect.Add_Click({
    if ($script:vcConnected) {
        # Disconnect
        try {
            Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue
            $script:vcConnected = $false
            $btnConnect.Text = "Connect"
            $btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
            $lblConnectionStatus.Text = "Disconnected"
            $lblConnectionStatus.ForeColor = $uiMuted
            $lstVMs.Items.Clear()
            $script:allVMs = @()
            $script:allVMsData = @()
            $lblVMCount.Text = "0 VMs"
            $txtFilter.Enabled = $false
            $cboOSFilter.Enabled = $false
            $btnRefresh.Enabled = $false
            $btnClearSelection.Enabled = $false
            $btnDelete.Enabled = $false
            $lstVMs.Enabled = $true
            $lstVMs.BackColor = $uiInput
            $lstVMs.ForeColor = $uiText
            $txtReason.Enabled = $false
            Write-Log "Disconnected from vCenter"
        } catch {
            Write-Log "Error disconnecting: $($_.Exception.Message)" -Level "ERROR"
        }
        return
    }
    
    # Connect
    $selectedIndex = $cboVCServer.SelectedIndex
    if ($selectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a vCenter server.", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $vcServer = $script:vCenterServers[$selectedIndex].Server
    $vcName = $script:vCenterServers[$selectedIndex].Name
    
    $lblConnectionStatus.Text = "Connecting to $vcName..."
    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnConnect.Text = 'Connecting...'
    $btnConnect.Enabled = $false
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    Sync-DeletionWizard
    if($script:WizardWindow){$script:WizardWindow.Dispatcher.Invoke([Action]{},[System.Windows.Threading.DispatcherPriority]::Render)}
    
    try {
        # Check if PowerCLI is loaded
        if (-not (Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) {
            Import-Module VMware.VimAutomation.Core -ErrorAction Stop
        }
        
        # Set PowerCLI configuration
        Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        
        # Connect to vCenter
        Connect-VIServer -Server $vcServer -ErrorAction Stop | Out-Null
        
        $script:vcConnected = $true
        $btnConnect.Text = "Disconnect"
        $btnConnect.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
        $lblConnectionStatus.Text = "Connected to $vcName ($vcServer)"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        
        # Enable controls
        $txtFilter.Enabled = $true
        $cboOSFilter.Enabled = $true
        $btnRefresh.Enabled = $true
        $btnClearSelection.Enabled = $true
        $btnLoadCSV.Enabled = $true
        $btnResume.Enabled = $true
        $btnDelete.Enabled = $true
        $lstVMs.Enabled = $true
        $txtReason.Enabled = $true
        $chkIncludeTemplates.Enabled = $true
        
        # Load VMs
        Get-VMList
        
        Write-Log "Connected to vCenter: $vcServer"
    }
    catch {
        $lblConnectionStatus.Text = "Connection failed: $($_.Exception.Message)"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
        if ($tooltip) {
            $tooltip.SetToolTip($lblConnectionStatus, "Connection failed: $($_.Exception.Message)")
        }
        Write-Log "vCenter connection failed: $($_.Exception.Message)" -Level "ERROR"
    }
    finally {
        $btnConnect.Text=$(if($script:vcConnected){'Disconnect'}else{'Connect'})
        $btnConnect.Enabled=$true
        Sync-DeletionWizard
    }
})

# ========== Filter TextBox ==========
$txtFilter.Add_TextChanged({
    Schedule-VMFilter
})

# ========== OS Filter ComboBox ==========
$cboOSFilter.Add_SelectedIndexChanged({
    Schedule-VMFilter
})

# ========== Include Templates Checkbox ==========
$chkIncludeTemplates.Add_CheckedChanged({
    $script:IncludeTemplates = $chkIncludeTemplates.Checked
    if ($script:vcConnected) {
        Get-VMList
    }
})

# ========== Refresh Button ==========
$btnRefresh.Add_Click({
    $lblConnectionStatus.Text = "Refreshing VM list..."
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    Get-VMList
    $lblConnectionStatus.Text = "VM list refreshed"
})
# ========== CSV List Button ==========
$btnLoadCSV.Add_Click({   
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV file with VM list"

    if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    try {
        $csvData = Import-Csv -Path $openFileDialog.FileName

        $hasVMName      = $csvData | Get-Member -Name "VMName"      -MemberType NoteProperty
        $hasVCenter     = $csvData | Get-Member -Name "vCenter"     -MemberType NoteProperty
        $hasEnvironment = $csvData | Get-Member -Name "Environment" -MemberType NoteProperty

        if (-not $hasVMName) {
            [System.Windows.Forms.MessageBox]::Show(
                "CSV file must contain a column named 'VMName'.",
                "Invalid CSV Format",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        # ===== Load all VM names once (fast) =====
        $vmHash = [System.Collections.Generic.HashSet[string]]::new()
        if ($script:vcConnected) {
            foreach ($vmData in $script:allVMsData) {
                $vmHash.Add($vmData.Name) | Out-Null
            }
        }

        # ===== Preview Form =====
        $previewForm = New-Object System.Windows.Forms.Form
        $previewForm.Text = "CSV Preview - Select VMs to Import"
        $previewForm.Size = New-Object System.Drawing.Size(760, 520)
        $previewForm.StartPosition = "CenterScreen"
        $previewForm.FormBorderStyle = "FixedDialog"
        $previewForm.MaximizeBox = $false
        $previewForm.TopMost = $true
        $previewForm.BackColor = $uiPanel

        # ===== DataGrid =====
        $dataGrid = New-Object System.Windows.Forms.DataGridView
        $dataGrid.Location = New-Object System.Drawing.Point(10, 10)
        $dataGrid.Size = New-Object System.Drawing.Size(720, 380)
        $dataGrid.AllowUserToAddRows = $false
        $dataGrid.AllowUserToDeleteRows = $false
        $dataGrid.RowHeadersVisible = $false
        $dataGrid.SelectionMode = "FullRowSelect"
        $dataGrid.MultiSelect = $false
        $dataGrid.AutoSizeColumnsMode = "Fill"
        $dataGrid.AllowUserToResizeRows = $false
        $dataGrid.BackgroundColor = $uiInput
        $dataGrid.GridColor = $uiBorderSoft
        $dataGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $dataGrid.EnableHeadersVisualStyles = $false
        $dataGrid.DefaultCellStyle.BackColor = $uiInput
        $dataGrid.DefaultCellStyle.ForeColor = $uiText
        $dataGrid.DefaultCellStyle.SelectionBackColor = $uiAccent2
        $dataGrid.DefaultCellStyle.SelectionForeColor = $uiButtonText
        $dataGrid.ColumnHeadersDefaultCellStyle.BackColor = $uiPanel2
        $dataGrid.ColumnHeadersDefaultCellStyle.ForeColor = $uiText
        $previewForm.Controls.Add($dataGrid)

        # ===== Columns =====
        $checkCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $checkCol.Name = "Select"
        $checkCol.HeaderText = "Select"
        $checkCol.Width = 40
        $checkCol.AutoSizeMode = "None"
        $dataGrid.Columns.Add($checkCol) | Out-Null

        $statusCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $statusCol.Name = "Status"
        $statusCol.HeaderText = "Status"
        $statusCol.Width = 90
        $statusCol.ReadOnly = $true
        $statusCol.AutoSizeMode = "None"
        $dataGrid.Columns.Add($statusCol) | Out-Null

        if ($hasVCenter) { 
            $dataGrid.Columns.Add("vCenter", "vCenter") | Out-Null 
        }
        if ($hasEnvironment) {
            $dataGrid.Columns.Add("Environment", "Environment") | Out-Null
        }
        $dataGrid.Columns.Add("VMName", "VM Name") | Out-Null

        # ===== Populate Rows =====
        $foundCount = 0
        $notFoundCount = 0
        
        foreach ($row in $csvData) {
            $vmName = $row.VMName.Trim()
            if ([string]::IsNullOrWhiteSpace($vmName)) { continue }

            $status = "Unknown"
            $exists = $false

            if ($script:vcConnected) {
                if ($vmHash.Contains($vmName)) {
                    $status = "Found"
                    $exists = $true
                    $foundCount++
                }
                else {
                    $status = "Not Found"
                    $notFoundCount++
                }
            }

            $rowData = @($exists, $status)  # רק Found מסומן כברירת מחדל
            if ($hasVCenter) { $rowData += $row.vCenter }
            if ($hasEnvironment) { $rowData += $row.Environment }
            $rowData += $vmName

            $index = $dataGrid.Rows.Add($rowData)

            # Coloring
            if ($script:vcConnected) {
                if ($exists) {
                    $dataGrid.Rows[$index].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(39, 72, 58)
                    $dataGrid.Rows[$index].DefaultCellStyle.ForeColor = $uiText
                }
                else {
                    $dataGrid.Rows[$index].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(82, 48, 55)
                    $dataGrid.Rows[$index].DefaultCellStyle.ForeColor = $uiMuted
                }
            }
        }

        # ===== Summary =====
        $lblSummary = New-Object System.Windows.Forms.Label
        $lblSummary.Location = New-Object System.Drawing.Point(10, 400)
        $lblSummary.Size = New-Object System.Drawing.Size(700, 20)
        $lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        if ($script:vcConnected) {
            $lblSummary.Text = "Total: $($dataGrid.Rows.Count) | Found: $foundCount | Not Found: $notFoundCount"
        }
        else {
            $lblSummary.Text = "Total VMs in CSV: $($dataGrid.Rows.Count) (Connect to vCenter to verify)"
        }
        $lblSummary.ForeColor = $uiMuted
        $previewForm.Controls.Add($lblSummary)

        # ===== Select All Button =====
        $btnSelectAll = New-Object System.Windows.Forms.Button
        $btnSelectAll.Location = New-Object System.Drawing.Point(10, 430)
        $btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
        $btnSelectAll.Text = "Select All"
        $btnSelectAll.FlatStyle = "Flat"
        Set-ModernButton -Button $btnSelectAll -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
        $btnSelectAll.Add_Click({
            foreach ($r in $dataGrid.Rows) {
                $r.Cells["Select"].Value = $true
            }
        })
        $previewForm.Controls.Add($btnSelectAll)

        # ===== Select None Button =====
        $btnSelectNone = New-Object System.Windows.Forms.Button
        $btnSelectNone.Location = New-Object System.Drawing.Point(120, 430)
        $btnSelectNone.Size = New-Object System.Drawing.Size(100, 30)
        $btnSelectNone.Text = "Select None"
        $btnSelectNone.FlatStyle = "Flat"
        Set-ModernButton -Button $btnSelectNone -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
        $btnSelectNone.Add_Click({
            foreach ($r in $dataGrid.Rows) {
                $r.Cells["Select"].Value = $false
            }
        })
        $previewForm.Controls.Add($btnSelectNone)

        # ===== Select Found Only Button =====
        $btnSelectFound = New-Object System.Windows.Forms.Button
        $btnSelectFound.Location = New-Object System.Drawing.Point(230, 430)
        $btnSelectFound.Size = New-Object System.Drawing.Size(140, 30)
        $btnSelectFound.Text = "Select Found Only"
        $btnSelectFound.FlatStyle = "Flat"
        Set-ModernButton -Button $btnSelectFound -BackColor $uiSuccess -ForeColor $uiButtonText -BorderSize 0 -Radius 10
        $btnSelectFound.Add_Click({
            foreach ($r in $dataGrid.Rows) {
                $r.Cells["Select"].Value = ($r.Cells["Status"].Value -eq "Found")
            }
        })
        $previewForm.Controls.Add($btnSelectFound)

        # ===== Import Button =====
        $btnImport = New-Object System.Windows.Forms.Button
        $btnImport.Location = New-Object System.Drawing.Point(520, 430)
        $btnImport.Size = New-Object System.Drawing.Size(100, 30)
        $btnImport.Text = "Import"
        $btnImport.FlatStyle = "Flat"
        Set-ModernButton -Button $btnImport -BackColor $uiAccent2 -ForeColor $uiButtonText -BorderSize 0 -Radius 10
        $btnImport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnImport.Add_Click({
            $script:SelectedVMNames.Clear()
            $script:CSVImportData = @()
            $importedCount = 0
            $skippedCount = 0

            foreach ($row in $dataGrid.Rows) {
                if ($row.Cells["Select"].Value -eq $true) {
                    $vmName = $row.Cells["VMName"].Value
                    if ([string]::IsNullOrWhiteSpace($vmName)) { continue }

                    # דלג על VMs שלא נמצאו
                    if ($script:vcConnected -and $row.Cells["Status"].Value -eq "Not Found") {
                        $skippedCount++
                        continue
                    }

                    $vmData = @{ VMName = $vmName }
                    if ($hasVCenter) {
                        $vmData.vCenter = $row.Cells["vCenter"].Value
                    }
                    if ($hasEnvironment) {
                        $vmData.Environment = $row.Cells["Environment"].Value
                    }

                    $script:CSVImportData += $vmData
                    $script:SelectedVMNames.Add($vmName) | Out-Null
                    $importedCount++
                }
            }

            $previewForm.Tag = @{
                Imported = $importedCount
                Skipped  = $skippedCount
            }
            $previewForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $previewForm.Close()
        })
        $previewForm.Controls.Add($btnImport)

        # ===== Cancel =====
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(630, 430)
        $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
        $btnCancel.Text = "Cancel"
        $btnCancel.FlatStyle = "Flat"
        $btnCancel.Add_Click({
            $previewForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $previewForm.Close()
        })
        $previewForm.Controls.Add($btnCancel)

        # ===== Show Dialog =====
        $result = $previewForm.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # עדכן CheckBoxes ברשימה הראשית
            $script:IsUpdatingList = $true
            foreach ($item in $lstVMs.Items) {
                $item.Checked = ($script:SelectedVMNames -contains $item.Text)
            }
            $script:IsUpdatingList = $false

            $lblSelectedCount.Text = "$($script:SelectedVMNames.Count) selected"

            # הודעה למשתמש
            $importInfo = $previewForm.Tag
            $message = "Imported $($importInfo.Imported) VMs."
            if ($importInfo.Skipped -gt 0) {
                $message += "`n`n$($importInfo.Skipped) VMs were not found and skipped."
            }

            [System.Windows.Forms.MessageBox]::Show(
                $message,
                "Import Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
    catch {
        Write-Log "Error loading CSV: $($_.Exception.Message)" -Level "ERROR"
    }
})
# ========== Clear Selection Button ==========
$btnClearSelection.Add_Click({
    $script:SelectedVMNames.Clear()
    foreach ($item in $lstVMs.Items) {
        $item.Checked = $false
    }
    $lblSelectedCount.Text = "0 selected"

})

# ========== SQL Integrated Auth Checkbox ==========
$chkSQLIntegrated.Add_CheckedChanged({
    $enabled = -not $chkSQLIntegrated.Checked
    $txtSQLUser.Enabled = $enabled
    $txtSQLPass.Enabled = $enabled
    $lblSQLUser.Enabled = $enabled
    $lblSQLPass.Enabled = $enabled
})

# ========== Check SQL Table Button ==========
$btnCheckTable.Add_Click({
    # Update settings from UI
    $script:DefaultSettings.SQLServer = $txtSQLServer.Text
    $script:DefaultSettings.SQLDatabase = $txtSQLDatabase.Text
    $script:DefaultSettings.SQLTable = $txtSQLTable.Text
    $script:DefaultSettings.SQLIntegratedAuth = $chkSQLIntegrated.Checked
    $script:DefaultSettings.SQLUsername = $txtSQLUser.Text
    $script:DefaultSettings.SQLPassword = $txtSQLPass.Text
    
    $lblArchiveStatus.Text = "Checking SQL connection..."
    $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    
    $connectionString = Get-SQLConnectionString
    
    if ([string]::IsNullOrWhiteSpace($connectionString)) {
        $lblArchiveStatus.Text = "Failed - Invalid SQL configuration"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    # Test connection
    $connTest = Test-SQLConnection -ConnectionString $connectionString
    
    if (-not $connTest.Success) {
        $lblArchiveStatus.Text = "Failed - $($connTest.Message)"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    # Check table exists
    $tableCheck = Test-SQLTableExists -ConnectionString $connectionString -TableName $($script:DefaultSettings.SQLTable)

    
    if (-not $tableCheck.Success) {
        $lblArchiveStatus.Text = "Failed - $($tableCheck.Message)"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    if ($tableCheck.Exists) {
        $lblArchiveStatus.Text = "OK - Table '$($script:DefaultSettings.SQLTable)' exists with $($tableCheck.Columns.Count) columns"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        $script:sqlTableExists = $true
        $btnRefreshArchive.Enabled = $true
        $btnExportCSV.Enabled = $true
        $btnExportExcel.Enabled = $true
        
        # Load archive data
        $btnRefreshArchive.PerformClick()
    } else {
        $createTable = [System.Windows.Forms.MessageBox]::Show(
            "Table '$($script:DefaultSettings.SQLTable)' does not exist.`n`nDo you want to create it?",
            "Create Table",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($createTable -eq [System.Windows.Forms.DialogResult]::Yes) {
            $createResult = New-SQLArchiveTable -ConnectionString $connectionString -TableName $($script:DefaultSettings.SQLTable)
            
            if ($createResult.Success) {
                $lblArchiveStatus.Text = "OK - $($createResult.Message)"
                $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
                $script:sqlTableExists = $true
                $btnRefreshArchive.Enabled = $true
                $btnExportCSV.Enabled = $true
                $btnExportExcel.Enabled = $true
            } else {
                $lblArchiveStatus.Text = "OK - $($createResult.Message)"
                $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
            }
        } else {
            $lblArchiveStatus.Text = "Table does not exist - create it before archiving"
            $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(255, 193, 7)
        }
    }
})

# ========== Refresh Archive Button ==========
$btnRefreshArchive.Add_Click({
    $connectionString = Get-SQLConnectionString
   
      # Each diagnostic query owns its resources; it is never used as an archive precondition.
    $testConn=$null;$testCmd=$null;$testReader=$null;$countCmd=$null
    # בדוק חיבור
    try {
        $testConn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $testConn.Open()
        $testCmd = $testConn.CreateCommand()
        $testCmd.CommandText = "SELECT @@SERVERNAME as Srv, DB_NAME() as DB"
        $testReader = $testCmd.ExecuteReader()
        if ($testReader.Read()) {
            Write-Log "📊 ARCHIVE TAB Connected to: Server='$($testReader['Srv'])', DB='$($testReader['DB'])'" -Level "INFO"
        }
        $testReader.Close()
        
        # ספור רשומות
        $countCmd = $testConn.CreateCommand()
        $countCmd.CommandText = "SELECT COUNT(*) FROM " + (ConvertTo-WizardSqlIdentifier -TableName $script:DefaultSettings.SQLTable)
        $count = $countCmd.ExecuteScalar()
        Write-Log "📊 Records in [$($script:DefaultSettings.SQLTable)]: $count" -Level "INFO"
        
        $testConn.Close()
    }
    catch {
        Write-Log "❌ Archive test connection failed: $($_.Exception.Message)" -Level "ERROR"
    }
    
    finally {if($testReader){$testReader.Dispose()};if($countCmd){$countCmd.Dispose()};if($testCmd){$testCmd.Dispose()};if($testConn){$testConn.Dispose()}}
    $archiveResult = Get-ArchiveData -ConnectionString $connectionString -TableName $script:DefaultSettings.SQLTable
    
    if ($archiveResult.Success) {
        $dgvArchive.DataSource = $archiveResult.Data
        $lblRecordCount.Text = "$($archiveResult.Data.Rows.Count) records"
        Write-Log "✅ Archive loaded: $($archiveResult.Data.Rows.Count) records" -Level "INFO"
        
        # Hide Id column if exists
        if ($dgvArchive.Columns["Id"]) {
            $dgvArchive.Columns["Id"].Visible = $false
        }
    } else {
        Write-Log "Failed to load archive: $($archiveResult.Message)" -Level "ERROR"
    }
})

# ========== Export CSV Button ==========
$btnExportCSV.Add_Click({
    if ($dgvArchive.DataSource -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.", "Export", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "VMDeletionArchive_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportResult = Export-ArchiveToFile -Data $dgvArchive.DataSource -FilePath $saveDialog.FileName -Format "CSV"
        if ($exportResult.Success) {
            [System.Windows.Forms.MessageBox]::Show($exportResult.Message, "Export", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            Write-Log "Failed to export archive to CSV: $($exportResult.Message)" -Level "ERROR"
        }
    }
})

# ========== Export Excel Button ==========
$btnExportExcel.Add_Click({
    if ($dgvArchive.DataSource -eq $null -or $dgvArchive.DataSource.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.", "Export", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "VMDeletionArchive_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportResult = Export-ArchiveToExcel -Data $dgvArchive.DataSource -FilePath $saveDialog.FileName
        if ($exportResult.Success) {
            [System.Windows.Forms.MessageBox]::Show($exportResult.Message, "Export", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            Write-Log "Failed to export archive to Excel: $($exportResult.Message)" -Level "ERROR"
        }
    }
})

# ========== Save SQL Settings Button ==========
$btnSaveSQLSettings.Add_Click({
    $script:DefaultSettings.SQLServer = $txtSQLServer.Text
    $script:DefaultSettings.SQLDatabase = $txtSQLDatabase.Text
    $script:DefaultSettings.SQLTable = $txtSQLTable.Text
    $script:DefaultSettings.SQLIntegratedAuth = $chkSQLIntegrated.Checked
    $script:DefaultSettings.SQLUsername = $txtSQLUser.Text
    $script:DefaultSettings.SQLPassword = $txtSQLPass.Text
    
    if (Save-Settings) {
        $lblArchiveStatus.Text = "OK - SQL settings saved"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    } else {
        $lblArchiveStatus.Text = "Failed - Failed to save settings"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

# ========== Browse Log Path Button ==========
$btnBrowseLog.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select Log Folder"
    $folderDialog.SelectedPath = $txtLogPath.Text
    
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLogPath.Text = $folderDialog.SelectedPath
    }
})

# ========== Save Settings Button ==========
$btnSaveSettings.Add_Click({
    # Update all settings from UI
    $script:DefaultSettings.SCCMServer = $txtSettingsSCCM.Text
    $script:DefaultSettings.SCCMSiteCode = $txtSettingsSiteCode.Text
    $script:DefaultSettings.DNSServer = $txtSettingsDNS.Text
    $script:DefaultSettings.DNSZone = $txtSettingsDNSZone.Text
    $script:DefaultSettings.DomainController = $txtSettingsDC.Text
    $script:DefaultSettings.EmailNotification = $txtSettingsEmail.Text
    $script:DefaultSettings.EnableLogging = $chkEnableLogging.Checked
    $script:DefaultSettings.ConfirmDeletion = $chkConfirmDeletion.Checked
    $script:DefaultSettings.DeleteFromDisk = $chkDeleteFromDisk.Checked
    $script:DefaultSettings.PowerOffTimeout = [int]$numTimeout.Value
    $script:DefaultSettings.LogPath = $txtLogPath.Text
    
    # Email settings
    $script:DefaultSettings.EnableEmailNotification = $chkEnableEmail.Checked
    $script:DefaultSettings.SMTPServer = $txtSMTPServer.Text
    $script:DefaultSettings.SMTPPort = [int]$numSMTPPort.Value
    $script:DefaultSettings.SMTPUseSSL = $chkSMTPSSL.Checked
    $script:DefaultSettings.SMTPFrom = $txtSMTPFrom.Text
    
    $script:DefaultSettings.SQLServer=$txtSQLServer.Text
    $script:DefaultSettings.SQLDatabase=$txtSQLDatabase.Text
    $script:DefaultSettings.SQLTable=$txtSQLTable.Text
    $script:DefaultSettings.SQLIntegratedAuth=$chkSQLIntegrated.Checked
    $script:DefaultSettings.SQLUsername=$txtSQLUser.Text
    $script:DefaultSettings.SQLPassword=$txtSQLPass.Text
    Save-DeletionStepPreferences
    # Save email groups
    Save-EmailGroups
    
    if (Save-Settings) {
        $lblSettingsStatus.Text = "OK - Settings saved successfully at $(Get-Date -Format 'HH:mm:ss')"
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        Write-Log "Settings saved"
    } else {
        $lblSettingsStatus.Text = "Failed - Failed to save settings"
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::Red
    }
    
})

# ========== Reset Settings Button ==========
$btnResetSettings.Add_Click({
    if (Reset-Configuration) {
        Refresh-SettingsForm
        [System.Windows.Forms.MessageBox]::Show(
            "Settings have been reset successfully!",
            "Settings Reset",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# ========== Delete VM Button ==========
$btnDelete.Add_Click({
    # בדיקה לפי SelectedVMNames
    if ($script:SelectedVMNames.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one VM to delete.", "No Selection", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # בניית רשימת VMs מ-SelectedVMNames
    $selectedVMs = @()
    $selectedVCIndex = $cboVCServer.SelectedIndex
    $defaultEnvironment = if ($selectedVCIndex -ge 0) {
        $script:vCenterServers[$selectedVCIndex].Name
    }
    else {
        $script:RunEnvironment
    }

    foreach ($vmName in $script:SelectedVMNames) {
        $vmData = $script:allVMsData | Where-Object { $_.Name -eq $vmName } | Select-Object -First 1
        $csvVMData = $script:CSVImportData | Where-Object { $_.VMName -eq $vmName } | Select-Object -First 1
        $osType = if ($vmData) { $vmData.OSType } else { "Unknown" }
        $vmEnvironment = if ($csvVMData -and $csvVMData.Environment) {
            $csvVMData.Environment
        }
        elseif ($csvVMData -and $csvVMData.vCenter) {
            Resolve-VCenterEnvironment -Value $csvVMData.vCenter
        }
        else {
            $defaultEnvironment
        }
        
        $selectedVMs += @{
            Name        = $vmName
            OSType      = $osType
            Environment = $vmEnvironment
        }
    }
    
    $vmNames = ($selectedVMs | ForEach-Object { $_.Name }) -join ", "
    $reason = $txtReason.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($reason)) {
        $reason = "No reason provided"
    }

# Confirmation
if($script:DefaultSettings.ConfirmDeletion){
    if(-not (Show-WizardDeletionConfirmation -SelectedVMs $selectedVMs -Reason $reason)){return}
}
    # Disable controls during deletion
    Set-ControlsEnabled -Enabled $false
    
    $totalVMs = $selectedVMs.Count
    $currentVM = 0
    
    # === רשימות לדוח סיכום ===
    $successList = @()
    $failedList = @()
    $skippedList = @()
    
    foreach ($vmInfo in $selectedVMs) {
        $vmName = $vmInfo.Name
        $osType = $vmInfo.OSType
        $vmEnvironment = $vmInfo.Environment
        $currentVM++
        $lblVMProgress.Text = "VM Progress: $vmName ($currentVM/$totalVMs)"
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        Write-Log "Starting deletion process for VM: $vmName ($currentVM of $totalVMs)"
        $script:CurrentVmDeletionReportRecord = $null
        Reset-AllStatus

        # Track deletion results for this VM
        $vmResult = @{
            VMName = $vmName
            OSType = $osType
            Environment = $vmEnvironment
            VCenter = (Get-VmDeletionReportVCenter)
            ReportSteps = (New-VmDeletionReportStepStates)
            Success = $true
            Errors = @()
            SCCMDeleted = $false
            ADDeleted = $false
            DNSDeleted = $false
            VMDeleted = $false
            Archived = $false
            EmailSent = $false
        }
        
        $script:CurrentVmDeletionReportRecord = $vmResult
        # Task 0: vCenter Connection (already connected)
        Update-TaskStatus -TaskIndex 0 -Status "Success" -Message "Connected to $($global:DefaultVIServer.Name)"
        $progressBar.Value = 14
        
        # Get VM details before deletion
        try {
            $vmDetails = Get-VMDetailsForArchive -VMName $vmName
        }
        catch {
            $vmResult.Errors += "Failed to get VM details: $($_.Exception.Message)"
            $vmDetails = @{ ServerName = $vmName; VMName = $vmName }
        }
        
        $osType=Get-DeletionOSType -Preferred $osType -OperatingSystem $vmDetails.OperatingSystem -GuestId $vmDetails.GuestId
        $vmResult.OSType=$osType
        $vmResult.IsTemplate=[bool]$vmDetails.IsTemplate
        $vmResult.ClusterName=$vmDetails.ClusterName
        # Task 1: Delete SCCM Record
        $stepSkipReason=Get-DeletionStepSkipReason -Step SCCM -OSType $osType -Selected $chkSCCM.Checked
        if (-not $stepSkipReason) {
            Update-TaskStatus -TaskIndex 1 -Status "Running" -Message ""
            try {
                $sccmResult = Remove-SCCMDevice -ComputerName $vmName -SCCMServer $script:DefaultSettings.SCCMServer -SiteCode $script:DefaultSettings.SCCMSiteCode
                
                if ($sccmResult.Success) {
                    $vmResult.SCCMDeleted = $sccmResult.Found
                    Update-TaskStatus -TaskIndex 1 -Status "Success" -Message $sccmResult.Message
                } else {
                    Update-TaskStatus -TaskIndex 1 -Status "Failed" -Message $sccmResult.Message
                    $vmResult.Errors += "SCCM: $($sccmResult.Message)"
                }
            }
            catch {
                Update-TaskStatus -TaskIndex 1 -Status "Failed" -Message $_.Exception.Message
                $vmResult.Errors += "SCCM: $($_.Exception.Message)"
            }
        } else {
            Update-TaskStatus -TaskIndex 1 -Status "Skipped" -Message $stepSkipReason
            Write-Log "$vmName - SCCM: $stepSkipReason"
        }
        $progressBar.Value = 28
        
        # Task 2: Delete AD Record
        $stepSkipReason=Get-DeletionStepSkipReason -Step AD -OSType $osType -Selected $chkAD.Checked
        if (-not $stepSkipReason) {
            Update-TaskStatus -TaskIndex 2 -Status "Running" -Message ""
            try {
                $adResult = Remove-ADDevice -ComputerName $vmName `
                    -DomainController $script:DefaultSettings.DomainController `
                    -SearchBase (Get-ADSearchBaseForEnvironment -EnvironmentName $vmEnvironment)
        
                if ($adResult.Success) {
                    $vmResult.ADDeleted = $adResult.Found
                    Update-TaskStatus -TaskIndex 2 -Status "Success" -Message $adResult.Message
                } else {
                    Update-TaskStatus -TaskIndex 2 -Status "Failed" -Message $adResult.Message
                    $vmResult.Errors += "AD: $($adResult.Message)"
                }
            }
            catch {
                Update-TaskStatus -TaskIndex 2 -Status "Failed" -Message $_.Exception.Message
                $vmResult.Errors += "AD: $($_.Exception.Message)"
            }
        } else {
            Update-TaskStatus -TaskIndex 2 -Status "Skipped" -Message $stepSkipReason
            Write-Log "$vmName - AD: $stepSkipReason"
        }
        $progressBar.Value = 42
        
        # Task 3: Delete DNS Records
        $stepSkipReason=Get-DeletionStepSkipReason -Step DNS -OSType $osType -Selected $chkDNS.Checked
        if (-not $stepSkipReason) {
            Update-TaskStatus -TaskIndex 3 -Status "Running" -Message ""
            try {
                $dnsResult = Remove-DNSRecords -ComputerName $vmName -DNSServer $script:DefaultSettings.DNSServer -DNSZone $script:DefaultSettings.DNSZone
                
                if ($dnsResult.Success) {
                    $vmResult.DNSDeleted = $dnsResult.Found
                    Update-TaskStatus -TaskIndex 3 -Status "Success" -Message $dnsResult.Message
                } else {
                    Update-TaskStatus -TaskIndex 3 -Status "Failed" -Message $dnsResult.Message
                    $vmResult.Errors += "DNS: $($dnsResult.Message)"
                }
            }
            catch {
                Update-TaskStatus -TaskIndex 3 -Status "Failed" -Message $_.Exception.Message
                $vmResult.Errors += "DNS: $($_.Exception.Message)"
            }
        } else {
            Update-TaskStatus -TaskIndex 3 -Status "Skipped" -Message $stepSkipReason
            Write-Log "$vmName - DNS: $stepSkipReason"
        }
        $progressBar.Value = 56
        
# Task 4: Delete VM or Template
$stepSkipReason=Get-DeletionStepSkipReason -Step VM -OSType $osType -Selected $chkVM.Checked
if (-not $stepSkipReason) {

    Update-TaskStatus -TaskIndex 4 -Status "Running" -Message ""

    try {
        # בדוק אם זה Template
        $vmData = $script:allVMsData | Where-Object { $_.Name -eq $vmName } | Select-Object -First 1
        $isTemplate = if ($vmData) { $vmData.IsTemplate } else { $false }
        
        if ($isTemplate) {
            # מחיקת Template
            Update-TaskStatus -TaskIndex 4 -Status "Running" -Message "Removing template..."
            
            $template = Get-Template -Name $vmName -ErrorAction Stop
            
            if ($script:DefaultSettings.DeleteFromDisk) {
                Remove-Template -Template $template -DeletePermanently -Confirm:$false -ErrorAction Stop
            }
            else {
                Remove-Template -Template $template -Confirm:$false -ErrorAction Stop
            }
            
            $vmResult.VMDeleted = $true
            Update-TaskStatus -TaskIndex 4 -Status "Success" -Message "Template deleted successfully"
        }
        else {
            # מחיקת VM רגיל
            $vm = Get-VM -Name $vmName -ErrorAction Stop

            # Power off if running
            if ($vm.PowerState -eq "PoweredOn") {

                Update-TaskStatus -TaskIndex 4 -Status "Running" -Message "Powering off VM..."

                try {
                    Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null

                    $timeout = $script:DefaultSettings.PowerOffTimeout
                    $elapsed = 0

                    while ((Get-VM -Name $vmName -ErrorAction SilentlyContinue).PowerState -eq "PoweredOn" -and $elapsed -lt $timeout) {
                        Start-Sleep -Seconds 2
                        $elapsed += 2
                    }
                }
                catch {
                    $vmResult.Errors += "Power off failed: $($_.Exception.Message)"
                }

            }

            # Delete VM
            if ($script:DefaultSettings.DeleteFromDisk) {
                Remove-VM -VM $vm -DeletePermanently -Confirm:$false -ErrorAction Stop
            }
            else {
                Remove-VM -VM $vm -Confirm:$false -ErrorAction Stop
            }

            $vmResult.VMDeleted = $true
            Update-TaskStatus -TaskIndex 4 -Status "Success" -Message "VM deleted successfully"
        }
    }
    catch {

        $vmResult.VMDeleted = $false
        $vmResult.Success = $false
        $vmResult.Errors += "VM/Template Delete: $($_.Exception.Message)"

        Update-TaskStatus -TaskIndex 4 -Status "Failed" -Message $_.Exception.Message
        Write-Log "VM/Template deletion failed for $vmName : $($_.Exception.Message)" -Level "ERROR"

    }

}
        else {
            Update-TaskStatus -TaskIndex 4 -Status 'Skipped' -Message $stepSkipReason
            Write-Log "$vmName - VM: $stepSkipReason"
        }
        $progressBar.Value = 70

        $vmResult.Success=($vmResult.Errors.Count -eq 0)
        # Task 5: Archive to SQL
        $stepSkipReason=Get-DeletionStepSkipReason -Step Archive -OSType $osType -Selected $chkArchive.Checked
        if (-not $stepSkipReason) {
            Update-TaskStatus -TaskIndex 5 -Status "Running" -Message ""
            try {
                $status = if ($vmResult.Success) { "Success" } else { "Failed" }

                $archiveResult = Save-VMDeletionToArchive -VMDetails $vmDetails -Reason $reason `
                    -SCCMDeleted $vmResult.SCCMDeleted -ADDeleted $vmResult.ADDeleted `
                    -DNSDeleted $vmResult.DNSDeleted -VMDeleted $vmResult.VMDeleted -Status $status

                if ($archiveResult.Success) {
                    $vmResult.Archived = $true
                    Update-TaskStatus -TaskIndex 5 -Status "Success" -Message $archiveResult.Message
                }
                else {
                    $vmResult.Errors += "Archive: $($archiveResult.Message)"
                    Update-TaskStatus -TaskIndex 5 -Status "Failed" -Message $archiveResult.Message
                }
            }
            catch {
                $vmResult.Errors += "Archive: $($_.Exception.Message)"
                Update-TaskStatus -TaskIndex 5 -Status "Failed" -Message $_.Exception.Message
                Write-Log "Archive failed for $vmName : $($_.Exception.Message)" -Level "ERROR"
            }
        }
        else {
            Update-TaskStatus -TaskIndex 5 -Status "Skipped" -Message $stepSkipReason
            Write-Log "$vmName - Archive: $stepSkipReason"
        }
        $progressBar.Value = 85

        # Task 6: Send Email Notification
        $stepSkipReason=Get-DeletionStepSkipReason -Step Email -OSType $osType -Selected $chkEmail.Checked
        if (-not $stepSkipReason) {
            Update-TaskStatus -TaskIndex 6 -Status "Running" -Message ""
            try {
                $emailResult = Send-VMDeletionNotification -VMDetails $vmDetails -Reason $reason -OSType $osType -DeletionSuccess $vmResult.VMDeleted -ReportRecord $vmResult
                
                if ($emailResult.Success) {
                    $vmResult.EmailSent = $true
                    Update-TaskStatus -TaskIndex 6 -Status "Success" -Message $emailResult.Message
                } else {
                    $vmResult.Errors += "Email: $($emailResult.Message)"
                    Update-TaskStatus -TaskIndex 6 -Status "Failed" -Message $emailResult.Message
                }
            }
            catch {
                $vmResult.Errors += "Email: $($_.Exception.Message)"
                Update-TaskStatus -TaskIndex 6 -Status "Failed" -Message $_.Exception.Message
            }
        } else {
            Update-TaskStatus -TaskIndex 6 -Status "Skipped" -Message $stepSkipReason
        }
        $progressBar.Value = 100
        
        # הוסף לרשימה המתאימה
        $vmResult.Success=($vmResult.Errors.Count -eq 0)
        if($vmResult.Errors.Count -gt 0){$failedList+=$vmResult}
        elseif($vmResult.VMDeleted -or -not $chkVM.Checked){$successList+=$vmResult}
        else {
            $skippedList += $vmResult
        }
        
        $script:CurrentVmDeletionReportRecord = $null
        Write-Log "Deletion process completed for VM: $vmName (Success: $($vmResult.Success))"
        
        # המתן קצת בין מכונות
        Start-Sleep -Milliseconds 500
        }
    $lblVMProgress.Text = "VM Progress: Completed"
    $lblVMProgress.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)

    # === שלח דוח סיכום במייל ===
    if ($chkEmail.Checked -and $script:DefaultSettings.EnableEmailNotification -and ($failedList.Count -gt 0 -or $successList.Count -gt 0)) {
        Send-DeletionSummaryReport -SuccessList $successList -FailedList $failedList -SkippedList $skippedList -Reason $reason
    }

    # Re-enable controls
    Set-ControlsEnabled -Enabled $true
    
    # Refresh VM list
    Get-VMList
   
    # איפוס Progress Bar וסטטוסים
    $progressBar.Value = 0
    Reset-AllStatus
    
    # נקה את הבחירות
    $script:SelectedVMNames.Clear()
    $lblSelectedCount.Text = "0 selected"
    
    # בטל סימון כל הפריטים ברשימה
    $script:IsUpdatingList = $true
    foreach ($item in $lstVMs.Items) {
        $item.Checked = $false
    }
    $script:IsUpdatingList = $false
    
    # הודעת סיכום - TopMost ראשון במסך
    $summaryMsg = "Deletion process completed!`n`n"
    $summaryMsg += "Successful: $($successList.Count)`n"
    $summaryMsg += "Failed: $($failedList.Count)`n"
    $summaryMsg += "Skipped: $($skippedList.Count)`n"

    if ($failedList.Count -gt 0) {
        $summaryMsg += "`nFailed VMs:`n"
        foreach ($failed in $failedList) {
            $summaryMsg += "  - $($failed.VMName): $($failed.Errors -join '; ')`n"
        }
    }
    
    $icon = if ($failedList.Count -eq 0) { 
        [System.Windows.Forms.MessageBoxIcon]::Information 
    } else { 
        [System.Windows.Forms.MessageBoxIcon]::Warning 
    }
    
    # הצג הודעה TopMost בראש המסך
    $topForm = New-Object System.Windows.Forms.Form
    $topForm.TopMost = $true
    $topForm.StartPosition = "CenterScreen"
    $topForm.ShowInTaskbar = $false
    $topForm.Size = New-Object System.Drawing.Size(1, 1)
    $topForm.Opacity = 0
    $topForm.Show()
    
    [System.Windows.Forms.MessageBox]::Show(
        $topForm,
        $summaryMsg,
        "Deletion Summary",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon
    )
    
    $topForm.Close()
    $topForm.Dispose()
    
    # החזר פוקוס לחלון הראשי
    $form.Activate()
})

# ========== Resume Button ==========
$btnResume.Add_Click({
    $progressFile = "$env:TEMP\VMDeletionTool\progress.json"
    $availableProgressFile = Wait-JsonFileAvailable -Path $progressFile
    if ($availableProgressFile) {
        $progress = Get-Content -LiteralPath $availableProgressFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Resume from: $($progress.LastVM)?`nLast run: $($progress.Timestamp)",
            "Resume Deletion",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            # מצא את ה-VM ברשימה וסמן אותו
            $script:SelectedVMNames.Clear()
            $script:SelectedVMNames.Add($progress.LastVM) | Out-Null
            
            # סמן ברשימה
            foreach ($item in $lstVMs.Items) {
                if ($item.Text -eq $progress.LastVM) {
                    $item.Selected = $true
                    $lstVMs.EnsureVisible($item.Index)
                    break
                }
            }
            
            $lblSelectedCount.Text = "1 selected"
            
            [System.Windows.Forms.MessageBox]::Show(
                "VM '$($progress.LastVM)' selected.`nClick 'Delete VM' to continue.",
                "Ready to Resume",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "No progress file found. Nothing to resume.",
            "Resume",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# ========== Close Button ==========
$btnClose.Add_Click({
    if ($script:vcConnected) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Disconnect from vCenter and close?",
            "Confirm Close",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue
            $form.Close()
        }
    } else {
        $form.Close()
    }
})

# ========== Form Closing Event ==========
$form.Add_FormClosing({})
#endregion Event Handlers

# SQL is checked manually from the Archive tab to keep GUI startup fast.
$form.Add_Shown({
    $lblArchiveStatus.Text = "SQL not checked yet. Click 'Check Table' when needed."
    $lblArchiveStatus.ForeColor = $uiMuted
})


#region WPF single-screen wizard
# Native WPF uses the reference's rounded templates. Existing WinForms controls
# remain the internal workflow contract; infrastructure handlers are unchanged.
function Initialize-DeletionWizard {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, WindowsFormsIntegration
    $script:WizardSyncing = $false
    $script:WizardBusy = $false
    $script:WizardListKey = ''
    $script:WizardSelectionKey = ''
    $script:WizardAdvanced = $null
    $wizardXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="VM Deletion Tool" Width="1440" Height="940" MinWidth="1180" MinHeight="860" WindowStartupLocation="CenterScreen" Background="#0C1730" FontFamily="Segoe UI" FontSize="13" Foreground="#E8F0FF" UseLayoutRounding="True" TextOptions.TextFormattingMode="Display"><Window.Resources>
    <SolidColorBrush x:Key="Bg" Color="#0C1730"/>
    <SolidColorBrush x:Key="Panel" Color="#12234A"/>
    <SolidColorBrush x:Key="Raised" Color="#1A3060"/>
    <SolidColorBrush x:Key="Line" Color="#2B4577"/>
    <SolidColorBrush x:Key="Text" Color="#E8F0FF"/>
    <SolidColorBrush x:Key="Muted" Color="#A7B9D9"/>
    <SolidColorBrush x:Key="Accent" Color="#62E4F4"/>
    <SolidColorBrush x:Key="AccentDim" Color="#14425F"/>
    <SolidColorBrush x:Key="Danger" Color="#F0647A"/>

    <Style x:Key="NoRepeat" TargetType="RepeatButton">
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="IsTabStop" Value="False"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="RepeatButton"><Border Background="Transparent"/></ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="8"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="ScrollBar">
          <Track x:Name="PART_Track" IsDirectionReversed="True">
            <Track.DecreaseRepeatButton><RepeatButton Command="ScrollBar.PageUpCommand" Style="{StaticResource NoRepeat}"/></Track.DecreaseRepeatButton>
            <Track.Thumb><Thumb><Thumb.Template><ControlTemplate TargetType="Thumb">
              <Border CornerRadius="4" Background="#2F4A7D" Margin="1,0"/>
            </ControlTemplate></Thumb.Template></Thumb></Track.Thumb>
            <Track.IncreaseRepeatButton><RepeatButton Command="ScrollBar.PageDownCommand" Style="{StaticResource NoRepeat}"/></Track.IncreaseRepeatButton>
          </Track>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <!-- Left navigation item: filled cyan when selected -->
    <Style x:Key="Nav" TargetType="RadioButton">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Height" Value="44"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="bd" CornerRadius="8" BorderThickness="1" BorderBrush="{StaticResource Line}" Background="Transparent" Padding="16,0">
            <ContentPresenter VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="bd" Property="Background" Value="{StaticResource Accent}"/>
              <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
              <Setter Property="Foreground" Value="#0C1730"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <!-- Small outlined cyan button (like "Add device") -->
    <Style x:Key="Pill" TargetType="Button">
      <Setter Property="Foreground" Value="#0C1730"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="bd" CornerRadius="6" Background="{StaticResource Accent}" Padding="14,6">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.85"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <Style x:Key="Ghost" TargetType="Button">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Margin" Value="4,0,0,0"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="bd" CornerRadius="6" BorderThickness="1" BorderBrush="{StaticResource Line}" Padding="12,5" Background="Transparent">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
              <Setter Property="Foreground" Value="{StaticResource Accent}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <Style x:Key="DangerBtn" TargetType="Button">
      <Setter Property="Foreground" Value="#1B0A12"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="bd" CornerRadius="10" Background="{StaticResource Danger}" Height="46">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.88"/></Trigger>
            <Trigger Property="IsEnabled" Value="False">
              <Setter TargetName="bd" Property="Background" Value="#2A3F6B"/>
              <Setter Property="Foreground" Value="#7185AB"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <!-- On/off switch -->
    <Style x:Key="Switch" TargetType="CheckBox">
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="CheckBox">
          <Border x:Name="track" Width="42" Height="22" CornerRadius="11" Background="#2A3F6B">
            <Ellipse x:Name="thumb" Width="16" Height="16" Fill="#C9D6F0" HorizontalAlignment="Left" Margin="3,0,0,0"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="track" Property="Background" Value="{StaticResource Accent}"/>
              <Setter TargetName="thumb" Property="HorizontalAlignment" Value="Right"/>
              <Setter TargetName="thumb" Property="Margin" Value="0,0,3,0"/>
              <Setter TargetName="thumb" Property="Fill" Value="#0C1730"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <Style x:Key="Tick" TargetType="CheckBox">
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="CheckBox">
          <Border x:Name="bd" Width="18" Height="18" CornerRadius="5" BorderThickness="1.5" BorderBrush="#4A6699" Background="Transparent">
            <Path x:Name="tick" Data="M3,8.5 L7,12.5 L13,4.5" Stroke="#0C1730" StrokeThickness="2" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="bd" Property="Background" Value="{StaticResource Accent}"/>
              <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
              <Setter TargetName="tick" Property="Visibility" Value="Visible"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <!-- Step tile: outlined cyan when included -->
    <Style x:Key="Tile" TargetType="ToggleButton">
      <Setter Property="Margin" Value="8"/>
      <Setter Property="Height" Value="104"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="ToggleButton">
          <Border x:Name="bd" CornerRadius="14" BorderThickness="1.5" BorderBrush="{StaticResource Line}" Background="{StaticResource Panel}" Padding="14,12">
            <Grid>
              <Ellipse x:Name="dot" Width="9" Height="9" Fill="{StaticResource Line}" HorizontalAlignment="Right" VerticalAlignment="Top"/>
              <ContentPresenter VerticalAlignment="Bottom"/>
            </Grid>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#4A6699"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
              <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentDim}"/>
              <Setter TargetName="dot" Property="Fill" Value="{StaticResource Accent}"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <Style x:Key="Field" TargetType="TextBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="CaretBrush" Value="{StaticResource Accent}"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="TextBox">
          <Border x:Name="bd" Background="{StaticResource Bg}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="8">
            <Grid>
              <ScrollViewer x:Name="PART_ContentHost" Margin="12,0" VerticalAlignment="Center"/>
              <TextBlock x:Name="hint" Text="Search VMs" Foreground="{StaticResource Muted}" Margin="14,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False" Visibility="Collapsed"/>
            </Grid>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="Text" Value=""><Setter TargetName="hint" Property="Visibility" Value="Visible"/></Trigger>
            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <Style x:Key="GridHeader" TargetType="DataGridColumnHeader">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Padding" Value="12,0,0,10"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="DataGridColumnHeader">
          <Border BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1" Padding="{TemplateBinding Padding}" Height="34">
            <ContentPresenter VerticalAlignment="Bottom"/>
          </Border>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <Style x:Key="GridCell" TargetType="DataGridCell">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="DataGridCell">
          <Border Background="Transparent" Padding="12,0"><ContentPresenter VerticalAlignment="Center"/></Border>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>

    <Style x:Key="GridRow" TargetType="DataGridRow">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#16295A"/></Trigger>
        <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#16295A"/></Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="Bar" TargetType="ProgressBar">
      <Setter Property="Height" Value="8"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="ProgressBar">
          <Grid>
            <Border CornerRadius="4" Background="#1F3563"/>
            <Border x:Name="PART_Track"/>
            <Border x:Name="PART_Indicator" CornerRadius="4" Background="{StaticResource Accent}" HorizontalAlignment="Left"/>
          </Grid>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>
  
    <Style TargetType="Control"><Setter Property="FocusVisualStyle"><Setter.Value><Style TargetType="Control"><Setter Property="Template"><Setter.Value><ControlTemplate><Rectangle Margin="2" RadiusX="5" RadiusY="5" Stroke="#62E4F4" StrokeDashArray="2 2" StrokeThickness="1"/></ControlTemplate></Setter.Value></Setter></Style></Setter.Value></Setter></Style>
    <Style x:Key="Editor" TargetType="TextBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/><Setter Property="CaretBrush" Value="{StaticResource Accent}"/>
      <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TextBox">
        <Border x:Name="edge" Background="{StaticResource Bg}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="8" Padding="10,7"><ScrollViewer x:Name="PART_ContentHost"/></Border>
        <ControlTemplate.Triggers><Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="edge" Property="BorderBrush" Value="{StaticResource Accent}"/></Trigger><Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.55"/></Trigger></ControlTemplate.Triggers>
      </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Picker" TargetType="ComboBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/><Setter Property="Background" Value="{StaticResource Bg}"/><Setter Property="MinHeight" Value="36"/>
      <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBox"><Grid>
        <ToggleButton Focusable="False" IsChecked="{Binding IsDropDownOpen,RelativeSource={RelativeSource TemplatedParent},Mode=TwoWay}"><ToggleButton.Template><ControlTemplate TargetType="ToggleButton"><Border Background="#0C1730" BorderBrush="#2B4577" BorderThickness="1" CornerRadius="8"><TextBlock Text="&#x2304;" Foreground="#A7B9D9" HorizontalAlignment="Right" Margin="0,0,12,0" VerticalAlignment="Center"/></Border></ControlTemplate></ToggleButton.Template></ToggleButton>
        <ContentPresenter Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" Margin="12,0,32,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
        <Popup x:Name="PART_Popup" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Placement="Bottom" Focusable="False"><Border Background="#12234A" BorderBrush="#2B4577" BorderThickness="1" CornerRadius="8" MinWidth="{Binding ActualWidth,RelativeSource={RelativeSource TemplatedParent}}" Padding="4"><ScrollViewer MaxHeight="260"><ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/></ScrollViewer></Border></Popup>
      </Grid><ControlTemplate.Triggers><Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.55"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="ComboBoxItem"><Setter Property="Foreground" Value="#E8F0FF"/><Setter Property="Padding" Value="10,8"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBoxItem"><Border x:Name="row" Background="Transparent" CornerRadius="5" Padding="{TemplateBinding Padding}"><ContentPresenter/></Border><ControlTemplate.Triggers><Trigger Property="IsHighlighted" Value="True"><Setter TargetName="row" Property="Background" Value="#14425F"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
    <Style TargetType="Button" BasedOn="{StaticResource {x:Type Control}}"><Style.Triggers><Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger></Style.Triggers></Style>
    <Style x:Key="ActionTile" TargetType="ToggleButton" BasedOn="{StaticResource Tile}"><Setter Property="Height" Value="88"/><Setter Property="Margin" Value="4"/><Style.Triggers><Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger></Style.Triggers></Style>
</Window.Resources>
  <Grid Margin="28,20,28,24">
    <Grid.RowDefinitions><RowDefinition Height="62"/><RowDefinition Height="50"/><RowDefinition Height="*"/><RowDefinition Height="216"/></Grid.RowDefinitions>
    <Grid Grid.Row="0"><Grid.ColumnDefinitions><ColumnDefinition Width="56"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <Border Width="42" Height="42" CornerRadius="12" Background="{StaticResource Panel}" BorderBrush="{StaticResource Line}" BorderThickness="1" HorizontalAlignment="Left" VerticalAlignment="Top"><TextBlock Text="VD" FontSize="16" FontWeight="Bold" Foreground="{StaticResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
      <StackPanel Grid.Column="1"><StackPanel Orientation="Horizontal"><TextBlock Text="VM cleanup" FontFamily="Segoe UI Light" FontSize="28"/><TextBlock Text="v1.0" Foreground="{StaticResource Muted}" Margin="16,14,0,0" FontSize="11"/></StackPanel><TextBlock Text="Connect, select and review. The entire workflow stays in one place." Foreground="{StaticResource Muted}" Margin="0,4,0,0"/></StackPanel>
      <Button x:Name="wAdvanced" Grid.Column="2" Style="{StaticResource Ghost}" VerticalAlignment="Top" Height="40" ToolTip="Archive and settings" AutomationProperties.Name="Advanced settings"><StackPanel Orientation="Horizontal"><TextBlock Text="&#xE713;" FontFamily="Segoe MDL2 Assets" FontSize="17" Margin="0,0,10,0"/><TextBlock Text="Advanced settings"/></StackPanel></Button>
    </Grid>
    <Border Grid.Row="1" CornerRadius="10" Background="{StaticResource Panel}" Margin="0,4,0,12" Padding="16,0"><Grid><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
      <TextBlock x:Name="wStage1" Text="01   Connect environment" Foreground="{StaticResource Accent}" VerticalAlignment="Center"/>
      <TextBlock x:Name="wStage2" Grid.Column="1" Text="02   Select machines" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
      <TextBlock x:Name="wStage3" Grid.Column="2" Text="03   Review cleanup" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
      <TextBlock x:Name="wStage4" Grid.Column="3" Text="04   Execute and monitor" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
    </Grid></Border>
    <Grid Grid.Row="2" Margin="0,0,0,16"><Grid.ColumnDefinitions><ColumnDefinition Width="1.35*" MinWidth="560"/><ColumnDefinition Width="20"/><ColumnDefinition Width="*" MinWidth="430"/></Grid.ColumnDefinitions>
      <Grid><Grid.RowDefinitions><RowDefinition Height="102"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        <Border Background="{StaticResource Panel}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="14" Padding="18,12" Margin="0,0,0,14"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="112"/></Grid.ColumnDefinitions>
          <TextBlock Text="ENVIRONMENT" FontSize="11" Foreground="{StaticResource Muted}"/>
          <ComboBox x:Name="wEnvironment" Grid.Row="1" Style="{StaticResource Picker}" Margin="0,6,12,0" AutomationProperties.Name="vCenter environment"/>
          <Button x:Name="wConnect" Grid.Row="1" Grid.Column="1" Content="Connect" Style="{StaticResource Pill}" Height="36" VerticalAlignment="Bottom"/>
        </Grid></Border>
        <Border Grid.Row="1" Background="{StaticResource Panel}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="14" Padding="18,14"><Grid><Grid.RowDefinitions><RowDefinition Height="36"/><RowDefinition Height="48"/><RowDefinition Height="*"/><RowDefinition Height="25"/></Grid.RowDefinitions>
          <Grid><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Text="Select machines" FontSize="18" FontWeight="SemiBold" VerticalAlignment="Center"/><Button x:Name="wRefresh" Grid.Column="1" Style="{StaticResource Ghost}" Content="Refresh" Height="30"/><Button x:Name="wImport" Grid.Column="2" Style="{StaticResource Pill}" Content="+ Import CSV" Height="30" Margin="8,0,0,0"/></Grid>
          <Grid Grid.Row="1" Margin="0,8,0,6"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="112"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBox x:Name="wSearch" Style="{StaticResource Field}" AutomationProperties.Name="Search VM names"/><ComboBox x:Name="wOS" Grid.Column="1" Style="{StaticResource Picker}" Margin="10,0,0,0" AutomationProperties.Name="Operating system filter"/><Button x:Name="wClear" Grid.Column="2" Style="{StaticResource Ghost}" Content="Clear" Margin="8,0,0,0"/></Grid>
          <DataGrid x:Name="wMachines" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" CanUserResizeRows="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="None" SelectionMode="Single" Background="Transparent" BorderThickness="0" RowBackground="Transparent" RowHeight="36" ColumnHeaderStyle="{StaticResource GridHeader}" CellStyle="{StaticResource GridCell}" RowStyle="{StaticResource GridRow}" EnableRowVirtualization="True" AutomationProperties.Name="Machines to delete">
            <DataGrid.Columns><DataGridTemplateColumn Width="40"><DataGridTemplateColumn.CellTemplate><DataTemplate><CheckBox Style="{StaticResource Tick}" IsChecked="{Binding Selected,Mode=TwoWay,UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" AutomationProperties.Name="{Binding Name}"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
              <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="1.3*"/><DataGridTextColumn Header="System" Binding="{Binding OS}" Width="*"/><DataGridTextColumn Header="Power" Binding="{Binding Power}" Width="100"/>
            </DataGrid.Columns>
          </DataGrid>
          <Grid Grid.Row="3" Margin="0,7,0,0"><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock x:Name="wLoaded" Text="Connect to load machines" Foreground="{StaticResource Muted}" FontSize="11"/><StackPanel Grid.Column="1" Orientation="Horizontal"><CheckBox x:Name="wTemplates" Style="{StaticResource Tick}" VerticalAlignment="Center" AutomationProperties.Name="Include VM templates"/><TextBlock Text="Include templates" Foreground="{StaticResource Muted}" FontSize="11" Margin="6,0,0,0"/></StackPanel></Grid>
        </Grid></Border>
      </Grid>
      <Border Grid.Column="2" Background="{StaticResource Panel}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="14" Padding="16,14"><Grid><Grid.RowDefinitions><RowDefinition Height="32"/><RowDefinition Height="200"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="48"/></Grid.RowDefinitions>
        <Grid><Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Text="Cleanup steps" FontSize="18" FontWeight="SemiBold"/><TextBlock x:Name="wStepCount" Grid.Column="1" Text="6 enabled" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/></Grid>
        <UniformGrid Grid.Row="1" Rows="2" Columns="3" Margin="-8,0,-8,0">
          <ToggleButton x:Name="wVM" Style="{StaticResource ActionTile}" IsChecked="True" AutomationProperties.Name="Delete VM or template"><StackPanel><TextBlock Text="vCenter" FontSize="15" FontWeight="SemiBold"/><TextBlock Text="VM / template" Foreground="{StaticResource Muted}" Margin="0,5,0,0" TextWrapping="Wrap" FontSize="11"/></StackPanel></ToggleButton>
          <ToggleButton x:Name="wSCCM" Style="{StaticResource ActionTile}" IsChecked="True" AutomationProperties.Name="Remove SCCM record"><StackPanel><TextBlock Text="SCCM" FontSize="15" FontWeight="SemiBold"/><TextBlock Text="Device record" Foreground="{StaticResource Muted}" Margin="0,5,0,0" TextWrapping="Wrap" FontSize="11"/></StackPanel></ToggleButton>
          <ToggleButton x:Name="wAD" Style="{StaticResource ActionTile}" IsChecked="True" AutomationProperties.Name="Delete Active Directory account"><StackPanel><TextBlock Text="Directory" FontSize="15" FontWeight="SemiBold"/><TextBlock Text="Computer account" Foreground="{StaticResource Muted}" Margin="0,5,0,0" TextWrapping="Wrap" FontSize="11"/></StackPanel></ToggleButton>
          <ToggleButton x:Name="wDNS" Style="{StaticResource ActionTile}" IsChecked="True" AutomationProperties.Name="Remove DNS records"><StackPanel><TextBlock Text="DNS" FontSize="15" FontWeight="SemiBold"/><TextBlock Text="Name records" Foreground="{StaticResource Muted}" Margin="0,5,0,0" TextWrapping="Wrap" FontSize="11"/></StackPanel></ToggleButton>
          <ToggleButton x:Name="wArchive" Style="{StaticResource ActionTile}" IsChecked="True" AutomationProperties.Name="Write SQL archive"><StackPanel><TextBlock Text="Archive" FontSize="15" FontWeight="SemiBold"/><TextBlock Text="SQL result" Foreground="{StaticResource Muted}" Margin="0,5,0,0" TextWrapping="Wrap" FontSize="11"/></StackPanel></ToggleButton>
          <ToggleButton x:Name="wEmail" Style="{StaticResource ActionTile}" IsChecked="True" AutomationProperties.Name="Send email notification"><StackPanel><TextBlock Text="Email" FontSize="15" FontWeight="SemiBold"/><TextBlock Text="Notification" Foreground="{StaticResource Muted}" Margin="0,5,0,0" TextWrapping="Wrap" FontSize="11"/></StackPanel></ToggleButton>
        </UniformGrid>
        <StackPanel Grid.Row="2" Margin="0,8,0,0"><TextBlock Text="Deletion reason" Foreground="{StaticResource Muted}" FontSize="12" Margin="0,0,0,7"/><TextBox x:Name="wReason" Style="{StaticResource Editor}" MaxLength="500" Height="36" AutomationProperties.Name="Deletion reason"/></StackPanel>
        <Grid Grid.Row="3" Margin="0,12,0,4" MinHeight="52"><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock x:Name="wCount" Text="0" FontSize="28" Foreground="{StaticResource Accent}" FontFamily="Segoe UI Light" VerticalAlignment="Center"/><StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="10,0,0,0"><TextBlock Text="machines selected" Foreground="{StaticResource Muted}" FontSize="12"/><TextBlock x:Name="wReady" TextTrimming="CharacterEllipsis" Text="Connect to begin" Foreground="{StaticResource Muted}" FontSize="11" Margin="0,3,0,0"/></StackPanel><Button x:Name="wResume" Grid.Column="2" Style="{StaticResource Ghost}" Content="Resume" Height="32" VerticalAlignment="Center"/></Grid>
        <Button x:Name="wDelete" Grid.Row="4" Style="{StaticResource DangerBtn}" Content="Delete selected VMs" IsEnabled="False" ToolTip="Review selected machines and actions before confirmation"/>
      </Grid></Border>
    </Grid>
    <Border Grid.Row="3" Background="{StaticResource Panel}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="14" Padding="18,14"><Grid><Grid.RowDefinitions><RowDefinition Height="28"/><RowDefinition Height="16"/><RowDefinition Height="*"/></Grid.RowDefinitions><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="24"/><ColumnDefinition Width="1.2*"/></Grid.ColumnDefinitions>
      <TextBlock x:Name="wNow" Text="Execution progress · ready when you are" FontSize="16" Grid.ColumnSpan="3"/>
      <ProgressBar x:Name="wProgress" Grid.Row="1" Grid.ColumnSpan="3" Style="{StaticResource Bar}" Maximum="100" VerticalAlignment="Top"/>
      <StackPanel x:Name="wTaskRows" Grid.Row="2" Margin="0,6,0,0"/>
      <Border Grid.Row="2" Grid.Column="2" Background="{StaticResource Bg}" CornerRadius="8" Padding="12,9" Margin="0,6,0,0"><TextBox x:Name="wLiveLog" Text="Progress and errors appear here during execution." IsReadOnly="True" Background="Transparent" BorderThickness="0" Foreground="{StaticResource Muted}" FontFamily="Consolas" FontSize="11" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" AutomationProperties.Name="Execution details"/></Border>
    </Grid></Border>
  </Grid>
</Window>
'@
    $script:WizardWindow = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$wizardXaml)))
    $script:WizardUi = @{}
    foreach ($match in [regex]::Matches($wizardXaml, 'x:Name="(w\w+)"')) {
        $script:WizardUi[$match.Groups[1].Value] = $script:WizardWindow.FindName($match.Groups[1].Value)
    }
    $script:WizardWindow.MaxHeight = [System.Windows.SystemParameters]::WorkArea.Height
    if ($script:WizardWindow.Height -gt $script:WizardWindow.MaxHeight) { $script:WizardWindow.Height = $script:WizardWindow.MaxHeight }
    $script:WizardWindow.MinHeight = [Math]::Min(860, $script:WizardWindow.MaxHeight)
    $script:WizardUi.wEnvironment.ItemsSource = @($script:vCenterServers | ForEach-Object { $_.Name.ToUpper() })
    $script:WizardUi.wEnvironment.SelectedIndex = $cboVCServer.SelectedIndex
    $script:WizardUi.wOS.ItemsSource = @('All','Windows','Linux','Other')
    $script:WizardUi.wOS.SelectedIndex = $cboOSFilter.SelectedIndex
    $script:WizardUi.wReason.Text = $txtReason.Text
    $script:WizardBindings = @{
        wVM=$chkVM; wSCCM=$chkSCCM; wAD=$chkAD; wDNS=$chkDNS
        wArchive=$chkArchive; wEmail=$chkEmail; wTemplates=$chkIncludeTemplates
    }
    foreach ($name in $script:WizardBindings.Keys) {
        $script:WizardUi[$name].IsChecked = $script:WizardBindings[$name].Checked
        $script:WizardUi[$name].Tag = $script:WizardBindings[$name]
        $script:WizardUi[$name].Add_Click({
            param($sender,$eventArgs)
            if (-not $script:WizardSyncing -and -not $script:WizardBusy) { $sender.Tag.Checked = [bool]$sender.IsChecked; Sync-DeletionWizard }
        })
    }
    $script:WizardTable = New-Object System.Data.DataTable
    [void]$script:WizardTable.Columns.Add('Selected',[bool])
    foreach ($name in @('Name','OS','Power')) { [void]$script:WizardTable.Columns.Add($name,[string]) }
    [void]$script:WizardTable.Columns.Add('LegacyItem',[object])
    [void]$lstVMs.Handle # ItemChecked must fire while the original form remains hidden.
    $script:WizardUi.wMachines.ItemsSource = $script:WizardTable.DefaultView
    $script:WizardUi.wMachines.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::CheckedEvent, [System.Windows.RoutedEventHandler]{param($sender,$eventArgs) Set-WizardMachineSelection -CheckBox $eventArgs.OriginalSource})
    $script:WizardUi.wMachines.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::UncheckedEvent, [System.Windows.RoutedEventHandler]{param($sender,$eventArgs) Set-WizardMachineSelection -CheckBox $eventArgs.OriginalSource})
    $script:WizardTaskBlocks = @()
    for ($i=0; $i -lt $tasks.Count; $i++) {
        $line = New-Object System.Windows.Controls.Grid
        $line.Height = 18
        foreach ($width in @('46','164','*')) { $column=New-Object System.Windows.Controls.ColumnDefinition; $column.Width=[System.Windows.GridLengthConverter]::new().ConvertFromString($width); [void]$line.ColumnDefinitions.Add($column) }
        $state=New-Object System.Windows.Controls.TextBlock; $state.FontSize=10; $state.Text='WAIT'; $state.Foreground=$script:WizardWindow.Resources['Muted']
        $title=New-Object System.Windows.Controls.TextBlock; $title.FontSize=11; $title.Text=$tasks[$i]; $title.Foreground=$script:WizardWindow.Resources['Text']; [System.Windows.Controls.Grid]::SetColumn($title,1)
        $result=New-Object System.Windows.Controls.TextBlock; $result.FontSize=10; $result.TextTrimming='CharacterEllipsis'; $result.Foreground=$script:WizardWindow.Resources['Muted']; [System.Windows.Controls.Grid]::SetColumn($result,2)
        [void]$line.Children.Add($state); [void]$line.Children.Add($title); [void]$line.Children.Add($result)
        [void]$script:WizardUi.wTaskRows.Children.Add($line)
        $script:WizardTaskBlocks += @{Status=$state; Result=$result}
    }
    $script:WizardUi.wEnvironment.Add_SelectionChanged({if(-not $script:WizardSyncing -and -not $script:WizardBusy){$cboVCServer.SelectedIndex=$script:WizardUi.wEnvironment.SelectedIndex}})
    $script:WizardUi.wOS.Add_SelectionChanged({if(-not $script:WizardSyncing -and -not $script:WizardBusy){$cboOSFilter.SelectedIndex=$script:WizardUi.wOS.SelectedIndex}})
    $script:WizardUi.wSearch.Add_TextChanged({if(-not $script:WizardSyncing -and -not $script:WizardBusy){$txtFilter.Text=$script:WizardUi.wSearch.Text}})
    $script:WizardUi.wReason.Add_TextChanged({if(-not $script:WizardSyncing -and -not $script:WizardBusy){$txtReason.Text=$script:WizardUi.wReason.Text}})
    $script:WizardUi.wConnect.Add_Click({Invoke-WizardAction -Button $btnConnect})
    $script:WizardUi.wRefresh.Add_Click({Invoke-WizardAction -Button $btnRefresh})
    $script:WizardUi.wImport.Add_Click({Invoke-WizardAction -Button $btnLoadCSV})
    $script:WizardUi.wClear.Add_Click({Invoke-WizardAction -Button $btnClearSelection})
    $script:WizardUi.wResume.Add_Click({Invoke-WizardAction -Button $btnResume})
    $script:WizardUi.wDelete.Add_Click({Invoke-WizardAction -Button $btnDelete})
    $script:WizardUi.wAdvanced.Add_Click({Show-WizardAdvanced})
    foreach($label in @($lblConnectionStatus,$lblVMProgress,$lblSelectedCount,$txtLogs)) { $label.Add_TextChanged({Sync-DeletionWizard}) }
    $script:WizardTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:WizardTimer.Interval = [TimeSpan]::FromMilliseconds(180)
    $script:WizardTimer.Add_Tick({Sync-DeletionWizard})
    $script:WizardWindow.Add_Closing({
        param($sender,$eventArgs)
        if ($script:WizardBusy) { $eventArgs.Cancel=$true; return }
        if ($script:vcConnected) {
            $answer=[System.Windows.MessageBox]::Show($script:WizardWindow,'Disconnect from vCenter and close?','Close tool',[System.Windows.MessageBoxButton]::YesNo,[System.Windows.MessageBoxImage]::Question)
            if($answer -ne [System.Windows.MessageBoxResult]::Yes){$eventArgs.Cancel=$true; return}
            try { Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction Stop; $script:vcConnected=$false }
            catch { [void][System.Windows.MessageBox]::Show($script:WizardWindow,$_.Exception.Message,'Disconnect failed'); $eventArgs.Cancel=$true }
        }
    })
    $script:WizardWindow.Add_Closed({$script:WizardTimer.Stop(); if($script:VMFilterTimer){$script:VMFilterTimer.Stop()}; $form.Dispose()})
    $lblArchiveStatus.Text = "SQL not checked yet. Click 'Check Table' when needed."
    Sync-DeletionWizard
}

function Set-WizardMachineSelection {
    param($CheckBox)
    if ($script:WizardSyncing -or $script:WizardBusy -or $CheckBox -isnot [System.Windows.Controls.CheckBox]) { return }
    $row=$CheckBox.DataContext
    if($row -is [System.Data.DataRowView]) {
        $row['Selected']=[bool]$CheckBox.IsChecked
        $row['LegacyItem'].Checked=[bool]$CheckBox.IsChecked
    }
}

function Invoke-WizardAction {
    param([System.Windows.Forms.Button]$Button)
    if($script:WizardBusy -or -not $Button.Enabled){return}
    $script:WizardBusy=$true
    try {
        Sync-DeletionWizard
        $script:WizardWindow.Dispatcher.Invoke([Action]{},[System.Windows.Threading.DispatcherPriority]::Render)
        # PerformClick cannot activate a button on an unshown form. Dispatch its
        # existing Click event after the same enabled check, on the GUI thread.
        $method=$Button.GetType().GetMethod('OnClick',[System.Reflection.BindingFlags]'Instance,NonPublic')
        [void]$method.Invoke($Button,@([System.EventArgs]::Empty))
    }
    catch {
        $failure=$_.Exception
        if($failure.InnerException){$failure=$failure.InnerException}
        Write-Log $failure.Message -Level 'ERROR'
        [void][System.Windows.MessageBox]::Show($script:WizardWindow,$failure.Message,'Operation failed',[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
    }
    finally {$script:WizardBusy=$false; Sync-DeletionWizard}
}

function Sync-DeletionWizard {
    if(-not $script:WizardWindow -or $script:WizardSyncing){return}
    $script:WizardSyncing=$true
    try {
        $ui=$script:WizardUi
        foreach($pair in @(@('wConnect',$btnConnect),@('wRefresh',$btnRefresh),@('wImport',$btnLoadCSV),@('wClear',$btnClearSelection),@('wResume',$btnResume))) {
            $ui[$pair[0]].IsEnabled=($pair[1].Enabled -and -not $script:WizardBusy)
        }
        $ui.wConnect.Content=$btnConnect.Text
        $ui.wEnvironment.IsEnabled=(-not $script:WizardBusy -and -not $script:vcConnected)
        $ui.wSearch.IsEnabled=($txtFilter.Enabled -and -not $script:WizardBusy)
        $ui.wOS.IsEnabled=($cboOSFilter.Enabled -and -not $script:WizardBusy)
        $ui.wReason.IsEnabled=($txtReason.Enabled -and -not $script:WizardBusy)
        $ui.wMachines.IsEnabled=($script:vcConnected -and -not $script:WizardBusy)
        $ui.wAdvanced.IsEnabled=(-not $script:WizardBusy)
        $ui.wDelete.IsEnabled=($btnDelete.Enabled -and $script:vcConnected -and $script:SelectedVMNames.Count -gt 0 -and -not $script:WizardBusy)
        $stepCount=0
        foreach($name in $script:WizardBindings.Keys) {
            $legacy=$script:WizardBindings[$name]
            $ui[$name].IsEnabled=($legacy.Enabled -and -not $script:WizardBusy)
            $ui[$name].IsChecked=$legacy.Checked
            if($name -ne 'wTemplates' -and $legacy.Checked){$stepCount++}
        }
        foreach($definition in (Get-DeletionStepDefinitions)){
            if(-not (Test-DeletionStepEnabled -Step $definition.Name)){
                $ui[$definition.Wizard].IsEnabled=$false
                $ui[$definition.Wizard].IsChecked=$false
                $ui[$definition.Wizard].ToolTip='Disabled in Advanced settings > Execution.'
            }
            else{$ui[$definition.Wizard].ToolTip=$(if($definition.Name -eq 'SCCM'){'Automatically skipped for Linux.'}else{'Skip this step for the current run by clearing this option.'})}
        }
        $ui.wStepCount.Text="$stepCount enabled"
        $ui.wCount.Text=[string]$script:SelectedVMNames.Count
        $ui.wReady.Text=$lblConnectionStatus.Text
        $ui.wReady.ToolTip=$lblConnectionStatus.Text
        $ui.wLoaded.Text=$lblVMCount.Text
        $ui.wNow.Text=$lblVMProgress.Text
        $ui.wProgress.Value=$progressBar.Value
        $ui.wStage1.Foreground=$script:WizardWindow.Resources[$(if($script:vcConnected){'Accent'}else{'Text'})]
        $ui.wStage2.Foreground=$script:WizardWindow.Resources[$(if($script:vcConnected){'Accent'}else{'Muted'})]
        $ui.wStage3.Foreground=$script:WizardWindow.Resources[$(if($script:SelectedVMNames.Count -gt 0){'Accent'}else{'Muted'})]
        $ui.wStage4.Foreground=$script:WizardWindow.Resources[$(if($script:WizardBusy){'Accent'}else{'Muted'})]
        for($i=0;$i -lt $script:WizardTaskBlocks.Count;$i++) {
            $script:WizardTaskBlocks[$i].Status.Text=$statusLabels[$i].Text
            $script:WizardTaskBlocks[$i].Result.Text=$resultLabels[$i].Text
            $script:WizardTaskBlocks[$i].Result.ToolTip=$resultLabels[$i].Text
            $c=$statusLabels[$i].ForeColor
            $script:WizardTaskBlocks[$i].Status.Foreground=New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb($c.R,$c.G,$c.B))
        }
        $logText=$txtLogs.Text
        if($logText.Length -gt 12000){$logText=$logText.Substring($logText.Length-12000)}
        if($logText -and $ui.wLiveLog.Text -ne $logText){$ui.wLiveLog.Text=$logText; $ui.wLiveLog.ScrollToEnd()}
        if(-not $script:IsUpdatingList) {
            $listKey=[string]$lstVMs.Items.Count
            if($lstVMs.Items.Count -gt 0){$listKey+=':'+$lstVMs.Items[0].GetHashCode()+':'+$lstVMs.Items[$lstVMs.Items.Count-1].GetHashCode()}
            $selectionKey=(@($script:SelectedVMNames | Sort-Object) -join "`n")
            if($listKey -ne $script:WizardListKey) {
                $script:WizardTable.Clear()
                foreach($item in $lstVMs.Items) {
                    $row=$script:WizardTable.NewRow(); $row['Selected']=($script:SelectedVMNames -contains $item.Text); $row['Name']=$item.Text
                    $row['OS']=$item.SubItems[2].Text; $row['Power']=$item.SubItems[1].Text; $row['LegacyItem']=$item
                    [void]$script:WizardTable.Rows.Add($row)
                }
                $script:WizardListKey=$listKey
            }
            elseif($selectionKey -ne $script:WizardSelectionKey) {
                foreach($row in $script:WizardTable.Rows){$row['Selected']=($script:SelectedVMNames -contains $row['Name'])}
            }
            $script:WizardSelectionKey=$selectionKey
        }
    }
    finally {$script:WizardSyncing=$false}
    if($script:WizardAdvanced){Sync-WizardAdvanced}
}


function Set-WizardAdvancedCaption {
    param([System.Windows.Window]$Window)
    try {
        if(-not ('VMDeletionAdvancedCaption' -as [type])){
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class VMDeletionAdvancedCaption {
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);
}
'@
        }
        $handle=(New-Object System.Windows.Interop.WindowInteropHelper($Window)).Handle
        $enabled=1
        $result=[VMDeletionAdvancedCaption]::DwmSetWindowAttribute($handle,20,[ref]$enabled,4)
        if($result -ne 0){[void][VMDeletionAdvancedCaption]::DwmSetWindowAttribute($handle,19,[ref]$enabled,4)}
    }
    catch { # Keep the window usable when this Windows version does not support caption theming.
    }
}

function New-WizardAdvancedAnimation {
    param($From, [double]$To, [int]$Milliseconds)
    $animation=New-Object System.Windows.Media.Animation.DoubleAnimation
    if($null -ne $From){$animation.From=[double]$From}
    $animation.To=$To
    $animation.Duration=[TimeSpan]::FromMilliseconds($Milliseconds)
    $easing=New-Object System.Windows.Media.Animation.CubicEase
    $easing.EasingMode='EaseOut'
    $animation.EasingFunction=$easing
    return $animation
}

function Reset-WizardAdvancedPage {
    param([int]$Index)
    foreach($prefix in @('aPageTransform','aGhostTransform','aContentTransform')){
        $transform=$script:AdvancedUi["$prefix$Index"]
        $transform.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty,$null)
        $transform.X=0
    }
    $page=$script:AdvancedUi["aPage$Index"]
    $page.BeginAnimation([System.Windows.UIElement]::OpacityProperty,$null)
    $page.Opacity=1
}

function Complete-WizardAdvancedTransition {
    if(-not $script:WizardAdvanced -or -not $script:AdvancedUi.aPageHost){return}
    if($script:AdvancedTransitionTimer){$script:AdvancedTransitionTimer.Stop()}
    foreach($index in 1..4){
        Reset-WizardAdvancedPage -Index $index
        $page=$script:AdvancedUi["aPage$index"]
        $active=($index -eq $script:AdvancedCurrentSection)
        $page.Visibility=$(if($active){'Visible'}else{'Collapsed'})
        $page.IsHitTestVisible=$active
        $page.IsEnabled=$active
    }
}

function Set-WizardAdvancedSection {
    param([ValidateRange(1,4)][int]$Index,[switch]$Immediate)
    if(-not $script:WizardAdvanced -or -not $script:AdvancedUi.aPageHost){return}
    if($Index -eq $script:AdvancedCurrentSection -and -not $Immediate){return}
    $from=$script:AdvancedCurrentSection
    if($script:AdvancedTransitionTimer){$script:AdvancedTransitionTimer.Stop()}
    $script:AdvancedCurrentSection=$Index
    foreach($number in 1..4){
        Reset-WizardAdvancedPage -Index $number
        $page=$script:AdvancedUi["aPage$number"]
        $page.Visibility=$(if($number -eq $from){'Visible'}else{'Collapsed'})
        $page.IsHitTestVisible=($number -eq $Index)
        $page.IsEnabled=($number -eq $Index)
    }
    $nav=@('aJumpConnections','aJumpExecution','aJumpEmail','aJumpSql')[$Index-1]
    $script:AdvancedUi[$nav].IsChecked=$true
    if($Immediate -or -not $script:WizardAdvanced.IsVisible -or -not [System.Windows.SystemParameters]::ClientAreaAnimation){
        Complete-WizardAdvancedTransition
        return
    }
    $width=[Math]::Max($script:AdvancedUi.aPageHost.ActualWidth,500)
    $direction=$(if($Index -gt $from){1}else{-1})
    $incoming=$script:AdvancedUi["aPage$Index"]
    $outgoing=$script:AdvancedUi["aPage$from"]
    [System.Windows.Controls.Panel]::SetZIndex($outgoing,1)
    [System.Windows.Controls.Panel]::SetZIndex($incoming,2)
    $script:AdvancedUi["aPageTransform$Index"].X=$direction*$width
    $script:AdvancedUi["aGhostTransform$Index"].X=-$direction*$width*0.50
    $script:AdvancedUi["aContentTransform$Index"].X=$direction*$width*0.14
    $incoming.Opacity=0
    $incoming.Visibility='Visible'
    $xProperty=[System.Windows.Media.TranslateTransform]::XProperty
    $opacityProperty=[System.Windows.UIElement]::OpacityProperty
    $script:AdvancedUi["aPageTransform$Index"].BeginAnimation($xProperty,(New-WizardAdvancedAnimation ($direction*$width) 0 430))
    $script:AdvancedUi["aGhostTransform$Index"].BeginAnimation($xProperty,(New-WizardAdvancedAnimation (-$direction*$width*0.50) 0 560))
    $script:AdvancedUi["aContentTransform$Index"].BeginAnimation($xProperty,(New-WizardAdvancedAnimation ($direction*$width*0.14) 0 480))
    $incoming.BeginAnimation($opacityProperty,(New-WizardAdvancedAnimation 0 1 260))
    $script:AdvancedUi["aPageTransform$from"].BeginAnimation($xProperty,(New-WizardAdvancedAnimation $null (-$direction*$width*0.30) 430))
    $outgoing.BeginAnimation($opacityProperty,(New-WizardAdvancedAnimation $null 0 300))
    # One reusable timer settles the latest transition; fast clicks cannot leave stale completion handlers.
    $script:AdvancedTransitionTimer.Start()
}


function Initialize-WizardAdvanced {
    $advancedXaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Advanced settings" Width="1060" Height="850" MinWidth="900" MinHeight="620" WindowStartupLocation="CenterOwner" Background="#0A1530" FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13" Foreground="#E8F0FF" UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
  <Window.Resources>
    <SolidColorBrush x:Key="Bg" Color="#0A1530" />
    <SolidColorBrush x:Key="Card" Color="#0F1F44" />
    <SolidColorBrush x:Key="FieldBg" Color="#0A1631" />
    <SolidColorBrush x:Key="Line" Color="#22396B" />
    <SolidColorBrush x:Key="Text" Color="#E8F0FF" />
    <SolidColorBrush x:Key="Muted" Color="#8CA2C9" />
    <SolidColorBrush x:Key="Accent" Color="#62E4F4" />
    <SolidColorBrush x:Key="AccentDim" Color="#14425F" />
    <Style x:Key="NoRepeat" TargetType="RepeatButton">
      <Setter Property="Focusable" Value="False" />
      <Setter Property="IsTabStop" Value="False" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RepeatButton">
            <Border Background="Transparent" />
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="8" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Track x:Name="PART_Track" Orientation="{TemplateBinding Orientation}" IsDirectionReversed="True">
              <Track.DecreaseRepeatButton>
                <RepeatButton Command="ScrollBar.PageUpCommand" Opacity="0" Focusable="False" />
              </Track.DecreaseRepeatButton>
              <Track.Thumb>
                <Thumb>
                  <Thumb.Template>
                    <ControlTemplate TargetType="Thumb">
                      <Border Background="#2F4A7D" CornerRadius="4" Margin="1" />
                    </ControlTemplate>
                  </Thumb.Template>
                </Thumb>
              </Track.Thumb>
              <Track.IncreaseRepeatButton>
                <RepeatButton Command="ScrollBar.PageDownCommand" Opacity="0" Focusable="False" />
              </Track.IncreaseRepeatButton>
            </Track>
            <ControlTemplate.Triggers>
              <Trigger Property="Orientation" Value="Horizontal">
                <Setter TargetName="PART_Track" Property="IsDirectionReversed" Value="False" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto" />
          <Setter Property="Height" Value="8" />
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="Primary" TargetType="Button">
      <Setter Property="Foreground" Value="#0A1530" />
      <Setter Property="Cursor" Value="Hand" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" CornerRadius="10" Background="{StaticResource Accent}" Padding="18,10">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.86" />
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.35" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Outline" TargetType="Button">
      <Setter Property="Foreground" Value="{StaticResource Text}" />
      <Setter Property="Cursor" Value="Hand" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" CornerRadius="10" BorderThickness="1" BorderBrush="#2B4577" Background="Transparent" Padding="18,10">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}" />
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.4" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Tab" TargetType="RadioButton">
      <Setter Property="Foreground" Value="{StaticResource Text}" />
      <Setter Property="Cursor" Value="Hand" />
      <Setter Property="Margin" Value="0,0,10,0" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="bd" CornerRadius="10" BorderThickness="1" BorderBrush="#2B4577" Background="Transparent" Padding="18,9">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#4A6699" />
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentDim}" />
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}" />
                <Setter Property="Foreground" Value="{StaticResource Accent}" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Field" TargetType="TextBox">
      <Setter Property="Foreground" Value="{StaticResource Text}" />
      <Setter Property="CaretBrush" Value="{StaticResource Accent}" />
      <Setter Property="SelectionBrush" Value="#2B6C8F" />
      <Setter Property="Height" Value="40" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{StaticResource FieldBg}" BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="12">
              <ScrollViewer x:Name="PART_ContentHost" Margin="14,0" VerticalAlignment="Center" />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsReadOnly" Value="True">
                <Setter Property="Foreground" Value="#B7C6E4" />
              </Trigger>
              <MultiTrigger>
                <MultiTrigger.Conditions>
                  <Condition Property="IsKeyboardFocused" Value="True" />
                  <Condition Property="IsReadOnly" Value="False" />
                </MultiTrigger.Conditions>
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}" />
              </MultiTrigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Tick" TargetType="CheckBox">
      <Setter Property="Cursor" Value="Hand" />
      <Setter Property="Foreground" Value="#B7C6E4" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Border x:Name="bd" Width="20" Height="20" CornerRadius="6" BorderThickness="1.5" BorderBrush="#4A6699" Background="Transparent" VerticalAlignment="Center">
                <Path x:Name="tick" Data="M4,9.5 L8,13.5 L14.5,5.5" Stroke="#0A1530" StrokeThickness="2" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed" />
              </Border>
              <ContentPresenter Margin="10,0,0,0" VerticalAlignment="Center" />
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource Accent}" />
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Accent}" />
                <Setter TargetName="tick" Property="Visibility" Value="Visible" />
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.55" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="StepTile" TargetType="CheckBox">
      <Setter Property="Cursor" Value="Hand"/><Setter Property="Margin" Value="0,0,12,12"/>
      <Setter Property="MinHeight" Value="100"/><Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="CheckBox">
        <Border x:Name="tile" CornerRadius="12" Background="{StaticResource FieldBg}" BorderBrush="{StaticResource Line}" BorderThickness="1" Padding="14">
          <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="30"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border x:Name="box" Width="20" Height="20" CornerRadius="6" BorderThickness="1.5" BorderBrush="#4A6699" VerticalAlignment="Top" Background="Transparent">
              <Path x:Name="mark" Data="M4,9.5 L8,13.5 L14.5,5.5" Stroke="#0A1530" StrokeThickness="2" Visibility="Collapsed"/>
            </Border><ContentPresenter Grid.Column="1" HorizontalAlignment="Stretch"/>
          </Grid>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsChecked" Value="True"><Setter TargetName="box" Property="Background" Value="{StaticResource Accent}"/><Setter TargetName="box" Property="BorderBrush" Value="{StaticResource Accent}"/><Setter TargetName="mark" Property="Visibility" Value="Visible"/></Trigger>
          <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="tile" Property="BorderBrush" Value="#4A6699"/></Trigger>
          <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="tile" Property="BorderBrush" Value="{StaticResource Accent}"/><Setter TargetName="tile" Property="BorderThickness" Value="2"/></Trigger>
          <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.55"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Lbl" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Muted}" />
      <Setter Property="FontSize" Value="12" />
      <Setter Property="Margin" Value="0,0,0,6" />
    </Style>
    <Style x:Key="GridHeader" TargetType="DataGridColumnHeader">
      <Setter Property="Foreground" Value="{StaticResource Muted}" />
      <Setter Property="Background" Value="Transparent" />
      <Setter Property="Padding" Value="12,0,0,8" />
      <Setter Property="HorizontalContentAlignment" Value="Left" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridColumnHeader">
            <Border BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1" Padding="{TemplateBinding Padding}" Height="32">
              <ContentPresenter VerticalAlignment="Bottom" />
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="GridCell" TargetType="DataGridCell">
      <Setter Property="Foreground" Value="{StaticResource Text}" />
      <Setter Property="BorderThickness" Value="0" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="Transparent" Padding="12,0">
              <ContentPresenter VerticalAlignment="Center" />
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="GridRow" TargetType="DataGridRow">
      <Setter Property="Background" Value="Transparent" />
      <Setter Property="Foreground" Value="{StaticResource Text}" />
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#16295A" />
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#16295A" />
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="AdvancedPassword" TargetType="PasswordBox">
      <Setter Property="Foreground" Value="#E8F0FF" />
      <Setter Property="CaretBrush" Value="#62E4F4" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="PasswordBox">
            <Border x:Name="edge" Background="#0C1730" BorderBrush="#2B4577" BorderThickness="1" CornerRadius="8" Padding="10,7">
              <ScrollViewer x:Name="PART_ContentHost" />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="edge" Property="BorderBrush" Value="#62E4F4" />
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.65" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="AdvancedList" TargetType="ListBox">
      <Setter Property="Background" Value="#0C1730" />
      <Setter Property="Foreground" Value="#E8F0FF" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBox">
            <Border Background="{TemplateBinding Background}" BorderBrush="#2B4577" BorderThickness="1" CornerRadius="8" Padding="5">
              <ScrollViewer Focusable="False" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                <ItemsPresenter />
              </ScrollViewer>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ListBoxItem">
      <Setter Property="Foreground" Value="#E8F0FF" />
      <Setter Property="Padding" Value="8,6" />
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border x:Name="item" CornerRadius="5" Background="Transparent" Padding="{TemplateBinding Padding}">
              <ContentPresenter />
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="item" Property="Background" Value="#14425F" />
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Grid Margin="26,22,26,22">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>
    <!-- Title and actions -->
    <Grid Grid.Row="0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />
        <ColumnDefinition Width="Auto" />
      </Grid.ColumnDefinitions>
      <StackPanel>
        <TextBlock Text="Advanced settings" FontFamily="Segoe UI Light" FontSize="32" />
        <TextBlock x:Name="aLockStatus" Foreground="{StaticResource Muted}" FontSize="12" Margin="2,6,0,0" TextTrimming="CharacterEllipsis" ToolTip="{Binding Text,RelativeSource={RelativeSource Self}}" />
      </StackPanel>
      <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Top">
        <Button x:Name="aSaveSettings" Content="Save configuration" Style="{StaticResource Primary}" />
        <Button x:Name="aResetSettings" Content="Reset configuration" Style="{StaticResource Outline}" Margin="10,0,0,0" />
        <Button x:Name="advancedUnlock" Content="Unlock settings" Style="{StaticResource Outline}" Margin="10,0,0,0" />
      </StackPanel>
    </Grid>
    <!-- Section buttons -->
    <WrapPanel Grid.Row="1" Margin="0,18,0,0">
      <RadioButton x:Name="aJumpConnections" Tag="1" Content="Connections" Style="{StaticResource Tab}" GroupName="sec" IsChecked="True" />
      <RadioButton x:Name="aJumpExecution" Tag="2" Content="Execution" Style="{StaticResource Tab}" GroupName="sec" />
      <RadioButton x:Name="aJumpEmail" Tag="3" Content="Email" Style="{StaticResource Tab}" GroupName="sec" />
      <RadioButton x:Name="aJumpSql" Tag="4" Content="SQL &amp; archive" Style="{StaticResource Tab}" GroupName="sec" />
    </WrapPanel>
    <!-- Pages live on top of each other; the code slides them -->
    <Grid x:Name="aPageHost" Grid.Row="2" Margin="0,16,0,0" ClipToBounds="True">
      <!-- 01 Connections -->
      <Grid x:Name="aPage1">
        <Grid.RenderTransform>
          <TranslateTransform x:Name="aPageTransform1" />
        </Grid.RenderTransform>
        <TextBlock Text="01" FontFamily="Segoe UI Light" FontSize="150" Foreground="#FFFFFF" Opacity="0.05" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,-8,46,0" IsHitTestVisible="False" Panel.ZIndex="5">
          <TextBlock.RenderTransform>
            <TranslateTransform x:Name="aGhostTransform1" />
          </TextBlock.RenderTransform>
        </TextBlock>
        <ScrollViewer x:Name="aScroll" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" CanContentScroll="False">
          <ScrollViewer.RenderTransform>
            <TranslateTransform x:Name="aContentTransform1" />
          </ScrollViewer.RenderTransform>
          <Border x:Name="aConnectionsSection" Background="{StaticResource Card}" BorderBrush="{DynamicResource Line}" BorderThickness="1" CornerRadius="18" Padding="26,22,26,24" Margin="0,0,14,0" VerticalAlignment="Top">
            <StackPanel>
              <Grid Margin="0,0,0,18">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="40" />
                  <ColumnDefinition />
                </Grid.ColumnDefinitions>
                <Border Width="28" Height="28" CornerRadius="8" Background="{DynamicResource AccentDim}" VerticalAlignment="Top">
                  <TextBlock Text="01" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                </Border>
                <StackPanel Grid.Column="1">
                  <TextBlock Text="Infrastructure connections" FontSize="20" FontWeight="SemiBold" />
                  <TextBlock Text="Servers used by the existing cleanup workflow." Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,4,0,0" TextWrapping="Wrap" />
                </StackPanel>
              </Grid>
              <UniformGrid Columns="2">
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="SCCM server" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aSccm" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="SCCM server" />
                </StackPanel>
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="SCCM site code" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aSite" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="SCCM site code" />
                </StackPanel>
              </UniformGrid>
              <UniformGrid Columns="2">
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="DNS server" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aDns" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="DNS server" />
                </StackPanel>
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="DNS zone" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aZone" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="DNS zone" />
                </StackPanel>
              </UniformGrid>
              <UniformGrid Columns="2">
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="Domain controller" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aDomain" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="Domain controller" />
                </StackPanel>
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="Notification email" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aNotification" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="Notification email" />
                </StackPanel>
              </UniformGrid>
            </StackPanel>
          </Border>
        </ScrollViewer>
      </Grid>
      <!-- 02 Execution -->
      <Grid x:Name="aPage2" Visibility="Collapsed">
        <Grid.RenderTransform>
          <TranslateTransform x:Name="aPageTransform2" />
        </Grid.RenderTransform>
        <TextBlock Text="02" FontFamily="Segoe UI Light" FontSize="150" Foreground="#FFFFFF" Opacity="0.05" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,-8,46,0" IsHitTestVisible="False" Panel.ZIndex="5">
          <TextBlock.RenderTransform>
            <TranslateTransform x:Name="aGhostTransform2" />
          </TextBlock.RenderTransform>
        </TextBlock>
        <ScrollViewer x:Name="aPageScroll2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" CanContentScroll="False">
          <ScrollViewer.RenderTransform>
            <TranslateTransform x:Name="aContentTransform2" />
          </ScrollViewer.RenderTransform>
          <Border x:Name="aExecutionSection" Background="{StaticResource Card}" BorderBrush="{DynamicResource Line}" BorderThickness="1" CornerRadius="18" Padding="26,22,26,24" Margin="0,0,14,0" VerticalAlignment="Top">
            <StackPanel>
              <Grid Margin="0,0,0,18">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="40" />
                  <ColumnDefinition />
                </Grid.ColumnDefinitions>
                <Border Width="28" Height="28" CornerRadius="8" Background="{DynamicResource AccentDim}" VerticalAlignment="Top">
                  <TextBlock Text="02" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                </Border>
                <StackPanel Grid.Column="1">
                  <TextBlock Text="Execution options" FontSize="20" FontWeight="SemiBold" />
                  <TextBlock Text="Logging, confirmation and VM removal preferences." Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,4,0,0" TextWrapping="Wrap" />
                </StackPanel>
              </Grid>
              <UniformGrid Columns="2">
                <CheckBox x:Name="aLogging" Content="Enable logging" Style="{StaticResource Tick}" Margin="0,0,14,14" />
                <CheckBox x:Name="aConfirmation" Content="Require deletion confirmation" Style="{StaticResource Tick}" Margin="0,0,14,14" />
              </UniformGrid>
              <TextBlock Text="Default workflow steps" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,6"/>
              <TextBlock Text="Clear a step to disable it for all future runs, then choose Save configuration. You can enable it again here. Linux skips SCCM only; AD follows your selection." Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,0,12,16"/>
              <UniformGrid Columns="3" Margin="0,0,0,16">
                <CheckBox x:Name="aStepSccm" Style="{StaticResource StepTile}" AutomationProperties.Name="Enable SCCM for future runs"><StackPanel><TextBlock Text="SCCM" FontWeight="SemiBold" Foreground="{StaticResource Text}" Margin="0,0,0,6"/><TextBlock Text="Remove the device record. Skipped for Linux." FontSize="11" Foreground="{StaticResource Muted}" TextWrapping="Wrap"/></StackPanel></CheckBox>
                <CheckBox x:Name="aStepAd" Style="{StaticResource StepTile}" AutomationProperties.Name="Enable Active Directory for future runs"><StackPanel><TextBlock Text="Active Directory" FontWeight="SemiBold" Foreground="{StaticResource Text}" Margin="0,0,0,6"/><TextBlock Text="Remove the computer account, including Linux when selected." FontSize="11" Foreground="{StaticResource Muted}" TextWrapping="Wrap"/></StackPanel></CheckBox>
                <CheckBox x:Name="aStepDns" Style="{StaticResource StepTile}" AutomationProperties.Name="Enable DNS for future runs"><StackPanel><TextBlock Text="DNS" FontWeight="SemiBold" Foreground="{StaticResource Text}" Margin="0,0,0,6"/><TextBlock Text="Remove the server DNS records." FontSize="11" Foreground="{StaticResource Muted}" TextWrapping="Wrap"/></StackPanel></CheckBox>
                <CheckBox x:Name="aStepVm" Style="{StaticResource StepTile}" AutomationProperties.Name="Enable VM / template for future runs"><StackPanel><TextBlock Text="VM / template" FontWeight="SemiBold" Foreground="{StaticResource Text}" Margin="0,0,0,6"/><TextBlock Text="Remove the VM or template from vCenter." FontSize="11" Foreground="{StaticResource Muted}" TextWrapping="Wrap"/></StackPanel></CheckBox>
                <CheckBox x:Name="aStepSql" Style="{StaticResource StepTile}" AutomationProperties.Name="Enable SQL archive for future runs"><StackPanel><TextBlock Text="SQL archive" FontWeight="SemiBold" Foreground="{StaticResource Text}" Margin="0,0,0,6"/><TextBlock Text="Write each result using an independent SQL connection." FontSize="11" Foreground="{StaticResource Muted}" TextWrapping="Wrap"/></StackPanel></CheckBox>
                <CheckBox x:Name="aStepEmail" Style="{StaticResource StepTile}" AutomationProperties.Name="Enable Email for future runs"><StackPanel><TextBlock Text="Email" FontWeight="SemiBold" Foreground="{StaticResource Text}" Margin="0,0,0,6"/><TextBlock Text="Send personal reports and the batch summary." FontSize="11" Foreground="{StaticResource Muted}" TextWrapping="Wrap"/></StackPanel></CheckBox>
              </UniformGrid>
              <Border Height="1" Background="{StaticResource Line}" Margin="0,0,12,20"/>
              <CheckBox x:Name="aPermanent" Content="Delete VM files from disk permanently" Style="{StaticResource Tick}" Margin="0,0,14,14" />
              <UniformGrid Columns="2">
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="Power-off timeout (10–300 seconds)" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aTimeout" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="Power-off timeout (10–300 seconds)" />
                </StackPanel>
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="Log folder" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aLogPath" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="Log folder" />
                </StackPanel>
              </UniformGrid>
              <StackPanel Orientation="Horizontal">
                <Button x:Name="aBrowse" Content="Browse log folder" Style="{StaticResource Outline}" Height="40" Margin="0,0,10,0" />
              </StackPanel>
            </StackPanel>
          </Border>
        </ScrollViewer>
      </Grid>
      <!-- 03 Email -->
      <Grid x:Name="aPage3" Visibility="Collapsed">
        <Grid.RenderTransform>
          <TranslateTransform x:Name="aPageTransform3" />
        </Grid.RenderTransform>
        <TextBlock Text="03" FontFamily="Segoe UI Light" FontSize="150" Foreground="#FFFFFF" Opacity="0.05" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,-8,46,0" IsHitTestVisible="False" Panel.ZIndex="5">
          <TextBlock.RenderTransform>
            <TranslateTransform x:Name="aGhostTransform3" />
          </TextBlock.RenderTransform>
        </TextBlock>
        <ScrollViewer x:Name="aPageScroll3" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" CanContentScroll="False">
          <ScrollViewer.RenderTransform>
            <TranslateTransform x:Name="aContentTransform3" />
          </ScrollViewer.RenderTransform>
          <Border x:Name="aEmailSection" Background="{StaticResource Card}" BorderBrush="{DynamicResource Line}" BorderThickness="1" CornerRadius="18" Padding="26,22,26,24" Margin="0,0,14,0" VerticalAlignment="Top">
            <StackPanel>
              <Grid Margin="0,0,0,18">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="40" />
                  <ColumnDefinition />
                </Grid.ColumnDefinitions>
                <Border Width="28" Height="28" CornerRadius="8" Background="{DynamicResource AccentDim}" VerticalAlignment="Top">
                  <TextBlock Text="03" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                </Border>
                <StackPanel Grid.Column="1">
                  <TextBlock Text="Email notifications" FontSize="20" FontWeight="SemiBold" />
                  <TextBlock Text="SMTP delivery and recipient groups, organized by operating system." Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,4,0,0" TextWrapping="Wrap" />
                </StackPanel>
              </Grid>
              <CheckBox x:Name="aEnableEmail" Content="Enable email notifications" Style="{StaticResource Tick}" Margin="0,0,14,14" />
              <UniformGrid Columns="3">
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="SMTP server" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aSmtp" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="SMTP server" />
                </StackPanel>
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="SMTP port (1–65535)" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aPort" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="SMTP port (1–65535)" />
                </StackPanel>
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="From address" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aFrom" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="From address" />
                </StackPanel>
              </UniformGrid>
              <CheckBox x:Name="aSsl" Content="Use SSL for SMTP" Style="{StaticResource Tick}" Margin="0,0,14,14" />
              <UniformGrid Columns="3">
                <StackPanel Margin="0,0,14,16">
                  <TextBlock Text="All machines" Foreground="{DynamicResource Muted}" Margin="0,0,0,7" />
                  <ListBox x:Name="aCommon" Style="{DynamicResource AdvancedList}" Height="104" AutomationProperties.Name="All machines" />
                </StackPanel>
                <StackPanel Margin="0,0,14,16">
                  <TextBlock Text="Windows only" Foreground="{DynamicResource Muted}" Margin="0,0,0,7" />
                  <ListBox x:Name="aWindows" Style="{DynamicResource AdvancedList}" Height="104" AutomationProperties.Name="Windows only" />
                </StackPanel>
                <StackPanel Margin="0,0,14,16">
                  <TextBlock Text="Linux only" Foreground="{DynamicResource Muted}" Margin="0,0,0,7" />
                  <ListBox x:Name="aLinux" Style="{DynamicResource AdvancedList}" Height="104" AutomationProperties.Name="Linux only" />
                </StackPanel>
              </UniformGrid>
              <UniformGrid Columns="2">
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="Group or email to add" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <TextBox x:Name="aNewGroup" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="Group or email to add" />
                </StackPanel>
                <StackPanel Margin="0,0,14,14">
                  <TextBlock Text="Recipient group" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                  <ComboBox x:Name="aGroupType" Style="{DynamicResource Picker}" Height="40" AutomationProperties.Name="Recipient group" />
                </StackPanel>
              </UniformGrid>
              <StackPanel Orientation="Horizontal">
                <Button x:Name="aAddGroup" Content="Add group" Style="{StaticResource Outline}" Height="40" Margin="0,0,10,0" />
                <Button x:Name="aRemoveGroup" Content="Remove selected" Style="{StaticResource Outline}" Height="40" Margin="0,0,10,0" />
                <Button x:Name="aTestEmail" Content="Send test email" Style="{StaticResource Outline}" Height="40" Margin="0,0,10,0" />
              </StackPanel>
            </StackPanel>
          </Border>
        </ScrollViewer>
      </Grid>
      <!-- 04 SQL and archive: settings on top, the table gets all the remaining height -->
      <Grid x:Name="aPage4" Visibility="Collapsed">
        <Grid.RenderTransform>
          <TranslateTransform x:Name="aPageTransform4" />
        </Grid.RenderTransform>
        <ScrollViewer x:Name="aPageScroll4" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" CanContentScroll="False">
          <ScrollViewer.RenderTransform>
            <TranslateTransform x:Name="aContentTransform4" />
          </ScrollViewer.RenderTransform>
          <StackPanel>
            <Border x:Name="aSqlSection" Background="{StaticResource Card}" BorderBrush="{DynamicResource Line}" BorderThickness="1" CornerRadius="18" Padding="26,22,26,24" Margin="0,0,14,0" VerticalAlignment="Top">
              <Grid>
                <TextBlock Text="04" FontFamily="Segoe UI Light" FontSize="150" Foreground="#FFFFFF" Opacity="0.05" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,-8,46,0" IsHitTestVisible="False" Panel.ZIndex="5">
                  <TextBlock.RenderTransform>
                    <TranslateTransform x:Name="aGhostTransform4" />
                  </TextBlock.RenderTransform>
                </TextBlock>
                <StackPanel>
                  <Grid Margin="0,0,0,18">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="40" />
                      <ColumnDefinition />
                    </Grid.ColumnDefinitions>
                    <Border Width="28" Height="28" CornerRadius="8" Background="{DynamicResource AccentDim}" VerticalAlignment="Top">
                      <TextBlock Text="04" Foreground="{DynamicResource Accent}" HorizontalAlignment="Center" VerticalAlignment="Center" FontSize="11" />
                    </Border>
                    <StackPanel Grid.Column="1">
                      <TextBlock Text="SQL archive connection" FontSize="20" FontWeight="SemiBold" />
                      <TextBlock Text="Connection settings for archive validation, queries and exports." Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,4,0,0" TextWrapping="Wrap" />
                    </StackPanel>
                  </Grid>
                  <UniformGrid Columns="3">
                    <StackPanel Margin="0,0,14,14">
                      <TextBlock Text="SQL server" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                      <TextBox x:Name="aSqlServer" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="SQL server" />
                    </StackPanel>
                    <StackPanel Margin="0,0,14,14">
                      <TextBlock Text="Database" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                      <TextBox x:Name="aDatabase" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="Database" />
                    </StackPanel>
                    <StackPanel Margin="0,0,14,14">
                      <TextBlock Text="Table" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                      <TextBox x:Name="aTable" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="Table" />
                    </StackPanel>
                  </UniformGrid>
                  <CheckBox x:Name="aIntegrated" Content="Use Windows authentication" Style="{StaticResource Tick}" Margin="0,0,14,14" />
                  <UniformGrid Columns="2">
                    <StackPanel Margin="0,0,14,14">
                      <TextBlock Text="SQL username" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                      <TextBox x:Name="aSqlUser" Style="{StaticResource Field}" Height="40" AutomationProperties.Name="SQL username" />
                    </StackPanel>
                    <StackPanel Margin="0,0,14,14">
                      <TextBlock Text="SQL password" Foreground="{DynamicResource Muted}" FontSize="12" Margin="0,0,0,7" />
                      <PasswordBox x:Name="aSqlPassword" Style="{DynamicResource AdvancedPassword}" Height="40" AutomationProperties.Name="SQL password" />
                    </StackPanel>
                  </UniformGrid>
                  <WrapPanel>
                    <Button x:Name="aCheckTable" Content="Check connection and table" Style="{StaticResource Primary}" Height="40" Margin="0,0,10,8" />
                    <Button x:Name="aSaveSql" Content="Save SQL settings" Style="{StaticResource Outline}" Height="40" Margin="0,0,10,8" />
                    <Button x:Name="aExportTable" Content="Export table" ToolTip="Read the complete current SQL table and export CSV or XLSX" Style="{StaticResource Outline}" Height="40" Margin="0,0,10,8" />
                    <Button x:Name="aOpenSqlTable" Content="Display SQL table" ToolTip="Display every current SQL record below the connection settings" Style="{StaticResource Outline}" Height="40" Margin="0,0,10,8" />
                  </WrapPanel>
                  <TextBlock x:Name="aSqlStatus" Foreground="{DynamicResource Muted}" Margin="0,14,0,0" TextWrapping="Wrap" />
                </StackPanel>
              </Grid>
            </Border>
            <Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="aSqlResults" Visibility="Collapsed" CornerRadius="18" Background="{StaticResource Card}" BorderBrush="{StaticResource Line}" BorderThickness="1" Padding="22,18" Margin="0,16,14,0">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto" />
                  <RowDefinition Height="Auto" />
                  <RowDefinition Height="240" />
                  <RowDefinition Height="Auto" />
                  <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>
                <Grid Margin="0,0,0,16">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="Auto" />
                  </Grid.ColumnDefinitions>
                  <StackPanel VerticalAlignment="Center" Margin="0,0,16,0">
                    <TextBlock Text="SQL table" FontSize="19" FontWeight="SemiBold" />
                    <TextBlock x:Name="tableSource" Text="Current SQL records" Foreground="{StaticResource Muted}" FontSize="12" Margin="0,5,0,0" TextTrimming="CharacterEllipsis" ToolTip="{Binding Text,RelativeSource={RelativeSource Self}}" />
                  </StackPanel>
                  <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="tableRefresh" Content="Refresh from SQL" Style="{StaticResource Outline}" Height="40" Margin="0,0,10,0" />
                    <Button x:Name="tableExport" Content="Export table" Style="{StaticResource Primary}" Height="40" ToolTip="Export the displayed records to CSV or XLSX" />
                  </StackPanel>
                </Grid>
                <DataGrid x:Name="tableGrid" Grid.Row="2" AutoGenerateColumns="True" IsReadOnly="True" CanUserAddRows="False" CanUserDeleteRows="False" CanUserResizeRows="False" HeadersVisibility="Column" GridLinesVisibility="None" Background="Transparent" BorderThickness="0" RowBackground="Transparent" RowHeight="34" MinColumnWidth="130" ColumnHeaderStyle="{StaticResource GridHeader}" CellStyle="{StaticResource GridCell}" RowStyle="{StaticResource GridRow}" EnableRowVirtualization="True" EnableColumnVirtualization="True" ScrollViewer.CanContentScroll="True" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" AutomationProperties.Name="Current SQL archive records" />
                <TextBlock x:Name="tableEmpty" Grid.Row="2" Text="No records loaded yet." Foreground="{StaticResource Muted}" HorizontalAlignment="Center" VerticalAlignment="Center" IsHitTestVisible="False" />
                <TextBlock x:Name="tableCount" Grid.Row="3" Text="No records loaded yet." Foreground="{StaticResource Muted}" FontSize="12" Margin="0,12,0,0" TextWrapping="Wrap" />
                <TextBlock x:Name="aSqlTableError" Grid.Row="4" Foreground="#F0647A" Margin="0,6,0,0" TextWrapping="Wrap" />
                <StackPanel xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,12">
                  <TextBlock Text="Search table" Foreground="{StaticResource Muted}" VerticalAlignment="Center" Margin="0,0,10,0" />
                  <TextBox x:Name="tableSearch" Width="250" Height="36" Style="{StaticResource Field}" AutomationProperties.Name="Search SQL table" />
                  <Button x:Name="tableClearSearch" Content="Clear" Height="36" Style="{StaticResource Outline}" Margin="10,0,0,0" />
                </StackPanel>
              </Grid>
            </Border>
          </StackPanel>
        </ScrollViewer>
      </Grid>
    </Grid>
    <StackPanel Grid.Row="3">
      <TextBlock x:Name="aSettingsStatus" Foreground="#A7B9D9" TextWrapping="Wrap" Margin="0,8,0,0" />
      <TextBlock x:Name="aValidation" Foreground="#F0647A" TextWrapping="Wrap" Margin="0,4,0,0" />
    </StackPanel>
  </Grid>
</Window>
'@
    $script:WizardAdvanced=[Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$advancedXaml)))
    [void]$script:WizardAdvanced.Resources.MergedDictionaries.Add($script:WizardWindow.Resources)
    $script:WizardAdvanced.Owner=$script:WizardWindow
    $script:WizardAdvanced.Add_SourceInitialized({param($sender,$eventArgs) Set-WizardAdvancedCaption -Window $sender})
    $script:WizardAdvanced.MaxHeight=[System.Windows.SystemParameters]::WorkArea.Height
    $script:WizardAdvanced.MinHeight=[Math]::Min(620,$script:WizardAdvanced.MaxHeight)
    $script:WizardAdvanced.Height=[Math]::Min(850,$script:WizardAdvanced.MaxHeight)
    $script:WizardAdvanced.MaxWidth=[System.Windows.SystemParameters]::WorkArea.Width
    $script:WizardAdvanced.MinWidth=[Math]::Min(900,$script:WizardAdvanced.MaxWidth)
    $script:WizardAdvanced.Width=[Math]::Min(1060,$script:WizardAdvanced.MaxWidth)
    $script:AdvancedUi=@{}
    foreach($match in [regex]::Matches($advancedXaml,'x:Name="([a-zA-Z]\w+)"')) {
        $script:AdvancedUi[$match.Groups[1].Value]=$script:WizardAdvanced.FindName($match.Groups[1].Value)
    }
    $script:AdvancedSyncing=$false
    $script:AdvancedFields=@{}
    $fieldDefinitions=@(
        @('aSccm',$txtSettingsSCCM,'Text'),@('aSite',$txtSettingsSiteCode,'Text'),
        @('aDns',$txtSettingsDNS,'Text'),@('aZone',$txtSettingsDNSZone,'Text'),
        @('aDomain',$txtSettingsDC,'Text'),@('aNotification',$txtSettingsEmail,'Text'),
        @('aLogPath',$txtLogPath,'Text'),@('aSmtp',$txtSMTPServer,'Text'),
        @('aFrom',$txtSMTPFrom,'Text'),@('aNewGroup',$txtNewGroup,'Text'),
        @('aSqlServer',$txtSQLServer,'Text'),@('aDatabase',$txtSQLDatabase,'Text'),
        @('aTable',$txtSQLTable,'Text'),@('aSqlUser',$txtSQLUser,'Text'),@('aSqlPassword',$txtSQLPass,'Password'),
        @('aTimeout',$numTimeout,'Number'),@('aPort',$numSMTPPort,'Number'),
        @('aLogging',$chkEnableLogging,'Check'),@('aConfirmation',$chkConfirmDeletion,'Check'),
        @('aPermanent',$chkDeleteFromDisk,'Check'),@('aEnableEmail',$chkEnableEmail,'Check'),
        @('aSsl',$chkSMTPSSL,'Check'),@('aIntegrated',$chkSQLIntegrated,'Check'),
        @('aGroupType',$cboGroupType,'Picker'),
        @('aCommon',$script:lstGroupsCommon,'List'),@('aWindows',$script:lstGroupsWindows,'List'),@('aLinux',$script:lstGroupsLinux,'List')
    )
    foreach($step in (Get-DeletionStepDefinitions)){$fieldDefinitions+=,@($step.Editor,$script:PersistentStepControls[$step.Name],'Check')}
    foreach($definition in $fieldDefinitions) {
        $editor=$script:AdvancedUi[$definition[0]]
        $binding=@{Name=$definition[0]; Legacy=$definition[1]; Kind=$definition[2]; Editor=$editor; ListKey=$null}
        $script:AdvancedFields[$definition[0]]=$binding
        $editor.Tag=$binding
        switch($binding.Kind) {
            'Text' { $editor.Add_TextChanged({param($sender,$eventArgs) if(-not $script:AdvancedSyncing -and -not $script:WizardBusy -and $sender.Tag.Legacy.Enabled){$sender.Tag.Legacy.Text=$sender.Text}}) }
            'Password' { $editor.Add_PasswordChanged({param($sender,$eventArgs) if(-not $script:AdvancedSyncing -and -not $script:WizardBusy -and $sender.Tag.Legacy.Enabled){$sender.Tag.Legacy.Text=$sender.Password}}) }
            'Check' { $editor.Add_Click({param($sender,$eventArgs) if(-not $script:AdvancedSyncing -and -not $script:WizardBusy -and $sender.Tag.Legacy.Enabled){$sender.Tag.Legacy.Checked=[bool]$sender.IsChecked; Sync-WizardAdvanced}}) }
            'Number' {
                $editor.Add_TextChanged({
                    param($sender,$eventArgs)
                    if($script:AdvancedSyncing -or $script:WizardBusy -or -not $sender.Tag.Legacy.Enabled){return}
                    $value=0
                    if([int]::TryParse($sender.Text,[ref]$value) -and $value -ge $sender.Tag.Legacy.Minimum -and $value -le $sender.Tag.Legacy.Maximum) {
                        $sender.Tag.Legacy.Value=$value; $sender.ToolTip=$null
                    }
                    else {$sender.ToolTip="Enter a whole number from $($sender.Tag.Legacy.Minimum) to $($sender.Tag.Legacy.Maximum)."}
                })
            }
            'Picker' {
                $editor.ItemsSource=@($binding.Legacy.Items | ForEach-Object {$_.ToString()})
                $editor.Add_SelectionChanged({param($sender,$eventArgs) if(-not $script:AdvancedSyncing -and -not $script:WizardBusy -and $sender.Tag.Legacy.Enabled){$sender.Tag.Legacy.SelectedIndex=$sender.SelectedIndex}})
            }
            'List' {
                $editor.Add_SelectionChanged({
                    param($sender,$eventArgs)
                    if($script:AdvancedSyncing -or $script:WizardBusy -or -not $sender.Tag.Legacy.Enabled){return}
                    $script:AdvancedSyncing=$true
                    try {
                        foreach($entry in $script:AdvancedFields.Values) {
                            if($entry.Kind -eq 'List' -and $entry.Editor -ne $sender){$entry.Legacy.ClearSelected(); $entry.Editor.SelectedIndex=-1}
                        }
                        $sender.Tag.Legacy.SelectedIndex=$sender.SelectedIndex
                    }
                    finally {$script:AdvancedSyncing=$false}
                })
            }
        }
    }
    $script:AdvancedActions=@{}
    foreach($definition in @(
        @('aBrowse',$btnBrowseLog,$false),@('aAddGroup',$btnAddGroup,$false),@('aRemoveGroup',$btnRemoveGroup,$false),
        @('aTestEmail',$btnTestEmail,$true),@('aSaveSettings',$btnSaveSettings,$true),@('aResetSettings',$btnResetSettings,$false),
        @('aCheckTable',$btnCheckTable,$false),@('aSaveSql',$btnSaveSQLSettings,$false)
    )) {
        $button=$script:AdvancedUi[$definition[0]]
        $button.Tag=@{Legacy=$definition[1]; ValidateNumbers=$definition[2]}
        $script:AdvancedActions[$definition[0]]=$button
        $button.Add_Click({param($sender,$eventArgs) Invoke-WizardAdvancedAction -Button $sender})
    }
    $script:AdvancedCurrentSection=1
    $script:AdvancedTransitionTimer=New-Object System.Windows.Threading.DispatcherTimer
    $script:AdvancedTransitionTimer.Interval=[TimeSpan]::FromMilliseconds(600)
    $script:AdvancedTransitionTimer.Add_Tick({Complete-WizardAdvancedTransition})
    foreach($definition in @(@('aJumpConnections',1),@('aJumpExecution',2),@('aJumpEmail',3),@('aJumpSql',4))) {
        $script:AdvancedUi[$definition[0]].Tag=[int]$definition[1]
        $script:AdvancedUi[$definition[0]].Add_Checked({param($sender,$eventArgs) Set-WizardAdvancedSection -Index ([int]$sender.Tag)})
    }
    $script:WizardAdvanced.Add_IsVisibleChanged({param($sender,$eventArgs) if(-not $sender.IsVisible){Complete-WizardAdvancedTransition}})
    $script:WizardAdvanced.Add_SizeChanged({if($script:AdvancedTransitionTimer.IsEnabled){Complete-WizardAdvancedTransition}})
    Set-WizardAdvancedSection -Index 1 -Immediate
    $script:AdvancedUi.advancedUnlock.Add_Click({Request-WizardAdvancedUnlock})
    $script:AdvancedUi.aExportTable.Add_Click({Invoke-WizardSqlTableAction -Action Export})
    $script:AdvancedUi.aOpenSqlTable.Add_Click({Invoke-WizardSqlTableAction -Action Open})
    $script:AdvancedUi.tableRefresh.Add_Click({Invoke-WizardSqlTableAction -Action Refresh})
    $script:AdvancedUi.tableExport.Add_Click({Invoke-WizardSqlTableAction -Action Export -Snapshot (Get-WizardSqlVisibleSnapshot)})
    $script:AdvancedUi.tableSearch.Add_TextChanged({Update-WizardSqlSearch})
    $script:AdvancedUi.tableClearSearch.Add_Click({$script:AdvancedUi.tableSearch.Clear()})
    $script:WizardAdvanced.Add_Closing({param($sender,$eventArgs) if($script:WizardWindow.IsVisible){$eventArgs.Cancel=$true; $sender.Hide()}})
    Sync-WizardAdvanced -Force
}

function Invoke-WizardAdvancedAction {
    param([System.Windows.Controls.Button]$Button)
    if($script:WizardBusy -or -not $Button.Tag.Legacy.Enabled){return}
    $script:AdvancedUi.aValidation.Text=''
    if($Button.Tag.ValidateNumbers) {
        foreach($name in @('aTimeout','aPort')) {
            $binding=$script:AdvancedFields[$name]; $value=0
            if(-not [int]::TryParse($binding.Editor.Text,[ref]$value) -or $value -lt $binding.Legacy.Minimum -or $value -gt $binding.Legacy.Maximum) {
                $script:AdvancedUi.aValidation.Text="$($binding.Editor.GetValue([System.Windows.Automation.AutomationProperties]::NameProperty)) must be a whole number from $($binding.Legacy.Minimum) to $($binding.Legacy.Maximum)."
                Set-WizardAdvancedSection -Index $(if($name -eq 'aTimeout'){2}else{3}) -Immediate
                $binding.Editor.BringIntoView(); [void]$binding.Editor.Focus(); return
            }
        }
    }
    Invoke-WizardAction -Button $Button.Tag.Legacy
    Sync-WizardAdvanced -Force
}

function Sync-WizardAdvanced {
    param([switch]$Force)
    if(-not $script:WizardAdvanced -or $script:AdvancedSyncing){return}
    if(-not $Force -and -not $script:WizardAdvanced.IsVisible){return}
    $script:AdvancedSyncing=$true
    try {
        foreach($binding in $script:AdvancedFields.Values) {
            $editor=$binding.Editor; $legacy=$binding.Legacy
            $canEdit=($legacy.Enabled -and -not $script:WizardBusy)
            switch($binding.Kind) {
                'Text' {
                    $editor.IsReadOnly=(-not $canEdit)
                    if(($Force -or -not $editor.IsKeyboardFocusWithin) -and $editor.Text -ne $legacy.Text){$editor.Text=$legacy.Text}
                }
                'Number' {
                    $editor.IsReadOnly=(-not $canEdit)
                    if(($Force -or -not $editor.IsKeyboardFocusWithin) -and $editor.Text -ne [string]$legacy.Value){$editor.Text=[string]$legacy.Value}
                }
                'Password' {
                    $editor.IsEnabled=$canEdit
                    if(($Force -or -not $editor.IsKeyboardFocusWithin) -and $editor.Password -ne $legacy.Text){$editor.Password=$legacy.Text}
                }
                'Check' {$editor.IsEnabled=$canEdit; $editor.IsChecked=$legacy.Checked}
                'Picker' {$editor.IsEnabled=$canEdit; $editor.SelectedIndex=$legacy.SelectedIndex}
                'List' {
                    $editor.IsEnabled=$canEdit
                    $key=(@($legacy.Items | ForEach-Object {$_.ToString()}) -join "`n")
                    if($key -cne $binding.ListKey) {
                        $editor.ItemsSource=@($legacy.Items | ForEach-Object {$_.ToString()})
                        $binding.ListKey=$key
                    }
                    $editor.SelectedIndex=$legacy.SelectedIndex
                }
            }
        }
        foreach($button in $script:AdvancedActions.Values){$button.IsEnabled=($button.Tag.Legacy.Enabled -and -not $script:WizardBusy)}
        $unlocked=($tabArchive.Enabled -and $tabSettings.Enabled)
        $script:AdvancedUi.advancedUnlock.Content=$(if($unlocked){'Lock settings'}else{'Unlock settings'})
        $script:AdvancedUi.advancedUnlock.IsEnabled=(-not $script:WizardBusy)
        $script:AdvancedUi.aExportTable.IsEnabled=($unlocked -and -not $script:WizardBusy)
        $script:AdvancedUi.aOpenSqlTable.IsEnabled=($unlocked -and -not $script:WizardBusy)
        $script:AdvancedUi.tableRefresh.IsEnabled=($unlocked -and -not $script:WizardBusy)
        $script:AdvancedUi.tableSearch.IsEnabled=($unlocked -and -not $script:WizardBusy -and $null -ne $script:AdvancedUi.aSqlResults.Tag)
        $script:AdvancedUi.tableClearSearch.IsEnabled=($script:AdvancedUi.tableSearch.IsEnabled -and $script:AdvancedUi.tableSearch.Text.Length -gt 0)
        $script:AdvancedUi.tableExport.IsEnabled=($unlocked -and -not $script:WizardBusy -and $null -ne $script:AdvancedUi.aSqlResults.Tag)
        $script:AdvancedUi.aLockStatus.Text=$(if($unlocked){'Unlocked · edits are saved only when you choose Save.'}else{'Locked · unlock to edit configuration and query the archive.'})
        $script:AdvancedUi.aSettingsStatus.Text=$lblSettingsStatus.Text
        $script:AdvancedUi.aSqlStatus.Text=$lblArchiveStatus.Text
    }
    finally {$script:AdvancedSyncing=$false}
}






function Get-WizardSqlTableSnapshot {
    if(-not $tabArchive.Enabled -or -not $tabSettings.Enabled){throw 'Unlock advanced settings before reading SQL data.'}
    $quotedTable=ConvertTo-WizardSqlIdentifier -TableName $txtSQLTable.Text
    $serverName=$txtSQLServer.Text.Trim(); $databaseName=$txtSQLDatabase.Text.Trim()
    $sqlCredential=$null
    try {
        if(-not $chkSQLIntegrated.Checked){$sqlCredential=ConvertTo-DeletionSqlCredential -UserName $txtSQLUser.Text -Secret $txtSQLPass.Text}
        $connectionString=New-WizardSqlConnectionString -Server $serverName -Database $databaseName -IntegratedSecurity $chkSQLIntegrated.Checked -Credential $sqlCredential
    }
    finally {if($sqlCredential){$sqlCredential.Password.Dispose()}}
    $connection=$null; $command=$null; $adapter=$null
    try {
        $connection=New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $connection.Open()
        $command=$connection.CreateCommand()
        $command.CommandTimeout=60
        # Identifiers are validated and quoted. This operation never creates or modifies a table.
        $command.CommandText='SELECT * FROM '+$quotedTable
        $adapter=New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $data=New-Object System.Data.DataTable
        [void]$adapter.Fill($data)
        return [pscustomobject]@{Data=$data; Server=$serverName; Database=$databaseName; Table=$txtSQLTable.Text.Trim(); LoadedAt=(Get-Date)}
    }
    finally {
        if($adapter){$adapter.Dispose()}; if($command){$command.Dispose()}; if($connection){$connection.Dispose()}
        $connectionString=$null
    }
}

function ConvertTo-WizardExportText {
    param($Value)
    if($null -eq $Value -or $Value -is [System.DBNull]){return ''}
    if($Value -is [DateTime]){return $Value.ToString('o',[Globalization.CultureInfo]::InvariantCulture)}
    if($Value -is [byte[]]){return [Convert]::ToBase64String($Value)}
    if($Value -is [IFormattable]){return $Value.ToString($null,[Globalization.CultureInfo]::InvariantCulture)}
    return [string]$Value
}

function Add-WizardXlsxEntry {
    param($Archive,[string]$Name,[string]$Content)
    $entry=$Archive.CreateEntry($Name)
    $stream=$entry.Open(); $writer=New-Object IO.StreamWriter($stream,(New-Object Text.UTF8Encoding($false)))
    try{$writer.Write($Content)}finally{$writer.Dispose()}
}

function Write-WizardXlsxCell {
    param($Writer,$Value,[switch]$Header)
    $Writer.WriteStartElement('c')
    if($Header){$Writer.WriteAttributeString('s','1')}
    if($null -eq $Value -or $Value -is [DBNull]){$Writer.WriteEndElement(); return}
    $text=ConvertTo-WizardExportText $Value
    if($text.Length -gt 32767){throw 'A table value exceeds the XLSX cell length limit. Choose CSV to preserve the complete value.'}
    $isNumeric=($Value -is [byte] -or $Value -is [sbyte] -or $Value -is [int16] -or $Value -is [uint16] -or $Value -is [int32] -or $Value -is [uint32] -or $Value -is [int64] -or $Value -is [uint64] -or $Value -is [decimal] -or $Value -is [single] -or $Value -is [double])
    # Excel preserves at most 15 significant digits; keep longer identifiers as text.
    $significantDigits=($text -replace '[^0-9]','').TrimStart('0').Length
    if($Value -is [bool]) {
        $Writer.WriteAttributeString('t','b'); $Writer.WriteElementString('v',$(if($Value){'1'}else{'0'}))
    }
    elseif(-not $Header -and $isNumeric -and $significantDigits -le 15 -and $text -match '^-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?$') {
        $Writer.WriteAttributeString('t','n'); $Writer.WriteElementString('v',$text)
    }
    else {
        # inlineStr keeps leading zeroes and formula-like text as literal values.
        $text=[regex]::Replace($text,'_x[0-9A-Fa-f]{4}_',{param($m) '_x005F_'+$m.Value.Substring(1)})
        $text=[regex]::Replace($text,'[\x00-\x08\x0B\x0C\x0E-\x1F]',{param($m) '_x{0:X4}_' -f [int][char]$m.Value[0]})
        $Writer.WriteAttributeString('t','inlineStr'); $Writer.WriteStartElement('is'); $Writer.WriteStartElement('t')
        $Writer.WriteAttributeString('xml','space','http://www.w3.org/XML/1998/namespace','preserve')
        $Writer.WriteString($text); $Writer.WriteEndElement(); $Writer.WriteEndElement()
    }
    $Writer.WriteEndElement()
}

function Export-WizardSqlTableData {
    param([Parameter(Mandatory=$true)][AllowEmptyCollection()][System.Data.DataTable]$Data,[Parameter(Mandatory=$true)][string]$FilePath)
    $destination=[IO.Path]::GetFullPath($FilePath)
    $extension=[IO.Path]::GetExtension($destination).ToLowerInvariant()
    if($extension -notin @('.csv','.xlsx')){throw 'Choose a .csv or .xlsx filename.'}
    if($Data.Columns.Count -eq 0){throw 'The query returned no column definitions.'}
    if($extension -eq '.xlsx' -and ($Data.Rows.Count -gt 1048575 -or $Data.Columns.Count -gt 16384)){throw 'This table exceeds XLSX limits. Choose CSV to export the complete table.'}
    $folder=[IO.Path]::GetDirectoryName($destination)
    $temporary=Join-Path $folder ('.vm-table-'+[Guid]::NewGuid().ToString('N')+'.tmp')
    $backup=Join-Path $folder ('.vm-table-'+[Guid]::NewGuid().ToString('N')+'.bak')
    try {
        if($extension -eq '.csv') {
            $writer=New-Object IO.StreamWriter($temporary,$false,(New-Object Text.UTF8Encoding($true)))
            try {
                $headers=@(foreach($column in $Data.Columns){'"'+$column.ColumnName.Replace('"','""')+'"'})
                $writer.WriteLine(($headers -join ','))
                foreach($row in $Data.Rows) {
                    $cells=@(foreach($column in $Data.Columns){'"'+(ConvertTo-WizardExportText $row[$column]).Replace('"','""')+'"'})
                    $writer.WriteLine(($cells -join ','))
                }
            }
            finally{$writer.Dispose()}
        }
        else {
            Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
            $zip=[IO.Compression.ZipFile]::Open($temporary,[IO.Compression.ZipArchiveMode]::Create)
            try {
                Add-WizardXlsxEntry $zip '[Content_Types].xml' '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
                Add-WizardXlsxEntry $zip '_rels/.rels' '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
                Add-WizardXlsxEntry $zip 'xl/workbook.xml' '<?xml version="1.0" encoding="utf-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SQL table" sheetId="1" r:id="rId1"/></sheets></workbook>'
                Add-WizardXlsxEntry $zip 'xl/_rels/workbook.xml.rels' '<?xml version="1.0" encoding="utf-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
                Add-WizardXlsxEntry $zip 'xl/styles.xml' '<?xml version="1.0" encoding="utf-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFE8F0FF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF12234A"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
                $entry=$zip.CreateEntry('xl/worksheets/sheet1.xml')
                $stream=$entry.Open(); $settings=New-Object Xml.XmlWriterSettings
                $settings.Encoding=New-Object Text.UTF8Encoding($false); $settings.CloseOutput=$true
                $xml=[Xml.XmlWriter]::Create($stream,$settings)
                try {
                    $last=''; $index=$Data.Columns.Count
                    while($index -gt 0){$index--; $last=[char](65+($index%26))+$last; $index=[int][Math]::Floor($index/26)}
                    $range='A1:'+ $last + [string]($Data.Rows.Count+1)
                    $xml.WriteStartDocument(); $xml.WriteStartElement('worksheet','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
                    $xml.WriteStartElement('dimension'); $xml.WriteAttributeString('ref',$range); $xml.WriteEndElement()
                    $xml.WriteStartElement('sheetViews'); $xml.WriteStartElement('sheetView'); $xml.WriteAttributeString('workbookViewId','0')
                    $xml.WriteStartElement('pane'); $xml.WriteAttributeString('ySplit','1'); $xml.WriteAttributeString('topLeftCell','A2'); $xml.WriteAttributeString('activePane','bottomLeft'); $xml.WriteAttributeString('state','frozen'); $xml.WriteEndElement(); $xml.WriteEndElement(); $xml.WriteEndElement()
                    $xml.WriteStartElement('cols')
                    for($i=1;$i -le $Data.Columns.Count;$i++){$xml.WriteStartElement('col'); $xml.WriteAttributeString('min',[string]$i); $xml.WriteAttributeString('max',[string]$i); $xml.WriteAttributeString('width',[string][Math]::Min(42,[Math]::Max(18,$Data.Columns[$i-1].ColumnName.Length+4))); $xml.WriteAttributeString('customWidth','1'); $xml.WriteEndElement()}
                    $xml.WriteEndElement(); $xml.WriteStartElement('sheetData'); $xml.WriteStartElement('row'); $xml.WriteAttributeString('r','1')
                    foreach($column in $Data.Columns){Write-WizardXlsxCell -Writer $xml -Value $column.ColumnName -Header}
                    $xml.WriteEndElement(); $rowNumber=2
                    foreach($row in $Data.Rows){$xml.WriteStartElement('row'); $xml.WriteAttributeString('r',[string]$rowNumber); foreach($column in $Data.Columns){Write-WizardXlsxCell -Writer $xml -Value $row[$column]}; $xml.WriteEndElement(); $rowNumber++}
                    $xml.WriteEndElement(); $xml.WriteStartElement('autoFilter'); $xml.WriteAttributeString('ref',$range); $xml.WriteEndElement(); $xml.WriteEndElement(); $xml.WriteEndDocument()
                }
                finally{$xml.Dispose()}
            }
            finally{$zip.Dispose()}
        }
        if([IO.File]::Exists($destination)){[IO.File]::Replace($temporary,$destination,$backup)}else{[IO.File]::Move($temporary,$destination)}
        return [pscustomobject]@{Success=$true; FilePath=$destination; Rows=$Data.Rows.Count; Columns=$Data.Columns.Count}
    }
    finally {
        if([IO.File]::Exists($temporary)){[IO.File]::Delete($temporary)}
        if([IO.File]::Exists($backup)){[IO.File]::Delete($backup)}
    }
}

function Set-WizardSqlViewerSnapshot {
    param($Window,$Snapshot)
    $Window.FindName('aSqlResults').Tag=$Snapshot
    $script:SqlTableView=New-Object System.Data.DataView($Snapshot.Data)
    $Window.FindName('tableGrid').ItemsSource=$script:SqlTableView
    $Window.FindName('tableSource').Text="$($Snapshot.Server) / $($Snapshot.Database) / $($Snapshot.Table)"
    $Window.FindName('aSqlTableError').Text=''
    Update-WizardSqlSearch
}

function Show-WizardSqlTableViewer {
    param($Snapshot)
    # The viewer lives below the SQL settings card in the existing advanced window.
    if($null -ne $Snapshot){Set-WizardSqlViewerSnapshot -Window $script:WizardAdvanced -Snapshot $Snapshot}
    $script:AdvancedUi.aSqlResults.Visibility='Visible'
    Set-WizardAdvancedSection -Index 4 -Immediate
    Sync-WizardAdvanced -Force
    $script:WizardAdvanced.UpdateLayout()
    $script:AdvancedUi.aSqlResults.BringIntoView()
}

function Invoke-WizardSqlTableAction {
    param([ValidateSet('Open','Refresh','Export')][string]$Action,$Owner=$script:WizardAdvanced,$Snapshot=$null)
    if($script:WizardBusy -or -not $tabArchive.Enabled -or -not $tabSettings.Enabled){return}
    $destination=$null
    if($Action -eq 'Export') {
        $dialog=New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Title='Export SQL table'
        $dialog.Filter='Excel workbook (*.xlsx)|*.xlsx|CSV UTF-8 (*.csv)|*.csv'
        $dialog.DefaultExt='.xlsx'; $dialog.AddExtension=$true; $dialog.OverwritePrompt=$true
        $safeName=($txtSQLTable.Text -replace '[^\p{L}\p{N}_-]','_')
        $dialog.FileName="SQL_${safeName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
        if($dialog.ShowDialog($Owner) -ne $true){return}
        $destination=$dialog.FileName
    }
    $script:WizardBusy=$true
    $loaded=$null
    try {
        $lblArchiveStatus.Text=$(if($Action -eq 'Export'){'Preparing full-table export...'}else{'Reading all current records from SQL...'})
        Sync-DeletionWizard
        $script:WizardWindow.Dispatcher.Invoke([Action]{},[System.Windows.Threading.DispatcherPriority]::Render)
        if($null -eq $Snapshot){$loaded=Get-WizardSqlTableSnapshot}else{$loaded=$Snapshot}
        if($Action -eq 'Export') {
            $exported=Export-WizardSqlTableData -Data $loaded.Data -FilePath $destination
            $lblArchiveStatus.Text="Exported $($exported.Rows) records to $($exported.FilePath)"
        }
        else {
            $dgvArchive.DataSource=$loaded.Data
            $lblRecordCount.Text="$($loaded.Data.Rows.Count) records"
            $lblArchiveStatus.Text="Loaded all $($loaded.Data.Rows.Count) records from $($loaded.Database).$($loaded.Table)."
            if($Action -eq 'Refresh'){Set-WizardSqlViewerSnapshot -Window $Owner -Snapshot $loaded}
        }
    }
    catch {
        $lblArchiveStatus.Text="SQL table $($Action.ToLowerInvariant()) failed: $($_.Exception.Message)"
        Write-Log $lblArchiveStatus.Text -Level 'ERROR'
        if($Action -eq 'Export') {
            [void][System.Windows.MessageBox]::Show($Owner,$lblArchiveStatus.Text,'SQL table',[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
        }
        else {
            $script:AdvancedUi.aSqlTableError.Text=$lblArchiveStatus.Text
            Show-WizardSqlTableViewer
        }
        $loaded=$null
    }
    finally {$script:WizardBusy=$false; Sync-DeletionWizard; Sync-WizardAdvanced -Force}
    if($Action -eq 'Open' -and $null -ne $loaded){Show-WizardSqlTableViewer -Snapshot $loaded}
    if($Action -eq 'Export' -and $null -ne $loaded){[void][System.Windows.MessageBox]::Show($Owner,$lblArchiveStatus.Text,'Export completed',[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information)}
}

function Request-WizardAdvancedUnlock {
    if($script:WizardBusy){return}
    if($tabArchive.Enabled -or $tabSettings.Enabled) {
        $tabArchive.Enabled=$false; $tabSettings.Enabled=$false
        $btnUnlockSettings.Text='Unlock settings'
        Sync-WizardAdvanced -Force; return
    }
    $unlockXaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Unlock settings" Width="410" Height="270" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" Background="#0C1730" Foreground="#E8F0FF" FontFamily="Segoe UI"><Border Background="#12234A" CornerRadius="14" BorderBrush="#2B4577" BorderThickness="1" Margin="16" Padding="20"><StackPanel><TextBlock Text="Unlock protected settings" FontSize="19"/><TextBlock Text="Enter the existing settings password." Foreground="#A7B9D9" Margin="0,8,0,12"/><PasswordBox x:Name="unlockPassword" Height="36" Style="{DynamicResource AdvancedPassword}"/><TextBlock x:Name="unlockError" Foreground="#F0647A" Margin="0,8,0,10"/><StackPanel Orientation="Horizontal" HorizontalAlignment="Right"><Button x:Name="unlockCancel" Content="Cancel" IsCancel="True" Style="{DynamicResource Ghost}" Height="34" Margin="0,0,8,0"/><Button x:Name="unlockConfirm" Content="Unlock" IsDefault="True" Style="{DynamicResource Pill}" Height="34"/></StackPanel></StackPanel></Border></Window>
'@
    $script:WizardUnlockDialog=[Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$unlockXaml)))
    [void]$script:WizardUnlockDialog.Resources.MergedDictionaries.Add($script:WizardAdvanced.Resources)
    $script:WizardUnlockDialog.Owner=$script:WizardAdvanced
    $script:WizardUnlockDialog.FindName('unlockConfirm').Add_Click({
        if($script:WizardUnlockDialog.FindName('unlockPassword').Password -eq $script:settingsPassword) {
            $tabArchive.Enabled=$true; $tabSettings.Enabled=$true
            $btnUnlockSettings.Text='Lock settings'
            $script:WizardUnlockDialog.DialogResult=$true
        }
        else {$script:WizardUnlockDialog.FindName('unlockError').Text='Incorrect password. Please try again.'}
    })
    $script:WizardUnlockDialog.Add_ContentRendered({[void]$script:WizardUnlockDialog.FindName('unlockPassword').Focus()})
    [void]$script:WizardUnlockDialog.ShowDialog()
    $script:WizardUnlockDialog.FindName('unlockPassword').Clear()
    $script:WizardUnlockDialog=$null
    Sync-WizardAdvanced -Force
}

function Show-WizardAdvanced {
    if($script:WizardBusy){return}
    if(-not $script:WizardAdvanced){Initialize-WizardAdvanced}
    Sync-WizardAdvanced -Force
    [void]$script:WizardAdvanced.ShowDialog()
}

#endregion WPF single-screen wizard

Initialize-DeletionWizard
$script:WizardTimer.Start()
[void]$script:WizardWindow.ShowDialog()


