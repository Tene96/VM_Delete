$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Data
$source=[IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) 'VM_DELETION_FINAL.ps1'))
$ast=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$null)
$names=@('Get-DeletionStepDefinitions','Test-DeletionStepEnabled','Get-DeletionStepSkipReason','Get-DeletionOSType','Sync-DeletionStepControls','Save-DeletionStepPreferences','New-VmDeletionReportHtml','ConvertTo-VmDeletionReportRows','Get-VmDeletionReportVCenter','Resolve-DeletionMailRecipients','Get-DeletionSummaryRecipients','Send-DeletionSummaryReport','Send-VMDeletionNotification','Get-DeletionAnnotationValues','Get-DeletionOwnerMetadata','Get-VMDetailsForArchive','Save-VMDeletionToArchive','Write-VMTrackingRecord','Get-SQLConnectionString','ConvertTo-DeletionSqlCredential','New-WizardSqlConnectionString','ConvertTo-WizardSqlIdentifier','Get-WizardSqlSearchFilter')
foreach($name in $names){$f=$ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)|Select-Object -First 1;if(-not $f){throw "Missing function $name"};Invoke-Expression $f.Extent.Text}
function Assert([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message}}
$script:Logs=New-Object Collections.ArrayList
function Write-Log{param($Message,$Level) [void]$script:Logs.Add([string]$Message)}
$script:DefaultSettings=@{SQLServer='sql.fixture.invalid';SQLDatabase='FixtureArchive';SQLTable='dbo.DeletedVMs';SQLIntegratedAuth=$false;SQLUsername='fixture';SQLPassword='fixture;"password';SMTPServer='smtp.fixture.invalid';SMTPFrom='tool@example.invalid';SMTPPort=2525;SMTPUseSSL=$true;EmailSubjectPrefix='[Fixture]';EnableEmailNotification=$true;DomainController='dc.fixture.invalid';GroupsCommon='common@example.invalid, DUP@example.invalid';GroupsWindows='WinTeam;dup@example.invalid';GroupsLinux='linux@example.invalid';AdminAlertEmail='old-admin@example.invalid'}
$global:DefaultVIServer=[pscustomobject]@{Name='vc.fixture.invalid'}
$script:allVMsData=@()
Add-Type -ReferencedAssemblies System.Data -TypeDefinition @'
using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
public static class FixtureSql {
 public static string Mode="ok", Query;
 public static int Opened, Executed, Commits, Rollbacks, Disposed;
 public static Dictionary<string,object> Values=new Dictionary<string,object>();
 public static void Reset(string mode){Mode=mode;Opened=Executed=Commits=Rollbacks=Disposed=0;Values.Clear();Query=null;}
 public class Connection:IDisposable {
  public void Open(){Opened++;}
  public Transaction BeginTransaction(){return new Transaction();}
  public Command CreateCommand(){return new Command();}
  public void Dispose(){Disposed++;}
 }
 public class Transaction:IDisposable {
  public void Commit(){if(Mode=="commit-fail")throw new Exception("Fixture commit failure");Commits++;}
  public void Rollback(){Rollbacks++;}
  public void Dispose(){}
 }
 public class Command:IDisposable {
  private SqlCommand backing=new SqlCommand();
  public string CommandText{get;set;}
  public int CommandTimeout{get;set;}
  public Transaction Transaction{get;set;}
  public SqlParameterCollection Parameters{get{return backing.Parameters;}}
  public int ExecuteNonQuery(){
   Executed++;Query=CommandText;
   foreach(SqlParameter p in Parameters)Values[p.ParameterName]=p.Value;
   if(Mode=="fail")throw new Exception("Fixture insert denied");
   Parameters["@InsertedRows"].Value=Mode=="zero"?0:1;
   return -1;
  }
  public void Dispose(){backing.Dispose();}
 }
}
'@
function New-DeletionSqlConnection{param($ConnectionString) return (New-Object FixtureSql+Connection)}
$details=@{VMName="VM-O'Brien";ServerName="VM-O'Brien";VCenterName='vc.fixture.invalid';Owner='owner';MemoryGB=4;CPUCores=2;OperatingSystem='Windows'}
[FixtureSql]::Reset('ok');$script:sqlTableExists=$false
$saved=Save-VMDeletionToArchive -VMDetails $details -Reason "Operator's test" -VMDeleted $true -Status Success
Assert ($saved.Success -and [FixtureSql]::Executed -eq 1 -and [FixtureSql]::Commits -eq 1 -and [FixtureSql]::Disposed -eq 1) 'Insert must commit exactly once even when ExecuteNonQuery returns -1'
Assert ([FixtureSql]::Query -match 'INSERT INTO \[dbo\]\.\[DeletedVMs\]' -and [FixtureSql]::Query -notmatch "VM-O'Brien") 'Archive SQL is not parameterized/identifier quoted'
Assert ([FixtureSql]::Values['@ServerName'] -eq "VM-O'Brien" -and [FixtureSql]::Values['@MemoryGB'] -eq 4 -and [FixtureSql]::Values['@PowerStateDate'] -is [DBNull]) 'Archive values or nullable fields changed'
foreach($mode in @('fail','zero','commit-fail')){
 [FixtureSql]::Reset($mode);$saved=Save-VMDeletionToArchive -VMDetails $details -Reason Test -VMDeleted $true
 Assert (-not $saved.Success -and [FixtureSql]::Commits -eq 0 -and [FixtureSql]::Rollbacks -eq 1 -and [FixtureSql]::Executed -eq 1) "Archive failure was lost or retried: $mode"
}
$builder=New-Object Data.SqlClient.SqlConnectionStringBuilder (Get-SQLConnectionString)
Assert ($builder['Password'] -eq 'fixture;"password' -and $builder['Initial Catalog'] -eq 'FixtureArchive') 'SQL connection string did not preserve quoted password/database'
'SQL_SELECTED_ARCHIVE_COMMIT_ACK_PARAMETERIZATION_ROLLBACK_AND_CONNECTION_SETTINGS_OK'

$script:Mails=New-Object Collections.ArrayList
function Send-MailMessage{[CmdletBinding()]param($From,$To,$Subject,$Body,[switch]$BodyAsHtml,$SmtpServer,$Port,$Encoding,$Priority,[switch]$UseSsl) [void]$script:Mails.Add(@{}+$PSBoundParameters)}
function Get-ADUserEmail{param($SamAccountName,$DomainController)
 $email=switch($SamAccountName){$env:USERNAME{'actor@example.invalid'};'ownerSam'{'owner@example.invalid'};'folderSam'{'folder@example.invalid'};'sameSam'{'actor@example.invalid'};default{''}}
 return @{Success=[bool]$email;Email=$email;Message='fixture'}
}
function Get-ADGroupMemberEmails{param($GroupName,$DomainController) if($GroupName -eq 'WinTeam'){@('windows@example.invalid','dup@example.invalid')}}
$record=@{VMName='VM<&';OSType='Windows';Errors=@();VMDeleted=$true;Archived=$true;VCenter='vc.fixture.invalid';ClusterName='PROD';IsTemplate=$false}
$personal=Send-VMDeletionNotification -VMDetails @{VMName='VM<&';ServerOwner='ownerSam';FolderOwner='folderSam';VCenterName='vc.fixture.invalid'} -Reason 'Activity Ended' -OSType Windows -DeletionSuccess $true -ReportRecord $record
Assert ($personal.Success -and $personal.SentTo.Count -eq 2 -and $personal.SentTo -contains 'actor@example.invalid' -and $personal.SentTo -contains 'owner@example.invalid') 'Personal email must go only to actor and server owner'
Assert ($Mails[0].Body -match '#0C1730|#12234A' -and $Mails[0].Body -match 'VM&lt;&amp;' -and $Mails[0].UseSsl -and $Mails[0].Port -eq 2525) 'Personal report design, escaping or SMTP settings changed'
$fallback=Send-VMDeletionNotification -VMDetails @{VMName='VM-2';ServerOwner='unknown';FolderOwner='folderSam'} -OSType Linux -DeletionSuccess $true
Assert ($fallback.SentTo -contains 'folder@example.invalid' -and $fallback.SentTo.Count -eq 2) 'Folder owner fallback failed'
$linux=@{VMName='LINUX-1';OSType='Linux';Errors=@();VMDeleted=$true;Archived=$true;ClusterName='LINUXHANA';VCenter='vc.fixture.invalid'}
Send-DeletionSummaryReport -SuccessList @($record,$linux) -Reason Test
Assert ($Mails.Count -eq 3 -and $Mails[2].To -contains 'common@example.invalid' -and $Mails[2].To -contains 'windows@example.invalid' -and $Mails[2].To -contains 'linux@example.invalid') 'One batch summary must go to the configured JSON teams'
Assert ($Mails[2].To -notcontains 'actor@example.invalid' -and $Mails[2].To -notcontains 'owner@example.invalid' -and $Mails[2].To -notcontains 'old-admin@example.invalid' -and $Mails[0].To -notcontains 'common@example.invalid') 'Personal and team recipient routes were mixed'
Assert (@($Mails[2].To|Where-Object{$_ -ieq 'dup@example.invalid'}).Count -eq 1) 'Summary recipients were duplicated'
$script:DefaultSettings.EnableEmailNotification=$false
Send-DeletionSummaryReport -SuccessList @($record)
[void](Send-VMDeletionNotification -VMDetails @{VMName='disabled'} -DeletionSuccess $true)
Assert ($Mails.Count -eq 3) 'Email disabled setting was ignored'
$script:DefaultSettings.EnableEmailNotification=$true
'PERSONAL_ACTOR_OWNER_FOLDER_FALLBACK_SUMMARY_JSON_ROUTES_DEDUP_AND_DISABLED_EMAIL_OK'

$rootFolder=[pscustomobject]@{Id='folder-root';Name='Root';Parent=$null}
$childFolder=[pscustomobject]@{Id='folder-child';Name='Child';Parent=$rootFolder}
$script:OwnerMode='server'
function Get-Annotation{param($Entity,$CustomAttribute,$ErrorAction)
 $values=@{}
 if($Entity.Name -eq 'OWNER-VM'){
  if($script:OwnerMode -eq 'server'){$values.SRVOWNER='ownerSam'}
  if($script:OwnerMode -eq 'legacy'){$values.srvOWN1='ownerSam'}
 }
 if($Entity.Name -eq 'Root'){$values.FOLDEROWNER='folderSam'}
 foreach($name in $values.Keys){if(-not $CustomAttribute -or $CustomAttribute -eq $name){[pscustomobject]@{Name=$name;Value=$values[$name]}}}
}
$ownerVm=[pscustomobject]@{Name='OWNER-VM';Notes='';Folder=$childFolder}
foreach($mode in @('server','legacy','folder')){
 $script:OwnerMode=$mode;$owners=Get-DeletionOwnerMetadata -VM $ownerVm
 Assert ($owners.FolderOwner -eq 'folderSam') 'Folder annotation/inheritance was not read'
 if($mode -ne 'folder'){Assert ($owners.ServerOwner -eq 'ownerSam') 'Server annotation alias was not read'}else{Assert (-not $owners.ServerOwner) 'Missing server owner was invented'}
}
$ownerVm.Notes='srvOWN1: ownerSam';$owners=Get-DeletionOwnerMetadata -VM $ownerVm
Assert ($owners.ServerOwner -eq 'ownerSam') 'Legacy Notes owner fallback was lost'
'VC_SERVER_OWNER_LEGACY_ALIAS_FOLDER_INHERITANCE_AND_NOTES_OK'

$table=New-Object Data.DataTable
[void]$table.Columns.Add('Name]\',[string]);[void]$table.Columns.Add('CPU',[int]);[void]$table.Columns.Add('When',[datetime])
[void]$table.Rows.Add([object[]]@("O'Brien[ops]%*",16,[datetime]'2026-10-20'))
[void]$table.Rows.Add([object[]]@('plain',4,[datetime]'2025-01-01'))
$view=New-Object Data.DataView($table)
foreach($term in @("O'Brien[ops]%*",'16','2026')){$view.RowFilter=Get-WizardSqlSearchFilter -Data $table -Text $term;Assert ($view.Count -eq 1) "Literal/type search failed: $term"}
$view.RowFilter=Get-WizardSqlSearchFilter -Data $table -Text "' OR 1=1 OR '"
Assert ($view.Count -eq 0) 'Search text was treated as an expression'
$view.RowFilter=Get-WizardSqlSearchFilter -Data $table -Text ''
Assert ($view.Count -eq 2 -and $table.DefaultView.RowFilter -eq '') 'Search changed the original table default view'
'LOCAL_SQL_SEARCH_TYPES_QUOTES_BRACKETS_WILDCARDS_AND_LITERAL_INPUT_OK'
$headlessOffset=$source.IndexOf('if ($NoGui) {',$source.IndexOf('function Invoke-HeadlessCsvDeletion'))
foreach($name in @('ConvertTo-DeletionSqlCredential','New-WizardSqlConnectionString','ConvertTo-WizardSqlIdentifier')){
 $function=$ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)|Select-Object -First 1
 Assert ($function.Extent.StartOffset -lt $headlessOffset) 'Shared SQL helper is unavailable in NoGui mode'
}
'SHARED_SQL_HELPERS_AVAILABLE_BEFORE_NOGUI_ENTRY_OK'
