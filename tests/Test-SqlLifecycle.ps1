$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data
$source = [IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) 'VM_DELETION_FINAL.ps1'))
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($name in @('Get-DeletionStepDefinitions','Test-DeletionStepEnabled','Get-DeletionStepSkipReason','Get-DeletionOSType','New-DeletionSqlConnection','New-SQLArchiveTable','ConvertTo-WizardSqlIdentifier','ConvertTo-DeletionSqlCredential','New-WizardSqlConnectionString','Get-SQLConnectionString','Write-VMTrackingRecord','Save-VMDeletionToArchive','Invoke-HeadlessCsvDeletion')) {
    $definition = $ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name}, $true) | Select-Object -First 1
    if (-not $definition) { throw "Missing $name" }
    Invoke-Expression $definition.Extent.Text
}
function Assert([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
$script:Logs = New-Object Collections.ArrayList
function Write-Log { param($Message, $Level) [void]$script:Logs.Add([string]$Message) }
$script:DefaultSettings = @{SQLServer='sql.fixture.invalid';SQLDatabase='FixtureArchive';SQLTable='dbo.DeletedVMs';SQLIntegratedAuth=$true;SQLUsername='';SQLPassword='';EnableEmailNotification=$false}
$global:DefaultVIServer = $null
$factory = $ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'New-DeletionSqlConnection'}, $true) | Select-Object -First 1
Assert ($factory.Extent.Text -match 'New-Object System.Data.SqlClient.SqlConnection\(\$ConnectionString\)' -and $factory.Extent.Text -notmatch '\$script:|\.Open\(') 'Factory must create a fresh SQL connection without consulting UI state or opening it.'

Add-Type -ReferencedAssemblies System.Data -TypeDefinition @'
using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
public static class LifecycleFixture {
 public static List<Connection> Connections=new List<Connection>();
 public static Queue<string> Modes=new Queue<string>();
 public static void Reset(params string[] modes){Connections.Clear();Modes.Clear();foreach(string mode in modes)Modes.Enqueue(mode);}
 public class Connection:IDisposable {
  public string Mode, ConnectionString;
  public int Opened, Disposed;
  public bool IsOpen;
  public Transaction Transaction;
  public Command Command;
  public Connection(string connectionString){ConnectionString=connectionString;Mode=Modes.Count>0?Modes.Dequeue():"ok";Connections.Add(this);}
  public void Open(){Opened++;if(Mode=="open-fail")throw new Exception("Fixture SQL unavailable");IsOpen=true;}
  public Transaction BeginTransaction(){if(Mode=="begin-fail")throw new Exception("Fixture transaction denied");Transaction=new Transaction(this);return Transaction;}
  public Command CreateCommand(){Command=new Command(this);return Command;}
  public void Dispose(){Disposed++;IsOpen=false;}
 }
 public class Transaction:IDisposable {
  Connection connection;
  public int Commits, Rollbacks, Disposed;
  public Transaction(Connection value){connection=value;}
  public void Commit(){if(connection.Mode=="commit-fail")throw new Exception("Fixture commit failed");Commits++;}
  public void Rollback(){Rollbacks++;}
  public void Dispose(){Disposed++;}
 }
 public class Command:IDisposable {
  Connection connection;
  SqlCommand backing=new SqlCommand();
  public string CommandText{get;set;}
  public int CommandTimeout{get;set;}
  public Transaction Transaction{get;set;}
  public int Executed, Disposed;
  public SqlParameterCollection Parameters{get{return backing.Parameters;}}
  public Command(Connection value){connection=value;}
  public int ExecuteNonQuery(){Executed++;if(connection.Mode=="insert-fail")throw new Exception("Fixture insert denied");if(Parameters.Contains("@InsertedRows"))Parameters["@InsertedRows"].Value=connection.Mode=="zero"?0:1;return -1;}
  public void Dispose(){Disposed++;backing.Dispose();}
 }
}
'@
function New-DeletionSqlConnection { param($ConnectionString) return (New-Object LifecycleFixture+Connection($ConnectionString)) }
$script:ManualChecks = 0
function Test-SQLConnection { $script:ManualChecks++; throw 'Archive must not require a manual connection check.' }
function Test-SQLTableExists { $script:ManualChecks++; throw 'Archive must not require a manual table check.' }
function Get-WizardSqlTableSnapshot { $script:ManualChecks++; throw 'Archive must not read or display the SQL table first.' }
$script:WizardSqlSnapshot = $null
$script:WizardSqlView = $null
$details = @{VMName='FIXTURE-VM';ServerName='FIXTURE-VM';VCenterName='vc.fixture.invalid';CPUCores=2;MemoryGB=4}
function Assert-DisposedConnections {
    foreach ($connection in [LifecycleFixture]::Connections) {
        Assert ($connection.Opened -eq 1 -and $connection.Disposed -eq 1 -and -not $connection.IsOpen) 'Every archive attempt must open once and dispose its own connection exactly once.'
        if ($connection.Transaction) { Assert ($connection.Transaction.Disposed -eq 1) 'Archive transaction was not disposed.' }
        if ($connection.Command) { Assert ($connection.Command.Disposed -eq 1) 'Archive command was not disposed.' }
    }
}
[LifecycleFixture]::Reset('ok','ok')
foreach ($index in 1..2) {
    $script:sqlTableExists = $false
    $script:DefaultSettings.SQLServer = "sql-$index.fixture.invalid"
    $result = Save-VMDeletionToArchive -VMDetails $details -Reason 'Fixture only' -VMDeleted $true -Status Success
    Assert $result.Success 'Archive required a prior table connection.'
    Assert ([LifecycleFixture]::Connections.Count -eq $index) 'Archive reused an existing connection.'
    Assert-DisposedConnections
}
Assert (-not [object]::ReferenceEquals([LifecycleFixture]::Connections[0],[LifecycleFixture]::Connections[1])) 'Separate records share a connection instance.'
foreach ($index in 0..1) {
    $builder = New-Object Data.SqlClient.SqlConnectionStringBuilder ([LifecycleFixture]::Connections[$index].ConnectionString)
    Assert ($builder.DataSource -eq "sql-$($index+1).fixture.invalid") 'Archive did not use the current saved SQL settings.'
}
Assert ($script:ManualChecks -eq 0) 'Manual SQL actions were called by archiving.'
'EACH_ARCHIVE_HAS_OWN_CONNECTION_WITHOUT_CHECK_OR_DISPLAY_OK'
foreach ($mode in @('open-fail','begin-fail','insert-fail','zero','commit-fail')) {
    [LifecycleFixture]::Reset($mode,'ok')
    $script:sqlTableExists = $false
    $result = Save-VMDeletionToArchive -VMDetails $details -Reason 'Fixture failure' -VMDeleted $true -Status Success
    Assert (-not $result.Success) "Archive failure was ignored: $mode"
    Assert-DisposedConnections
    $first = [LifecycleFixture]::Connections[0]
    if ($first.Transaction) { Assert ($first.Transaction.Commits -eq 0 -and $first.Transaction.Rollbacks -eq 1) "Failed archive did not roll back: $mode" }
    $result = Save-VMDeletionToArchive -VMDetails $details -Reason 'Fixture recovery' -VMDeleted $true -Status Success
    Assert ($result.Success -and [LifecycleFixture]::Connections.Count -eq 2) "Next record did not independently reconnect: $mode"
    Assert-DisposedConnections
}
'CONNECTION_DISPOSED_ON_OPEN_BEGIN_INSERT_ACK_OR_COMMIT_FAILURE_AND_RECONNECT_OK'

# Exercise the actual NoGui loop with all deletion, email and infrastructure actions disabled.
$SkipVM=$true; $SkipSCCM=$true; $SkipAD=$true; $SkipDNS=$true; $SkipEmail=$true; $SkipArchive=$false
$Reason='Fixture only'; $script:RunEnvironment='Fixture'
function Get-NoGuiActionSummary { return @{Enabled='Archive';Skipped='VM,SCCM,AD,DNS,Email'} }
function Import-VMDeletionCsv { param($Path) @([pscustomobject]@{VMName='FIXTURE-ONE';Environment='Fixture'},[pscustomobject]@{VMName='FIXTURE-TWO';Environment='Fixture'}) }
function Get-VCenterServerForEnvironment { param($EnvironmentName) return @{Name='Fixture';Server='vc.fixture.invalid'} }
function Connect-VIServer { throw 'Live connection is forbidden in this fixture.' }
function Remove-VM { throw 'VM deletion is forbidden in this fixture.' }
function Remove-SCCMDevice { throw 'SCCM deletion is forbidden in this fixture.' }
function Remove-ADDevice { throw 'AD deletion is forbidden in this fixture.' }
function Remove-DNSRecords { throw 'DNS deletion is forbidden in this fixture.' }
function Send-VMDeletionNotification { throw 'Email is forbidden in this fixture.' }
[LifecycleFixture]::Reset('ok','ok'); $script:sqlTableExists=$false
$exitCode = Invoke-HeadlessCsvDeletion -Path 'fixture-only.csv'
Assert ($exitCode -eq 0 -and [LifecycleFixture]::Connections.Count -eq 2) 'NoGui did not archive both records independently.'
Assert-DisposedConnections
[LifecycleFixture]::Reset('insert-fail','ok'); $script:sqlTableExists=$false; $script:Logs.Clear()
$exitCode = Invoke-HeadlessCsvDeletion -Path 'fixture-only.csv'
Assert ($exitCode -eq 1 -and [LifecycleFixture]::Connections.Count -eq 2 -and ($script:Logs -join "`n") -match "NoGui deletion failed for 'FIXTURE-ONE': Archive:" -and ($script:Logs -join "`n") -match "NoGui deletion completed successfully for 'FIXTURE-TWO'") 'NoGui skipped SQL, hid failure, or failed to reconnect for the next record.'
Assert-DisposedConnections
$SkipArchive=$true; [LifecycleFixture]::Reset('ok')
$exitCode = Invoke-HeadlessCsvDeletion -Path 'fixture-only.csv'
Assert ($exitCode -eq 0 -and [LifecycleFixture]::Connections.Count -eq 0) 'SkipArchive opened a SQL connection.'
Assert ($script:ManualChecks -eq 0) 'NoGui depended on check/display SQL actions.'
'NOGUI_INDEPENDENT_ARCHIVE_FAILURE_REPORTING_RECOVERY_AND_SKIP_OK'

[LifecycleFixture]::Reset('ok')
$created=New-SQLArchiveTable -ConnectionString 'fixture-only' -TableName '[audit].[VM records]'
Assert ($created.Success -and [LifecycleFixture]::Connections.Count -eq 1) 'Missing-table creation action did not complete.'
$command=[LifecycleFixture]::Connections[0].Command
Assert ($command.CommandText -match 'CREATE TABLE \[audit\]\.\[VM records\]' -and $command.Parameters['@TableName'].Value -eq '[audit].[VM records]') 'Table creation identifier quoting or parameterized existence check failed.'
Assert-DisposedConnections
[LifecycleFixture]::Reset('insert-fail')
$created=New-SQLArchiveTable -ConnectionString 'fixture-only' -TableName 'dbo.VMs'
Assert (-not $created.Success) 'Table creation failure was hidden.'
Assert-DisposedConnections
[LifecycleFixture]::Reset('ok')
$created=New-SQLArchiveTable -ConnectionString 'fixture-only' -TableName 'dbo.X; DROP TABLE X'
Assert (-not $created.Success -and [LifecycleFixture]::Connections.Count -eq 0) 'Unsafe table name reached the SQL factory.'
'SQL_CREATE_TABLE_ACTION_SUCCESS_FAILURE_QUOTING_AND_CLEANUP_OK'
