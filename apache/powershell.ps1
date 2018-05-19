# https://powershell.org/forums/topic/export-ldapgc-query-to-csv/

## Separate CSV
## --------------------------------------------- ###

$users = Import-Csv ./ldap.csv
foreach ($user in $users ) {
  $roles = @()
  if ($user.Role -match '^(\{)?CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(,\ CN=(((?!\,).)*),CN=groups,CN=(((?!\,).)*),O=ecm)?(\})?$') {
    $matches
    if ($matches[2]) { $roles += $matches[2] }
    if ($matches[7]) { $roles += $matches[7] }
    if ($matches[12]) { $roles += $matches[12] }
    if ($matches[17]) { $roles += $matches[17] }
    if ($matches[22]) { $roles += $matches[22] }
    if ($matches[27]) { $roles += $matches[27] }
    if ($matches[32]) { $roles += $matches[32] }
    if ($matches[37]) { $roles += $matches[37] }
    if ($matches[42]) { $roles += $matches[42] }
    if ($matches[47]) { $roles += $matches[47] }
  }
  $roles -Join ", "
  "---"
}

## Logger
## --------------------------------------------- ###

function Global:Get-Logger{
  Param(
    [CmdletBinding()]
    [Parameter()]
    [String]$Delimiter = " ",
    [Parameter()]
    [String]$Logfile,
    [Parameter()]
    [String]$Encoding = "Default",
    [Parameter()]
    [Switch]$NoDisplay
  )
  if (!(Test-Path -LiteralPath (Split-Path $Logfile -parent) -PathType container)) {
    New-Item $Logfile -type file -Force
  }
  $logger = @{}
  $logger.Set_Item('info', (Put-Log -Delimiter $Delimiter -Logfile $logfile -Encoding $Encoding -NoDisplay $NoDisplay -Info))
  $logger.Set_Item('warn', (Put-Log -Delimiter $Delimiter -Logfile $logfile -Encoding $Encoding -NoDisplay $NoDisplay -Warn))
  $logger.Set_Item('error', (Put-Log -Delimiter $Delimiter -Logfile $logfile -Encoding $Encoding -NoDisplay $NoDisplay -Err))
  return $logger
}
function Global:Put-Log {
  Param(
    [CmdletBinding()]
    [Parameter()]
    [String]$Delimiter = " ",
    [Parameter()]
    [String]$Logfile,
    [Parameter()]
    [String]$Encoding,
    [Parameter()]
    [bool]$NoDisplay,
    [Parameter()]
    [Switch]$Info,
    [Parameter()]
    [Switch]$Warn,
    [Parameter()]
    [Switch]$Err
  )
  return {
    param([String]$msg = "")

    # Initialize variables
    $logparam = @("White", "INFO")
    if ($Warn)  { $logparam = @("Yellow", "WARN") }
    if ($Err) { $logparam = @("Red", "ERROR") }
    $txt = "[$(Get-Date -Format "yyyy/MM/dd HH:mm:ss")]${Delimiter}{0}${Delimiter}{1}" -f $logparam[1], $msg

    # Output Display
    if(!$NoDisplay) {
      Write-Host -ForegroundColor $logparam[0] $txt
    }
    # Output logfile
    if($Logfile) {
      Write-Output $txt | Out-File -FilePath $Logfile -Append -Encoding $Encoding
    }
  }.GetNewClosure()
}

## Global Initialize
## --------------------------------------------- ###

## Logger

$logger = Get-Logger
$logger.info.Invoke('PowerShell Started.')
$logfile_path   = 'C:\\tmp\Datacap\data\logs\'
$logfile_name   = 'userslist_asof_' + $Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log"
$logfile        = Join-Path $logfile_path $logfile_name

if( -not (Test-Path $logfile_path) ) { New-Item $logfile_path -Type Directory }

## Audit Logs

$auditlog_path  = 'C:\\tmp\Datacap\data\audit_logs\'
$auditlog_name  = $auditlog_path + 'userslist_asof_' + $Now.ToString("yyyy-MM-dd")
$auditlog       = Join-Path $auditlog_path ($auditlog_name + ".log")
$auditlog_csv   = Join-Path $auditlog_path ($auditlog_name + ".csv")

if( -not (Test-Path $auditlog_path) ) { New-Item $auditlog_path -Type Directory }


## CRESCO Users List (Export list from LDAP to CSV)
## --------------------------------------------- ###

$root = [ADSI]"LDAP://O=ecm"
$search = [ADSISearcher]$root
$search.Filter = '(&(objectClass=user))'
$users = $search.FindAll()
$output = @()
foreach ($user in $users) {
  $user = $user.getdirectoryentry()
  $props = @{'name'=$user.name.value;'distinguishedname'=$user.distinguishedname.value}
  $obj = New-Object -Type userObject -Prop $props
  $result = $result + $obj
}
$result | Export-Csv C:\export.csv -append

Export-Csv $logfile -Encoding Default
Write-Output 'CRESCO Users List | Successfully exported to csv file.' | Out-File -FilePath $logfile -Encoding Default -append


## CRESCO Users List (Create Audit log)
## --------------------------------------------- ###

$cre_users = Import-Csv .\cre_usrs_ldap.csv
$audit_list = $cre_users | Where-Object { $_.name -eq "yamada" } | Where-Object { $_.name -like "営業*" } | Where-Object { $_.name -like "営業[12]課" }
$audit_list

Write-Output 'CRESCO Users List | Successfully exported to audit log.' | Out-File -FilePath $logfile -Encoding Default -append


## Datacap's temporarily PDF directory (Daily House Keeping)
## --------------------------------------------- ###



## Datacap's temporarily PDF directory (Grant Permissions)
## --------------------------------------------- ###

$audit_list | ForEach-Object { $_.name }
$cre_users = $cre_users | Where-Object { $_.name -eq "yamada" } | Where-Object { $_.name -like "営業*" } | Where-Object { $_.name -like "営業[12]課" }

Write-Output 'CRESCO Users List | Successfully exported to audit log.' | Out-File -FilePath $logfile -Encoding Default -append
