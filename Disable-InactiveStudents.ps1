<#
.SYNOPSIS
 Queries Aeries Student Inforamtion System and Active Directory and determines which AD accounts need to be disabled.
.DESCRIPTION
 EmployeeIDs Queried from Aeries and Active Directory student user objects are compared. If AD Object employeeID attribute
 is not present in Aeries results then the AD account is disabled,
 and if present will be used to determine if the account should remain active until the hold date expires.
.EXAMPLE
 .\Disable-InactiveStudents.ps1 -OrgUnitNoLicAD 'OU=NoLicAD,DN=Mars,DN=Colony' -OrgUnitNoLicGoogle 'OU=NoLicGoogle,DN=Mars,DN=Colony' -ADCredential $adCreds -SISServer $sisServer -SISDatabase $sisDatabase -SISCredential $sisCreds -MailCredential $malCred -MailTarget meohmy@mars.com
.INPUTS
  AD account credential object with various user object permissions
.OUTPUTS
Active Directory Object updates
GSuite account updates
Chromebook device updates
Email Messages
.NOTES
 In special cases an account can be held open until a set date.
 Use the AccountExpirationDate AD attribue to keep a student's account active.
 https://developers.google.com/admin-sdk/licensing/v1/how-tos/products
#>
[cmdletbinding()]
param (
 # [Parameter(Mandatory = $True)][Alias('DCs')][string[]]$DomainControllers,
 [Parameter(Mandatory = $True)][string]$OrgUnitNoLicAD,
 [Parameter(Mandatory = $True)][string]$OrgUnitNoLicGoogle,
 [Parameter(Mandatory = $True)][PSCredential]$ADCredential,
 [Parameter(Mandatory = $True)][string]$SISServer,
 [Parameter(Mandatory = $True)][string]$SISDatabase,
 # Aeries SQL user account with SELECT permission to STU table
 [Parameter(Mandatory = $True)][PSCredential]$SISCredential,
 [Parameter(Mandatory = $True)][string[]]$ExportMailTarget,
 [Parameter(Mandatory = $True)][PSCredential]$MailCredential,
 [Parameter(Mandatory = $True)][string[]]$MailTarget,
 [string[]]$BccAddress,
 [string[]]$CCAddress,
 [SWITCH]$Wait,
 [SWITCH]$Slow,
 [Alias('wi')][SWITCH]$WhatIf
)

# Script Functions =========================================================================

function Disable-ADObjects ([pscredential]$cred) {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.info) -F magenta
  $params = @{
   Identity   = $_.ad.ObjectGUID
   Enabled    = $false
   Confirm    = $false
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Set-ADUser @params
  $_
 }
}

function Disable-Chromebook {
 process {
  $id = $_.deviceId
  if ($crosDev.status -ne 'ACTIVE') { return }
  $msg = $MyInvocation.MyCommand.name, $_.serialNumber, "& $gam redirect stderr null update cros $id action disable"
  Write-Host ('{0},[{1}],[{2}]' -f $msg) -F DarkCyan
  if ($WhatIf) { return }
  $ErrorActionPreference = 'Continue'
  & $gam redirect stderr null update cros $id action disable *>$null
  $ErrorActionPreference = 'Stop'
 }
}

function Export-Data ([string]$filePath) {
 begin {
  'SamAccountName,LastLogonDate,WhenCreated,ExpiredGrad' | Out-File -FilePath $filePath -Force
  $global:myList = [System.Collections.Generic.List[string]]::new()
 }
 process {
  $msg = ('{0},{1},{2},{3}' -f $_.ad.SamAccountName, $_.ad.LastLogonDate, $_.ad.WhenCreated, $_.expiredGrad)
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $msg) -F DarkCyan
  $global:myList.Add($msg)
  $_
 }
 end {
  $global:myList | Out-File -FilePath $filePath -Force -Append
 }
}

function Export-Report ($ExportData) {
 $exportFileName = 'Recover_Devices-' + (Get-Date -f yyyy-MM-dd)
 $ExportBody = Get-Content -Path .\html\report_export.html -Raw
 Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, ".\reports\$exportFileName") -F DarkCyan
 if (-not(Test-Path -Path .\reports\)) { New-Item -Type Directory -Name reports -Force }
 Write-Host 'Export data to Excel file'
 Import-Module 'ImportExcel'
 $ExportData | Export-Excel -Path .\reports\$exportFileName.xlsx
 Send-ReportData -AttachmentPath .\reports\$exportFileName.xlsx -ExportHTML $ExportBody
}

function Format-Html {
 begin {
  $html = Get-Content -Path .\html\return_chromebook_message.html -Raw
 }
 process {
  $data = $_.group[0]
  $stuName = $data.FirstName + ' ' + $data.LastName
  $output = @{html = $html; stuName = $stuName ; gmail = $data.mail }
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $data.mail) -F DarkCyan
  $parentEmails = $_ | Format-ParentEmailAddresses
  $output.html = $output.html.Replace('{email}', $parentEmails)
  $output.html = $output.html.Replace('{student}', $stuName)
  $output.html = $output.html.Replace('{barcode}', $data.Barcode)
  $output
 }
}

function Format-Object {
 begin {
  Write-Host ('{0}, May take some time...' -f $MyInvocation.MyCommand.Name) -F Yellow
 }
 process {
  [pscustomobject]@{
   ad          = $null
   expiredGrad = $null
   grad        = $null
   id          = $_
   info        = $null
  }
 }
}

function Format-ParentEmailAddresses {
 process {
  # Build a string containing any parent emails
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.group[0].mail) -F DarkCyan
  foreach ($obj in $_.group) {
   if ( -not([DBNull]::Value).Equals($obj.ParentEmail) -and ($null -ne $obj.ParentEmail) -and ($obj.ParentEmail -like '*@*')) {
    if ($parentEmailList -notmatch $obj.ParentEmail) {
     $parentEmailList = $obj.ParentEmail, $parentEmailList -join '; '
    }
   }
   if ( -not([DBNull]::Value).Equals($obj.ParentPortalEmail) ) {
    if ($parentEmailList -notmatch $obj.ParentPortalEmail) {
     $parentEmailList = $obj.ParentPortalEmail, $parentEmailList -join '; '
    }
   }
  }
  $parentEmailList.TrimEnd(', ')
 }
}

function Get-ADData ($props, [pscredential]$cred) {
 $filter = "
 employeeType -eq 'student' -and
 -not(Description -like '*test*')
 "
 $params = @{
  Filter     = $filter
  Properties = $props
  Credential = $cred
 }
 $objs = Get-ADUser @params
 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.Name, @($objs).count) -F green
 $objs
}

function Get-ActiveSiS ($sqlParams) {
 $query = Get-Content -Path '.\sql\active-students.sql' -Raw
 $results = New-SqlOperation @sqlParams -Query $query | Sort-Object employeeId
 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.name, @($results).count) -F green
 $results
}

# function Get-InactiveADObj ($adData, $inactiveIDs) {
#  Write-Host ('{0}, May take some time...' -f $MyInvocation.MyCommand.Name) -F Yellow
#  $result = foreach ($id in $inactiveIDs.employeeId) {
#   $adData.Where({ $_.employeeId -eq $id })
#  }
#  Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($results).count) -F Green
#  $result
# }

function Get-InactiveIDs ($adData, $sisData) {
 Write-Host ('{0}, May take some time...' -f $MyInvocation.MyCommand.Name) -F yellow
 $results = Compare-Object -ReferenceObject $adData.EmployeeId -DifferenceObject $sisData.ID |
  Where-Object { $_.SideIndicator -eq '<=' } |
   Select-Object -ExpandProperty InputObject
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($results).count) -F Green
 $results
}

filter Get-AssignedDeviceUsers ($sqlParams) {
 begin { $query = Get-Content -Path .\sql\student_return_cb.sq.sql -Raw }
 process {
  $sqlVars = "permId=$($_.ad.EmployeeId)"
  Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.info, ($sqlVars -join ','))
  New-SqlOperation @sqlParams -Query $query -Parameters $sqlVars | Group-Object
 }
}

function Get-InactiveSeniors ($sqlParams) {
 $query = Get-Content -Path '.\sql\get-inactive-seniors.sql' -Raw
 $results = New-SqlOperation @sqlParams -Query $query | Sort-Object employeeId
 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.name, @($results).count) -F green
 $results
}

filter Get-SecondaryStudents {
 if ($null -eq $_.group) {
  $wMsg = $MyInvocation.MyCommand.name, $_.samAccountName, $_.gecos
  Write-Warning ('{0},[{1}],Grade: [{2}],Grade error.' -f $wMsg)
  return
 }
 $data = $_.group[0]
 $msg = $MyInvocation.MyCommand.name, $data.Mail, $data.Grade
 if (($data.Grade) -and ([int]$data.Grade -is [int])) {
  if ([int]$data.Grade -ge 6) {
   Write-Host ('{0},[{1}],Grade: [{2}]' -f $msg) -F green
   $_
   return
  }
  Write-Host ('{0},[{1}],Grade: [{2}],Primary student detected. Skipping.' -f $msg) -F Yellow
 }
}

# function Get-StaleAD ([int]$months) {
#  process {
#   if ($_.ad.LastLogonDate -gt $cutOff -and $_.ad.WhenCreated -gt $cutOff) { return }
#   $_
#  }
# }

function Set-ChromebookOU {
 begin {
  $targOu = '/Chromebooks/Missing'
 }
 process {
  $id = $_.deviceId
  if ($_.orgUnitPath -match $targOu) { return } # Skip is OU is correct
  $msg = $MyInvocation.MyCommand.name, $_.serialNumber, "& $gam redirect stderr null update cros $id ou $targOu"
  Write-Host ('{0},[{1}],[{2}]' -f $msg) -F magenta
  if ($WhatIf) { return }
  $ErrorActionPreference = 'Continue'
  & $gam redirect stderr null update cros $id ou $targOu *>$null
  $ErrorActionPreference = 'Stop'
 }
}

function Remove-GoogleLicense ($ou) {
 begin {
  $license = @(
   1010310008 # Google Workspace for Education Plus
  )
 }
 process {
  # if (!($_.gSuiteData)) { return $_ } # Skip if no GSuite data
  if (!$WhatIf) {
   $i = 20
   if (!$WhatIf) {
    do {
     # Wait for Google Workspace to update user orgUnit
     ($ouCheck = & $gam redirect stderr null print users query "email:$($_.HomePage)" fields 'orgUnitPath' | ConvertFrom-Csv)*>$null
     if (!$WhatIf -and !$ouCheck) { Start-Sleep 7 }
     $i--
    } until ($WhatIf -or $ouCheck.orgUnitPath -match $ou -or ($i -eq 0))
   }
  }

  $ErrorActionPreference = 'SilentlyContinue'

  ($gamUser = & $gam redirect stderr null info user $_.HomePage) *>$null
  foreach ($lic in $license) {
   if ($gamUser -match $lic) { continue }
   $msg = $MyInvocation.MyCommand.name, $_.info, $lic
   Write-Host ('{0},{1},{2}' -f $msg) -F DarkMagenta
   if (!$WhatIf) {
    try { (& $gam redirect stderr null user "$($_.ad.HomePage)" del license $lic)*>$null }
    catch {
     Write-Host ('{0},{1},{2},Error Removing License' -f $msg) -F Red
    }
   }
  }

  $ErrorActionPreference = 'Stop'
  $_
 }
}

function Remove-StaleAD ([pscredential]$cred) {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.SamAccountName) -F yellow
  $params = @{
   Identity   = $_.ad.ObjectGUID
   Recursive  = $true
   Confirm    = $false
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Remove-ADObject @params
  $_
 }
}

function Remove-StaleGSuite {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info) -F yellow
  Write-Verbose ("& $gam redirect stderr null delete user {0}" -f $_.ad.HomePage)
  if ($WhatIf) { return }
  $ErrorActionPreference = 'Continue'
  & $gam redirect stderr null delete user $_.ad.HomePage
  $ErrorActionPreference = 'Stop'
  # pause
 }
}

function Select-ADDisabled {
 process {
  $_ | Where-Object { $_.ad.Enabled -eq $false }
 }
}

function Select-ADStale ([int]$months) {
 begin {
  $logonCutOff = (Get-Date).AddMonths(-$months)
 }
 process {
  if (($_.ad.LastLogonDate -is [datetime] -and $_.ad.LastLogonDate -lt $logonCutOff) -or
   ($_.ad.LastLogonDate -isnot [datetime] -and $_.ad.WhenCreated -lt $logonCutOff)) { $_ }
 }
}

function Send-AlertEmail ([pscredential]$cred) {
 begin {
  $subject = 'Exiting Student Chromebook Return'
  $i = 0
 }
 process {
  $msg = $MyInvocation.MyCommand.name, ($MailTarget -join ','), ($CCAddress -join ','), ($BccAddress -join ',')
  Write-Host ('{0},To: [{1}],CC: [{2}],BCc: [{3}]' -f $msg) -F blue
  $mailParams = @{
   To         = $MailTarget
   From       = $cred.Username
   Subject    = $subject
   HTML       = $_.html
   SMTPServer = 'smtp.office365.com'
   Cred       = $cred
   UseSSL     = $True
   Port       = 587
   WhatIf     = $WhatIf
  }
  if ($BccAddress) { $mailParams += @{Bcc = $BccAddress } }
  if ($CCAddress) { $mailParams += @{CC = $CCAddress } }
  Write-Verbose ($_.html | Out-String)
  Send-EmailMessage @mailParams
  if (!$WhatIf) { Start-Sleep -Seconds 60 } # Avoid throttling
  $i++
 }
 end {
  Write-Host ('Emails sent: [{0}]' -f $i) -F DarkGreen
 }
}

function Send-ReportData {
 param (
  $AttachmentPath,
  $ExportHTML
 )
 Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, ($ExportMailTarget -join ',')  ) -F blue
 $mailParams = @{
  To         = $ExportMailTarget
  From       = $MailCredential.Username
  Subject    = (Get-Date -f MM/dd/yyyy) + ' - Student Device Recovery Report'
  HTML       = $ExportHTML
  Attachment = $AttachmentPath
  SMTPServer = 'smtp.office365.com'
  Cred       = $MailCredential
  UseSSL     = $True
  Port       = 587
  WhatIf     = $WhatIf
 }
 Write-Verbose ($_.html | Out-String)
 Send-EmailMessage @mailParams
}

function Set-RandomPassword ([pscredential]$cred) {
 process {
  # if ($_.randomPW -ne $true) { return $_ }
  # Password Randomizer - only for users disabled and not logged in for over 60 days.
  # Discussed with the team and this may not be needed except in specific cases.
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.name) -F magenta
  $params = @{
   Identity    = $_.ObjectGUID
   NewPassword = ConvertTo-SecureString -String (New-RandomPassword -length 16) -AsPlainText -Force
   Confirm     = $false
   Credential  = $cred
   WhatIf      = $WhatIf
  }
  Set-ADAccountPassword @params
  $_
 }
}

function Set-PropAD ($data) {
 process {
  $id = $_.id
  $_.ad = $data.Where({ $_.EmployeeId -eq $id })
  $_
 }
}

function Set-PropGrad ($data) {
 process {
  $id = $_.ad.EmployeeId
  $_.grad = $data.Where({ $_.ID -eq $id }) | ConvertTo-Csv | ConvertFrom-Csv
  $_
 }
}

function Set-PropInfo {
 process {
  $_.info = '[{0},LastLogon: {1}]' -f $_.ad.SamAccountName, ($_.ad.LastLogonDate -split ' ')[0]
  $_
 }
}

# function Set-PropSIS ($data) {
#  process {
#   $id = $_.id
#   $_.sis = $data.Where({ $_.ID -eq $id }) | ConvertTo-Csv | ConvertFrom-Csv
#   $_
#  }
# }

function Set-UserAccountControl ([pscredential]$cred) {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.name) -F magenta
  $params = @{
   Identity   = $_.ObjectGUID
   Replace    = @{ UserAccountControl = 546 }
   Confirm    = $false
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Set-ADUser @params
  $_
 }
}

function Show-MissingData {
 process {
  if (!$_.grad -and !$_.sis) {
   Write-Verbose ('{0},{1},Missing grad and sis data' -f $MyInvocation.MyCommand.Name, $_.info)
  }
  $_
 }
}

function Show-Obj {
 begin { $i = 0 }
 process {
  $i++
  Write-Verbose ($i, $MyInvocation.MyCommand.Name, $_ | Out-String)
  if ($Wait) { Read-Host 'Press Enter to continue...' }
  elseif ($Slow) { Start-Sleep 2 }
  else { Start-Sleep 0 }
 }
 end {
  Write-Verbose ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $i)
 }
}

function Skip-ActiveSis {
 process {
  $_ | Where-Object { !$_.sis }
 }
}

function Skip-Disabled ($ou) {
 process {
  if (($_.ad.Enabled -eq $false -or $_.ad.Enabled -eq 'false') -and $_.ad.DistinguishedName -match $ou) { return }
  $_
 }
}

function Skip-SaturdayResets {
 process {
  if ($null -eq $_.LastLogonDate) { return }
  #   Write-Verbose (get-date $_.LastLogonDate).dayofweek -f Green
  if ((Get-Date $_.LastLogonDate).dayofweek -eq 'Saturday') { return }
  $_
 }
}

function Skip-RecentGraduates ($months) {
 process {
  if (!$_.grad) { return $_ }
  $completionGraceCutoffDate = (Get-Date $_.grad.completionDate).AddMonths($months)
  $msg = $MyInvocation.MyCommand.Name, $_.info, $completionGraceCutoffDate
  if ((Get-Date) -lt $completionGraceCutoffDate) {
   return (Write-Host ('{0},{1},{2},Qualifying Senior Detected. Skipping' -f $msg) -f Magenta)
  }
  Write-Verbose ('{0},{1},{2},Expired Senior detected.' -f $msg)
  $_.expiredGrad = $true
  $_
 }
}

function Skip-TestAccounts {
 process {
  if ($_.ad.Description -match 'test') { return }
  $_
 }
}

function Update-AccountExpirationDate ([pscredential]$cred) {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.info) -F magenta
  $params = @{
   Identity   = $_.ad.ObjectGUID
   DateTime   = [datetime]::Today
   Confirm    = $false
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Set-ADAccountExpiration @params
  $_
 }
}

function Update-Chromebooks {
 begin {
  $crosFields = 'serialNumber,orgUnitPath,deviceId,status'
 }
 process {
  if ($null -eq $_.group) { return }
  $data = $_.group[0]
  $sn = $data.serialNumber
  $msg = $MyInvocation.MyCommand.name, $data.mail, $sn, "& $gam redirect stderr null print cros query `"id: $sn`" fields $crosFields"
  Write-Host ('{0},[{1}],[{2}],[{3}]' -f $msg) -F magenta
  $ErrorActionPreference = 'Continue'
  ($crosDev = & $gam redirect stderr null print cros query "id: $sn" fields $crosFields | ConvertFrom-Csv)*>$null
  $ErrorActionPreference = 'Stop'
  if ($crosDev) {
   $crosDev | Set-ChromebookOU
   $crosDev | Disable-Chromebook
   $_
  }
 }
}

function Update-GoogleArchive {
 process {
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.info) -F magenta
  if (!$WhatIf -and $_.ad.HomePage) {
   $ErrorActionPreference = 'Continue'
   (& $gam redirect stderr null update user $_.ad.HomePage archived on) *>$null
   $ErrorActionPreference = 'Stop'
  }
  $_
 }
}

function Update-GoogleSuspended {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.info) -F magenta
  if ($_.ad.HomePage -and -not$WhatIf) {
   $ErrorActionPreference = 'Continue'
   (& $gam update user $_.ad.HomePage suspended on) *>$null
   $ErrorActionPreference = 'Stop'
  }
  $_
 }
}

function Update-Grade ([pscredential]$cred) {
 process {
  # Set grade to 9999 to indicate inactive student.
  if ($_.ad.gecos -eq '9999') { return $_ }
  $params = @{
   Identity   = $_.ad.ObjectGUID
   Replace    = @{gecos = '9999' }
   Confirm    = $false
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Write-Verbose ('{0},[{1}],Gecos = 9999' -f $MyInvocation.MyCommand.Name, $_.info)
  Set-ADUser @params
  $_
 }
}

function Update-OrgUnitAD ($ou, [pscredential]$cred) {
 process {
  if ($_.DistinguishedName -match $ou ) { return }
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Magenta
  $params = @{
   Identity   = $_.ad.ObjectGUID
   TargetPath = $ou
   Confirm    = $false
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Move-ADObject @params
  $_
 }
}

function Update-OrgUnitGoogle ($ou) {
 process {
  # if ($_.gSuiteData.orgUnitPath -match $ou) { return $_ }
  Write-Host ('{0},{1},[{2}]' -f $MyInvocation.MyCommand.Name, $_.info, $ou) -F Magenta
  if (!$WhatIf) { (& $gam redirect stderr nullupdate user "$($_.ad.HomePage)" org "$ou")*>null }
  $_
 }
}

# ======================================= Processing ======================================
if ($WhatIf) { Show-TestRun }

Import-Module CommonScriptFunctions -Cmdlet Clear-SessionData, Connect-ADSession, Show-TestRun, New-SqlOperation, New-RandomPassword
Import-Module -Name dbatools -Cmdlet Invoke-DbaQuery, Set-DbatoolsConfig, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module ImportExcel -Cmdlet Export-Excel
Import-Module -Name Mailozaurr -Cmdlet Send-EMailMessage

Show-BlockInfo main
Clear-SessionData
$gam = 'C:\GAM7\gam.exe'

$sqlParams = @{
 Server     = $SISServer
 Database   = $SISDatabase
 Credential = $SISCredential
}

$adProps = 'AccountExpirationDate', 'Description', 'EmployeeID', 'gecos', 'HomePage',
'info', 'LastLogonDate', 'title', 'WhenCreated'
$adData = Get-ADData -props $adProps -cred $ADCredential

$activeSiS = Get-ActiveSiS $sqlParams
$inactiveIds = Get-InactiveIds -ad $adData -sis $activeSiS

$inactiveSeniors = Get-InactiveSeniors $sqlParams -Query (Get-Content .\sql\get-inactive-seniors.sql -Raw)

# Export-Report -ExportData (($aDObjs | Get-AssignedDeviceUsers $sqlParams).group)
Show-BlockInfo 'Preparing AD objects'
$inactive = $inactiveIds |
 Format-Object |
  Set-PropAD -data $adData |
   Set-PropInfo |
    Set-PropGrad -data $inactiveSeniors |
     Skip-RecentGraduates -months 2
# Show-MissingData |
# Export-Data -filePath '.\export\inactive-students.csv' |
# Show-Obj

Show-BlockInfo 'Disabling inactive student accounts'
$inactive |
 Skip-Disabled -ou $OrgUnitNoLicAD |
  Update-Grade -cred $ADCredential |
   Update-GoogleArchive |
    Update-OrgUnitAD -ou $OrgUnitNoLicAD -cred $ADCredential |
     Update-OrgUnitGoogle -ou $OrgUnitNoLicGoogle |
      Remove-GoogleLicense -ou $OrgUnitNoLicGoogle |
       Update-GoogleSuspended |
        # Set-RandomPassword -cred $ADCredential |
        Update-AccountExpirationDate -cred $ADCredential |
         Show-Obj

Show-BlockInfo 'Chrome devices'
$inactive |
 Get-AssignedDeviceUsers $sqlParams |
  Update-Chromebooks |
   Get-SecondaryStudents |
    Format-Html |
     Send-AlertEmail -cred $MailCredential |
      Show-Obj

Show-BlockInfo 'Removing stale student accounts'
$inactive |
 Select-ADDisabled |
  Select-ADStale -months 18 |
   Remove-StaleAD -cred $ADCredential |
    Remove-StaleGSuite |
     Show-Obj

Clear-SessionData
if ($WhatIf) { Show-TestRun }