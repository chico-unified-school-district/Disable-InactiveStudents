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
 [Parameter(Mandatory = $True)][string]$OrgUnitGoogleCrOS,
 [Parameter(Mandatory = $True)][PSCredential]$ADCredential,
 [Parameter(Mandatory = $True)][string]$SISServer,
 [Parameter(Mandatory = $True)][string]$SISDatabase,
 # Aeries SQL user account with SELECT permission to STU table
 [Parameter(Mandatory = $True)][PSCredential]$SISCredential,
 [Parameter(Mandatory = $True)][string[]]$ExportMailTarget,
 [Parameter(Mandatory = $True)][PSCredential]$MailCredential,
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
  $id = $_.googleCrOS.deviceId
  if ($_.googleCrOS.status -ne 'ACTIVE') { return $_ }
  $msg = $MyInvocation.MyCommand.name, $_.info, "& $gam redirect stderr null update cros $id action disable"
  Write-Host ('{0},[{1}],[{2}]' -f $msg) -F DarkCyan
  if ($WhatIf) {
   $ErrorActionPreference = 'Continue'
   (& $gam redirect stderr null update cros $id action disable)*>$null
   $ErrorActionPreference = 'Stop'
  }
  $_
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
   google      = $null
   googleCroS  = $null
   grad        = $null
   id          = $_
   info        = $null
   sisCrOS     = $null
  }
 }
}

function Format-ParentEmailAddresses {
 process {
  # Build a string containing any parent emails
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.group[0].mail) -F DarkCyan
  foreach ($obj in $_.group) {
   if ( -not([DBNull]::Value).Equals($obj.ParentEmail) -and ($null -ne $obj.ParentEmail) -and ($obj.ParentEmail -like '*@*')) {
    if ($parentEmailList -notmatch [regex]::Escape($obj.ParentEmail)) {
     $parentEmailList = $obj.ParentEmail, $parentEmailList -join '; '
    }
   }
   if ( -not([DBNull]::Value).Equals($obj.ParentPortalEmail) ) {
    if ($parentEmailList -notmatch [regex]::Escape($obj.ParentPortalEmail)) {
     $parentEmailList = $obj.ParentPortalEmail, $parentEmailList -join '; '
    }
   }
  }
  $parentEmailList.TrimEnd(', ')
 }
}

function Get-ADData ($props, $months, [pscredential]$cred) {
 $cutOff = (Get-Date).AddMonths(-$months)
 $filter = "
 employeeType -eq 'student' -and
 -not(Description -like '*test*')
 -and (
  (LastLogonDate -lt '$cutOff' -and Enabled -eq 'false') -or
  (LastLogonDate -gt '$cutOff' -and Enabled -eq 'true')
 )
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

function Get-GoogleCrOSData {
 Get-ChildItem -Path '.\data\*' -Filter *cros* -Exclude "cros-$(Get-Date -Format 'yyyy-MM-dd').csv", '.gitkeep' |
  Remove-Item -Force -Confirm:$false

 if (Test-Path -Path ".\data\cros-$(Get-Date -Format 'yyyy-MM-dd').csv") {
  Write-Host ('{0},Using cached Google CrOS data.' -f $MyInvocation.MyCommand.Name) -F Green
  $results = Import-Csv -Path ".\data\cros-$(Get-Date -Format 'yyyy-MM-dd').csv"
 }
 else {
  Write-Host ('{0},May take some time...' -f $MyInvocation.MyCommand.Name) -F Yellow
  $fields = 'serialNumber,orgUnitPath,deviceId,status'
  ($results = gam redirect stderr null print cros query 'status:provisioned' fields $fields | ConvertFrom-Csv)*>$null
  # Cache the Google data for future use
  $results | Export-Csv -Path ".\data\cros-$(Get-Date -Format 'yyyy-MM-dd').csv" -NoTypeInformation
 }

 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.Name, @($results).count) -F green
 $results
}

function Get-GoogleUserData {
 Get-ChildItem -Path '.\data\*' -Filter *google* -Exclude "google-$(Get-Date -Format 'yyyy-MM-dd').csv", '.gitkeep' |
  Remove-Item -Force -Confirm:$false

 if (Test-Path -Path ".\data\google-$(Get-Date -Format 'yyyy-MM-dd').csv") {
  Write-Host ('{0},Using cached Google user data.' -f $MyInvocation.MyCommand.Name) -F Green
  $results = Import-Csv -Path ".\data\google-$(Get-Date -Format 'yyyy-MM-dd').csv"
 }
 else {
  Write-Host ('{0},May take some time...' -f $MyInvocation.MyCommand.Name) -F Yellow
  $fields = 'archived,suspended,orgUnitPath'
  ($results = gam redirect stderr null print users query 'orgTitle=student' fields $fields | ConvertFrom-Csv)*>$null
  # Cache the Google data for future use
  $results | Export-Csv -Path ".\data\google-$(Get-Date -Format 'yyyy-MM-dd').csv" -NoTypeInformation
 }

 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.Name, @($results).count) -F green
 $results
}

function Get-InactiveIDs ($adData, $sisData) {
 Write-Host ('{0},May take some time...' -f $MyInvocation.MyCommand.Name) -F yellow
 $results = Compare-Object -ReferenceObject $adData.EmployeeId -DifferenceObject $sisData.ID |
  Where-Object { $_.SideIndicator -eq '<=' } |
   Select-Object -ExpandProperty InputObject
 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.Name, @($results).count) -F Green
 $results
}

filter Get-AssignedDeviceUser ($sqlParams) {
 begin { $query = Get-Content -Path .\sql\student_return_cb.sq.sql -Raw }
 process {
  $results = New-SqlOperation @sqlParams -Query $query -Parameters $sqlVars | ConvertTo-Csv | ConvertFrom-Csv
  Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.Name, @($results).count) -F green
  $results
 }
}

function Get-InactiveSeniors ($sqlParams) {
 $query = Get-Content -Path '.\sql\get-inactive-seniors.sql' -Raw
 $results = New-SqlOperation @sqlParams -Query $query | Sort-Object employeeId
 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.name, @($results).count) -F green
 $results
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
   do {
    # Wait for Google Workspace to update user orgUnit
    ($ouCheck = & $gam redirect stderr null print users query "email:$($_.HomePage)" fields 'orgUnitPath' | ConvertFrom-Csv)*>$null
    if (!$ouCheck) { Start-Sleep 7 }
    $i--
   } until ($ouCheck.orgUnitPath -like [regex]::Escape($ou) -or ($i -eq 0))
  }

  $ErrorActionPreference = 'SilentlyContinue'

  ($gamUser = & $gam redirect stderr null info user $_.HomePage) *>$null
  foreach ($lic in $license) {
   if ($gamUser -match [regex]::Escape($lic)) { continue }
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
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info) -F Magenta
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
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.info) -F Magenta
  Write-Verbose ("& $gam redirect stderr null delete user {0}" -f $_.ad.HomePage)
  if ($WhatIf) { return }
  $ErrorActionPreference = 'Continue'
  (& $gam redirect stderr null delete user $_.ad.HomePage)*>$null
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

filter Select-Secondary {
 process {
  if ($_.ad.gecos -isnot [int]) { return $_ }
  $_ | Where-Object { [int]$_.ad.gecos -ge 6 }
 }
}

function Send-MissingCrOSReport ($to, $from, $bcc) {
 process {
  $_.data | Export-Csv -Path '.\reports\missing_cros_report.csv' -NoTypeInformation -Force
  $params = @{
   To         = $to
   From       = '<{0}>' -f $from.Username
   Subject    = $_.Subject
   Html       = $_.Body
   Attachment = '.\reports\missing_cros_report.csv'
   SMTPServer = 'smtp.office365.com'
   Cred       = $from
   UseSSL     = $True
   Port       = 587
   WhatIf     = $WhatIf
   Suppress   = $True
  }
  if ($bcc) { $params += @{Bcc = $bcc } }
  Send-EmailMessage @params
  Write-Verbose ($params | Out-String)
  $msg = $MyInvocation.MyCommand.name, ($params.To -join ','), $params.Subject
  Write-Host ('{0},Recipient: [{1}],Subject: [{2}]' -f $msg) -F Green
  $_
 }
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

function Set-PropEmailParams {
 begin {
  $tableData = @()
  $mailObj = [PSCustomObject]@{
   Subject = 'Exiting Student Chromebook Return - {0}' -f (Get-Date -f 'dddd, MMMM dd, yyyy')
   Body    = $null
   data    = $null
  }
 }
 process { $tableData += $_.sisCrOS }
 end {
  if (@($tableData).count -lt 1) { return }
  $mailObj.data = $tableData
  $css = Get-Content -Path '.\html\style.css' -Raw
  $head = '<style TYPE="TEXT/CSS">' + $css + '</style>'
  $message = 'Hello,<br><br>Attached is the missing Chromebooks report.'
  $sig = Get-Content -Path '.\html\emailSig.html' -Raw
  $params = @{
   Head        = $head
   PreContent  = $message
   Property    = 'test'
   PostContent = $sig
  }
  $mailObj.body = @{'See attached' = $null } | ConvertTo-Html @params | Out-String
  Write-Verbose ($mailObj | Out-String )
  Write-Verbose ( $mailObj.Body )
  $mailObj
 }
}

function Set-PropGoogleCroS ($data) {
 process {
  $sn = $_.sisCrOS.SerialNumber
  $_.googleCrOS = $data.Where({ $_.serialNumber -eq $sn })
  if (!$_.googleCrOS) { return }
  $_
 }
}

function Set-PropGoogleData ($data) {
 process {
  $mail = $_.ad.HomePage
  $_.google = $data.Where({ $_.primaryEmail -eq $mail })
  # if (!$_.google) { Write-Verbose ('{0},No Google data found for {1}' -f $MyInvocation.MyCommand.Name, $mail); Read-Host 'continue' }
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

function Set-PropSisCrOS ($data) {
 process {
  $id = $_.id
  $_.sisCrOS = $data.Where({ $_.PermID -eq $id })
  if (!$_.sisCrOS) { return }
  $_
 }
}

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

function Show-Obj ($data) {
 begin {
  $i = 0
  $total = @($data).count
 }
 process {
  $i++
  if ($data) {
   Write-Verbose (('{0},{1}/{2}' -f $MyInvocation.MyCommand.Name, $i, $total), $_ | Out-String)
  }
  else {
   Write-Verbose ("$($MyInvocation.MyCommand.Name): $i", $_ | Out-String)
  }
  if ($Wait) { Read-Host 'Press Enter to continue...' }
  elseif ($Slow) { Start-Sleep 2 }
  else { Start-Sleep 0 }
 }
 end {
  Write-Verbose ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $i)
 }
}

function Skip-Disabled ($ou) {
 process {
  if (($_.ad.Enabled -eq $false -or $_.ad.Enabled -eq 'false') -and $_.ad.DistinguishedName -match [regex]::Escape($ou)) { return }
  $_
 }
}

function Skip-RecentGraduate ($months) {
 process {
  if (!$_.grad) { return $_ }
  $completionGraceCutoffDate = (Get-Date $_.grad.completionDate).AddMonths($months)
  $msg = $MyInvocation.MyCommand.Name, $_.info, $completionGraceCutoffDate
  if ((Get-Date) -lt $completionGraceCutoffDate) {
   return (Write-Host ('{0},{1},{2},Qualifying Senior Detected. Skipping' -f $msg) -f Magenta)
  }
  # Write-Verbose ('{0},{1},{2},Expired Senior detected.' -f $msg)
  $_.expiredGrad = $true
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

function Update-AccountExpirationDate ([pscredential]$cred) {
 process {
  if ($_.ad.AccountExpirationDate -is [datetime]) { return $_ }
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

function Update-GoogleArchive {
 process {
  if ($_.google.archived) { return $_ }
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
  if ($_.google.suspended) { return $_ }
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
  if ($_.DistinguishedName -match [regex]::Escape($ou)) { return }
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
  if ($_.google.orgUnitPath -match [regex]::Escape($ou)) { return $_ }
  Write-Host ('{0},{1},[{2}]' -f $MyInvocation.MyCommand.Name, $_.info, $ou) -F Magenta
  if (!$WhatIf) { (& $gam redirect stderr null update user "$($_.ad.HomePage)" org "$ou")*>null }
  $_
 }
}

function Update-OrgUnitGoogleCrOS ($ou) {
 process {
  if ($_.googleCrOS.orgUnitPath -match [regex]::Escape($ou)) { return }
  $id = $_.googleCrOS.deviceId
  $msg = $MyInvocation.MyCommand.name, $_.info, "& $gam redirect stderr null update cros $id ou $ou"
  Write-Host ('{0},[{1}],[{2}]' -f $msg) -F magenta
  if (!$WhatIf) {
   $ErrorActionPreference = 'Continue'
   (& $gam redirect stderr null update cros $id ou $ou)*>$null
   $ErrorActionPreference = 'Stop'
  }
  $_
 }
}


# ======================================= Processing ======================================
Import-Module CommonScriptFunctions -Cmdlet Clear-SessionData, Connect-ADSession, Show-TestRun, New-SqlOperation, New-RandomPassword
Import-Module -Name dbatools -Cmdlet Invoke-DbaQuery, Set-DbatoolsConfig, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module -Name Mailozaurr -Cmdlet Send-EMailMessage

Show-BlockInfo ('Process Start: ' + (Get-Date))
if ($WhatIf) { Show-TestRun }

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
$adData = Get-ADData -props $adProps -months 18 -cred $ADCredential

$activeSiS = Get-ActiveSiS -sqlParams $sqlParams
$inactiveIds = Get-InactiveIds -ad $adData -sis $activeSiS
$inactiveSeniors = Get-InactiveSeniors -sqlParams $sqlParams
$assignedDeviceUser = Get-AssignedDeviceUser -sqlParams $sqlParams

$googleData = Get-GoogleUserData
$googleCrOSData = Get-GoogleCrOSData

Show-BlockInfo 'Preparing objects'
$inactive = $inactiveIds |
 Format-Object |
  Set-PropAD -data $adData |
   Set-PropInfo |
    Set-PropGoogleData -data $googleData |
     Set-PropGrad -data $inactiveSeniors |
      Skip-RecentGraduate -months 2

Show-BlockInfo 'Disabling inactive student accounts'
$inactive |
 Skip-Disabled -ou $OrgUnitNoLicAD |
  Update-Grade -cred $ADCredential |
   Update-GoogleArchive |
    Update-OrgUnitAD -ou $OrgUnitNoLicAD -cred $ADCredential |
     Update-OrgUnitGoogle -ou $OrgUnitNoLicGoogle |
      Remove-GoogleLicense -ou $OrgUnitNoLicGoogle |
       Update-GoogleSuspended |
        Update-AccountExpirationDate -cred $ADCredential |
         Show-Obj

Show-BlockInfo 'Update CrOS devices'
$inactiveStudentsWithCrOSDevices = $inactive |
 Select-Secondary |
  Set-PropSisCrOS -data $assignedDeviceUser |
   Set-PropGoogleCrOS -data $googleCrOSData |
    Update-OrgUnitGoogleCrOS -ou $OrgUnitGoogleCrOS |
     Disable-Chromebook

Show-BlockInfo 'Email CrOS report'
$inactiveStudentsWithCrOSDevices |
 Set-PropEmailParams |
  Send-MissingCrOSReport -to $ExportMailTarget -from $MailCredential -bcc $BccAddress |
   Show-Obj

Show-BlockInfo 'Removing stale student accounts'
$inactive |
 Select-ADDisabled |
  Select-ADStale -months 18 |
   Remove-StaleAD -cred $ADCredential |
    Remove-StaleGSuite |
     Show-Obj -data $inactiveIds

Clear-SessionData
if ($WhatIf) { Show-TestRun }
Show-BlockInfo ('Process End: ' + (Get-Date))
