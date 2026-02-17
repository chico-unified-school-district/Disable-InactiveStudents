#Requires -Version 5.0
<#
.SYNOPSIS
 Queries Aeries Student Inforamtion System and Active Directory and determines which AD accounts need to be disabled.
.DESCRIPTION
 EmployeeIDs Queried from Aeries and Active Directory student user objects are compared. If AD Object employeeID attribute
 is not present in Aeries results then the AD account is disabled,
 and if present will be used to determine if the account should remain active until the hold date expires.
.EXAMPLE
 .\Disable-InactiveStudents.ps1 -DC $dc -RootOU 'OU=StuOU,DN=Mars,DN=Colony' -ADCredential $adCreds -SISConnection $sisConn -SISCredential $sisCreds -MailCred $malCred -MailTarg meohmy@mars.com
.INPUTS
 Active Directory Domain controller, AD account credentail object with server permissions and various user object permissions
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
 [Parameter(Mandatory = $True)]
 [Alias('DCs')]
 [string[]]$DomainControllers,
 [Parameter(Mandatory = $True)]
 [string]$RootOU,
 # PSSession to Domain Controller and Use Active Directory CMDLETS
 [Parameter(Mandatory = $True)]
 [Alias('ADCred')]
 [System.Management.Automation.PSCredential]$ADCredential,
 # Aeries Server\Database combination
 [Parameter(Mandatory = $True)]
 [string]$SISServer,
 [Parameter(Mandatory = $True)]
 [string]$SISDatabase,
 # Aeries SQL user account with SELECT permission to STU table
 [Parameter(Mandatory = $True)]
 [Alias('SISCred')]
 [System.Management.Automation.PSCredential]$SISCredential,
 [Parameter(Mandatory = $True)]
 [string[]]$ExportMailTarget,
 [Parameter(Mandatory = $True)]
 [System.Management.Automation.PSCredential]$MailCredential,
 [Parameter(Mandatory = $True)]
 [string[]]$MailTarget,
 [string[]]$BccAddress,
 [string[]]$CCAddress,
 [Alias('wi')]
 [SWITCH]$WhatIf
)

# Output Colors
$info = 'Blue'
$alert = 'Yellow'
$get = 'Green'
$update = 'Magenta'

# Script Functions =========================================================================
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

function Get-ADData {
 $properties = 'AccountExpirationDate', 'EmployeeID', 'HomePage', 'info', 'title', 'gecos'
 $allStuParams = @{
  Filter     = { (homepage -like '*@*') -and (employeeID -like '*') }
  SearchBase = $RootOU
  Properties = $properties
 }

 $objs = Get-ADUser @allStuParams | Where-Object {
  $_.SamAccountName -match '^\b[a-zA-Z][a-zA-Z]\d{5,6}\b$' -and
  $_.employeeID -match '^\d{5,6}$' -and
  $_.title -notmatch 'test' -and
  $_.AccountExpirationDate -isnot [datetime] -and
  $_.Enabled -eq $True
 }
 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.Name, @($objs).count) -F $get
 $objs | Sort-Object employeeId
}

function Get-SuperStaleAD {
 $cutOff = (Get-Date).AddMonths(-18) # Ask Director of IT before changing.
 $properties = 'LastLogonDate', 'EmployeeID', 'HomePage', 'title', 'WhenCreated'
 $allStuParams = @{
  Filter     = { (homepage -like '*@*') -and (employeeID -like '*') -and (Enabled -eq 'False') }
  SearchBase = $RootOU
  Properties = $properties
 }
 $objs = Get-ADUser @allStuParams | Where-Object {
  $_.SamAccountName -match '^\b[a-zA-Z][a-zA-Z]\d{5,6}\b$' -and
  $_.employeeID -match '^\d{5,6}$' -and
  $_.title -notmatch 'test' -and
  $_.LastLogonDate -lt $cutOff -and
  $_.WhenCreated -lt $cutOff
 }
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($objs).count) -F $get
 # Start-Sleep 3 # why?
 $objs | Sort-Object employeeId
}

function Get-ActiveSiS ($sqlParams) {
 $query = Get-Content -Path '.\sql\active-students.sql' -Raw
 $results = New-SqlOperation @sqlParams -Query $query | Sort-Object employeeId
 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.name, @($results).count) -F $get
 $results
}

function Get-InactiveADObj ($adData, $inactiveIDs) {
 foreach ($id in $inactiveIDs.employeeId) {
  $adData.Where({ $_.employeeId -eq $id })
 }
}

function Get-InactiveIDs ($adData, $sisData) {
 Write-Host $MyInvocation.MyCommand.name -F $get
 Compare-Object -ReferenceObject $sisData -DifferenceObject $adData -Property employeeId |
  Where-Object { $_.SideIndicator -eq '=>' }
}

filter Get-AssignedDeviceUsers ($sqlParams) {
 begin { $query = Get-Content -Path .\sql\student_return_cb.sq.sql -Raw }
 process {
  $sqlVars = "permId=$($_.employeeId)"
  Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.name, ($sqlVars -join ','))
  New-SqlOperation @sqlParams -Query $query -Parameters $sqlVars | Group-Object
 }
}

function Get-InactiveSeniors ($sqlParams) {
 $query = Get-Content -Path '.\sql\get-inactive-seniors.sql' -Raw
 $results = New-SqlOperation @sqlParams -Query $query | Sort-Object employeeId
 Write-Host ('{0},Count: [{1}]' -f $MyInvocation.MyCommand.name, @($results).count) -F $get
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
   Write-Host ('{0},[{1}],Grade: [{2}]' -f $msg) -F $get
   $_
   return
  }
  Write-Host ('{0},[{1}],Grade: [{2}],Primary student detected. Skipping.' -f $msg) -F Yellow
 }
}

function Get-StaleAD {
 $cutOff = (Get-Date).AddMonths(-1) # Ask Director of IT before changing.
 $properties = 'LastLogonDate', 'EmployeeID', 'HomePage', 'title', 'WhenCreated'
 $allStuParams = @{
  Filter     = { (homepage -like '*@*') -and (employeeID -like '*') -and (Enabled -eq 'False') }
  SearchBase = $RootOU
  Properties = $properties
 }
 $objs = Get-ADUser @allStuParams | Where-Object {
  $_.SamAccountName -match '^\b[a-zA-Z][a-zA-Z]\d{5,6}\b$' -and
  $_.employeeID -match '^\d{5,6}$' -and
  $_.title -notmatch 'test' -and
  $_.LastLogonDate -lt $cutOff -and
  $_.WhenCreated -lt $cutOff
 }
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($objs).count) -F $get
 $objs | Sort-Object employeeId
}

function Disable-ADObjects {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.name) -F $update
  Set-ADUser -Identity $_.ObjectGUID -Enabled:$false -Confirm:$false -WhatIf:$WhatIf
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
  $msg = $MyInvocation.MyCommand.name, $data.mail, $sn, "& $gam print cros query `"id: $sn`" fields $crosFields"
  Write-Host ('{0},[{1}],[{2}],[{3}]' -f $msg) -F $update
  $ErrorActionPreference = 'Continue'
  ($crosDev = & $gam print cros query "id: $sn" fields $crosFields | ConvertFrom-Csv)*>$null
  $ErrorActionPreference = 'Stop'
  if ($crosDev) {
   $crosDev | Set-ChromebookOU
   $crosDev | Disable-Chromebook
   $_
  }
 }
}

function Set-ChromebookOU {
 begin {
  $targOu = '/Chromebooks/Missing'
 }
 process {
  $id = $_.deviceId
  if ($_.orgUnitPath -match $targOu) { return } # Skip is OU is correct
  $msg = $MyInvocation.MyCommand.name, $_.serialNumber, "& $gam update cros $id ou $targOu"
  Write-Host ('{0},[{1}],[{2}]' -f $msg) -F $update
  if ($WhatIf) { return }
  $ErrorActionPreference = 'Continue'
  & $gam update cros $id ou $targOu *>$null
  $ErrorActionPreference = 'Stop'
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

function Disable-Chromebook {
 process {
  $id = $_.deviceId
  if ($crosDev.status -ne 'ACTIVE') { return }
  $msg = $MyInvocation.MyCommand.name, $_.serialNumber, "& $gam update cros $id action disable"
  Write-Host ('{0},[{1}],[{2}]' -f $msg) -F DarkCyan
  if ($WhatIf) { return }
  $ErrorActionPreference = 'Continue'
  & $gam update cros $id action disable *>$null
  $ErrorActionPreference = 'Stop'
 }
}

function Remove-GSuiteLicense {
 process {
  #SKU: 1010310003 = Google Workspace for Education Plus - Legacy (Student)
  #SKU: 1010310008 = Google Workspace for Education Plus
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.HomePage) -F $update
  $cmd = "& $gam user {0} delete license 1010310008" -f $_.HomePage
  Write-Verbose $cmd
  if ($_.HomePage -and -not$WhatIf) {
   $ErrorActionPreference = 'Continue'
   (& $gam user $_.HomePage delete license 1010310008) *>$null
   $ErrorActionPreference = 'Stop'
  }
  $_
 }
}

function Send-AlertEmail {
 begin {
  $subject = 'Exiting Student Chromebook Return'
  $i = 0
 }
 process {
  $msg = $MyInvocation.MyCommand.name, ($MailTarget -join ','), ($CCAddress -join ','), ($BccAddress -join ',')
  Write-Host ('{0},To: [{1}],CC: [{2}],BCc: [{3}]' -f $msg) -F $info
  $mailParams = @{
   To         = $MailTarget
   From       = $MailCredential.Username
   Subject    = $subject
   HTML       = $_.html
   SMTPServer = 'smtp.office365.com'
   Cred       = $MailCredential
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
 Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, ($ExportMailTarget -join ',')  ) -F $info
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

function Set-RandomPassword {
 process {
  if ($_.randomPW -ne $true) { return $_ }
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.name) -F $update
  $randomPW = ConvertTo-SecureString -String (New-RandomPassword) -AsPlainText -Force
  Set-ADAccountPassword -Identity $_.ObjectGUID -NewPassword $randomPW -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Set-GSuiteArchiveOn {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.name) -F $update
  if ($_.HomePage -and -not$WhatIf) {
   $ErrorActionPreference = 'Continue'
   (& $gam update user $_.HomePage archived on) *>$null
   $ErrorActionPreference = 'Stop'
  }
  $_
 }
}

function Set-GSuiteSuspended {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.name) -F $update
  if ($_.HomePage -and -not$WhatIf) {
   $ErrorActionPreference = 'Continue'
   (& $gam update user $_.HomePage suspended on) *>$null
   $ErrorActionPreference = 'Stop'
  }
  $_
 }
}

function Set-UserAccountControl {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.name, $_.name) -F $update
  Set-ADUser -Identity $_.ObjectGUID -Replace @{UserAccountControl = 546 } -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Remove-StaleAD {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.SamAccountName) -F $alert
  Remove-ADObject -Identity $_.ObjectGUID -Recursive -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Remove-StaleGSuite {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.HomePage) -F $alert
  Write-Verbose ("& $gam delete user {0}" -f $_.HomePage)
  if ($WhatIf) { return }
  $ErrorActionPreference = 'Continue'
  & $gam delete user $_.HomePage
  $ErrorActionPreference = 'Stop'
  # pause
 }
}

function Show-Obj {
 begin { $i = 0 }
 process {
  $i++
  Write-Verbose ($i, $MyInvocation.MyCommand.Name, $_ | Out-String)
  # Write-Debug 'Proceed?'
 }
}

function Skip-SeniorGrads ($inactiveSeniors) {
 process {
  if ($inactiveSeniors.permId -contains $_.EmployeeID) {
   $msg = $MyInvocation.MyCommand.Name, $_.EmployeeID
   return (Write-Host ('{0},{1},Qualifying Senior Detected. Skipping.' -f $msg))
  }
  $_
 }
}


function Update-Grade {
 process {
  # Set grade to 9999 to indicate inactive student. Other processing can use this discrepancy to identify and fix if needed.
  Write-Host ('{0},[{1}] Gecos = 9999' -f $MyInvocation.MyCommand.Name, $_.SamAccountName) -F Cyan
  Set-ADUser -Identity $_.ObjectGUID -Replace @{gecos = '9999' } -Confirm:$false -WhatIf:$WhatIf
 }
}

# ======================================= Processing ======================================
if ($WhatIf) { Show-TestRun }

Import-Module CommonScriptFunctions -Cmdlet Clear-SessionData, Connect-ADSession, Show-TestRun, New-SqlOperation
Import-Module -Name dbatools -Cmdlet Invoke-DbaQuery, Set-DbatoolsConfig, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module ImportExcel -Cmdlet Export-Excel
Import-Module -Name Mailozaurr -Cmdlet Send-EMailMessage

Show-BlockInfo main
Clear-SessionData
$gam = 'C:\GAM7\gam.exe'

$cmdlets = 'Get-ADUser', 'Set-ADUser', 'Set-ADAccountPassword', 'Remove-ADobject'
Connect-ADSession -DomainControllers $DomainControllers -Cmdlets $cmdlets -Credential $ADCredential

$sqlParams = @{
 Server     = $SISServer
 Database   = $SISDatabase
 Credential = $SISCredential
}

$studentADData = Get-ADData
$activeSiS = Get-ActiveSiS $sqlParams
$inactiveIDs = Get-InactiveIDs -adData $studentADData -sisData $activeSiS

$inactiveSeniors = Get-InactiveSeniors $sqlParams -Query (Get-Content .\sql\get-inactive-seniors.sql -Raw)

$aDObjs = Get-InactiveADObj -adData $studentADData -inactiveIDs $inactiveIDs

Export-Report -ExportData (($aDObjs | Get-AssignedDeviceUsers $sqlParams).group)

# Processing inactive student accounts
$adObjs |
 Skip-SeniorGrads $inactiveSeniors |
  Update-Grade |
   # Disable-ADObjects |
   # Set-UserAccountControl |
   # Set-GsuiteSuspended |
   Remove-GsuiteLicense |
    Set-GSuiteArchiveOn |
     Get-AssignedDeviceUsers $sqlParams |
      Update-Chromebooks |
       Get-SecondaryStudents |
        Format-Html |
         Send-AlertEmail |
          Show-Obj

Write-Debug 'Process stale?'
# Password Randomizer - only for users disabled and not logged in for over 60 days.
Get-StaleAD | Skip-SaturdayResets | Set-RandomPassword | Show-Obj

Write-Debug 'Process super stale?'
# Remove old student accounts
Get-SuperStaleAD | Remove-StaleAD | Remove-StaleGsuite | Show-Obj

Clear-SessionData
if ($WhatIf) { Show-TestRun }