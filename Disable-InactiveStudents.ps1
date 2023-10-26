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
# Script Functions =========================================================================
function Export-Report ($ExportData) {
 $exportFileName = 'Recover_Devices-' + (Get-Date -f yyyy-MM-dd)
 $ExportBody = Get-Content -Path .\html\report_export.html -Raw
 Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, ".\reports\$exportFileName") -Fore DarkCyan
 if (-not(Test-Path -Path .\reports\)) { New-Item -Type Directory -Name reports -Force -WhatIf:$WhatIf }
 if (-not$WhatIf) {
  Write-Host 'Export data to Excel file'
  'ImportExcel' | Load-Module
  $ExportData | Export-Excel -Path .\reports\$exportFileName.xlsx
 }
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
  Write-Host ('{0},{1}' -f $data.mail, $MyInvocation.MyCommand.name) -Fore DarkCyan
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
  Write-Host ('{0},{1}' -f $_.group[0].mail, $MyInvocation.MyCommand.name) -Fore DarkCyan
  foreach ($obj in $_.group) {
   if ( -not([DBNull]::Value).Equals($obj.ParentEmail) -and ($null -ne $obj.ParentEmail) -and ($obj.ParentEmail -like '*@*')) {
    if ($parentEmailList -notmatch $obj.ParentEmail) {
     $parentEmailList = $obj.ParentEmail, $parentEmailList -join ', '
    }
   }
   if ( -not([DBNull]::Value).Equals($obj.ParentPortalEmail) ) {
    if ($parentEmailList -notmatch $obj.ParentPortalEmail) {
     $parentEmailList = $obj.ParentPortalEmail, $parentEmailList -join ', '
    }
   }
  }
  $parentEmailList.TrimEnd(', ')
 }
}

function Get-ActiveAD {
 Write-Host $MyInvocation.MyCommand.name
 $properties = 'AccountExpirationDate', 'EmployeeID', 'HomePage', 'info', 'title'
 $allStuParams = @{
  Filter     = { (homepage -like "*@chicousd.net*") -and (employeeID -like "*") }
  SearchBase = $RootOU
  Properties = $properties
 }

 $objs = Get-ADUser @allStuParams | Where-Object {
  $_.samaccountname -match "^\b[a-zA-Z][a-zA-Z]\d{5,6}\b$" -and
  $_.employeeID -match "^\d{5,6}$" -and
  $_.title -notmatch 'test' -and
  $_.AccountExpirationDate -isnot [datetime] -and
  $_.Enabled -eq $True
 }
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $objs.count) -Fore Green
 $objs | Sort-Object employeeId
}

function Get-StaleAD {
 Write-Host $MyInvocation.MyCommand.name
 $cutOff = (Get-Date).AddMonths(-9)
 $properties = 'LastLogonDate', 'EmployeeID', 'HomePage', 'title', 'WhenCreated'
 $allStuParams = @{
  Filter     = { (homepage -like "*@*") -and (employeeID -like "*") -and (Enabled -eq 'False') }
  SearchBase = $RootOU
  Properties = $properties
 }
 $objs = Get-ADUser @allStuParams | Where-Object {
  $_.samaccountname -match "^\b[a-zA-Z][a-zA-Z]\d{5,6}\b$" -and
  $_.employeeID -match "^\d{5,6}$" -and
  $_.title -notmatch 'test' -and
  $_.LastLogonDate -lt $cutOff -and
  $_.WhenCreated -lt $cutOff
 }
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $objs.count) -Fore Green
 Start-Sleep 3
 $objs | Sort-Object employeeId
}

function Get-ActiveAeries {
 Write-Host $MyInvocation.MyCommand.name
 $sqlParams = @{
  Server     = $SISServer
  Database   = $SISDatabase
  Credential = $SISCredential
  # TrustServerCertificate = $true
 }
 $query = Get-Content -Path '.\sql\active-students.sql' -Raw
 Invoke-SqlCmd @sqlParams -Query $query | Sort-Object employeeId
}

function Get-InactiveADObj ($activeAD, $inactiveIDs) {
 foreach ($id in $inactiveIDs.employeeId) {
  $activeAD.Where({ $_.employeeId -eq $id })
 }
}

function Get-InactiveIDs ($activeAD, $activeAeries) {
 Write-Host $MyInvocation.MyCommand.name
 Compare-Object -ReferenceObject $activeAeries -DifferenceObject $activeAD -Property employeeId |
 Where-Object { $_.SideIndicator -eq '=>' }
}

filter Get-AssignedDeviceUsers {
 $sqlParams = @{
  Server     = $SISServer
  Database   = $SISDatabase
  Credential = $SISCredential
  # TrustServerCertificate = $true
 }
 Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -Fore DarkCyan
 $sql = (Get-Content -Path .\sql\student_return_cb.sq.sql -Raw) -f $_.employeeId
 Invoke-SqlCmd @sqlParams -Query $sql | Group-Object
}

filter Get-SecondaryStudents {
 $data = $_.group[0]
 if (($data.Grade) -and ([int]$data.Grade -is [int])) {
  if ([int]$data.Grade -ge 6) {
   Write-Host ('{0},{1},Grade: {2}' -f $data.Mail, $MyInvocation.MyCommand.name, $data.Grade)
   $_
  }
  else {
   Write-Host ('{0},{1},Grade: {2},Primary student detected. Skipping.' -f $data.Mail, $MyInvocation.MyCommand.name, $data.Grade) -Fore Yellow
  }
 }
 else {
  Write-Warning ('{0},{1},Grade: {2},Grade error.' -f $data.Mail, $MyInvocation.MyCommand.name, $data.Grade)
 }
}

function Disable-ADObjects {
 process {
  Write-Debug ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name)
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -Fore DarkCyan
  Set-ADUser -Identity $_.ObjectGUID -Enabled:$false -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Update-Chromebooks {
 begin {
  $crosFields = 'serialNumber,orgUnitPath,deviceId,status'
 }
 process {
  $data = $_.group[0]
  Write-Host ('{0},{1}' -f $data.mail, $MyInvocation.MyCommand.name) -Fore DarkCyan
  $sn = $data.serialNumber
  Write-Host ('{0},{1}' -f $sn, $MyInvocation.MyCommand.name) -Fore DarkCyan
  # ' *>$null suppresses noisy output '
  Write-Host "& $gam print cros query `"id: $sn`" fields $crosFields"
 ($crosDev = & $gam print cros query "id: $sn" fields $crosFields | ConvertFrom-CSV) | Out-Null
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
  if ($_.orgUnitPath -notmatch $targOu) {
   Write-Host ('{0},{1}' -f $_.deviceId, $MyInvocation.MyCommand.name) -Fore DarkCyan
   Write-Host "& $gam update cros $id ou /Chromebooks/Missing *>$null"
   if (-not$WhatIf) {
    & $gam update cros $id ou $targOu *>$null
   }
  }
  else { Write-Verbose "$id,Skipping. OrgUnitPath already $targOu" }
 }
}

function Disable-Chromebook {
 process {
  $id = $_.deviceId
  if ($crosDev.status -eq "ACTIVE") {
   Write-Host ('{0},{1}' -f $_.deviceId, $MyInvocation.MyCommand.name) -Fore DarkCyan
   Write-Host "& $gam update cros $id action disable *>$null"
   if (-not$WhatIf) {
    & $gam update cros $id action disable *>$null
   }
  }
  else { Write-Verbose "$id,Skipping. Status already `'Disabled`'" }
 }
}

function Remove-GsuiteLicense {
 process {
  #SKU: 1010310003 = Google Workspace for Education Plus - Legacy (Student)
  $cmd = "& $gam user {0} delete license 1010310003" -f $_.HomePage
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $cmd) -Fore DarkCyan
  if ($_.HomePage -and -not$WhatIf) { (& $gam user $_.HomePage delete license 1010310003) *>$null }
  $_
 }
}

function Send-AlertEmail {
 begin {
  $subject = 'Exiting Student Chromebook Return'
  $i = 0
 }
 process {
  Write-Host ('To: {0},CC: {1},BCc: {2}, {3}' -f ($MailTarget -join ','), ($CCAddress -join ','), ($BccAddress -join ','), $MyInvocation.MyCommand.name)
  Write-Debug ('{0},{1}' -f ($_.gmail -join ','), $MyInvocation.MyCommand.name)
  # Write-Debug ( $mailParams | Out-String )
  $mailParams = @{
   To         = $MailTarget
   From       = $MailCredential.Username
   Subject    = $subject
   bodyAsHTML = $true
   Body       = $_.html
   SMTPServer = 'smtp.office365.com'
   Cred       = $MailCredential
   UseSSL     = $True
   Port       = 587
  }
  if ($BccAddress) { $mailParams += @{Bcc = $BccAddress } }
  if ($CCAddress) { $mailParams += @{CC = $CCAddress } }
  Write-Verbose ($_.html | Out-String)
  if (-not$WhatIf) { Send-MailMessage @mailParams }
  $i++
 }
 end {
  Write-Host ('Emails sent: {0}' -f $i) -Fore DarkGreen
 }
}

function Send-ReportData {
 param (
  $AttachmentPath,
  $ExportHTML
 )
 Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, ($ExportMailTarget -join ',')  )
 $mailParams = @{
  To         = $ExportMailTarget
  From       = $MailCredential.Username
  Subject    = (Get-Date -f MM/dd/yyyy) + ' - Student Device Recovery Report'
  bodyAsHTML = $true
  Body       = $ExportHTML
  Attachment = $AttachmentPath
  SMTPServer = 'smtp.office365.com'
  Cred       = $MailCredential
  UseSSL     = $True
  Port       = 587
 }
 Write-Verbose ($_.html | Out-String)
 if (-not$WhatIf) { Send-MailMessage @mailParams }
}

function Set-RandomPassword {
 Process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -Fore DarkCyan
  $randomPW = ConvertTo-SecureString -String (New-RandomPassword) -AsPlainText -Force
  Set-ADAccountPassword -Identity $_.ObjectGUID -NewPassword $randomPW -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Set-GsuiteSuspended {
 process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -Fore DarkCyan
  if ($_.HomePage -and -not$WhatIf) { (& $gam update user $_.HomePage suspended on) *>$null }
  $_
 }
}

function Set-UserAccountControl {
 process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -Fore DarkCyan
  # Set uac to 514 (0x0202) to notify Bradford to stop access to network
  Set-ADUser -Identity $_.ObjectGUID -Replace @{UserAccountControl = 0x0202 } -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Remove-StaleAD {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.samaccountname) -Fore Yellow
  Remove-ADObject -Identity $_.ObjectGUID -Recursive -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Remove-StaleGSuite {
 process {
  Write-Host ('{0},[{1}]' -f $MyInvocation.MyCommand.Name, $_.HomePage) -Fore Cyan
  Write-Verbose ("& $gam delete user {0}" -f $_.HomePage)
  if ($WhatIf) { return }
  & $gam delete user $_.HomePage
  # pause
 }
}

# =========================================================================================
# Imported Functions
. .\lib\Clear-SessionData.ps1
. .\lib\Load-Module.ps1
. .\lib\New-ADSession.ps1
. .\lib\New-RandomPassword.ps1
. .\lib\Select-DomainController.ps1
. .\lib\Show-TestRun.ps1

# ======================================= Processing ======================================
Show-TestRun
Clear-SessionData

'SQLServer' | Load-Module

$gam = '.\bin\gam.exe'

$dc = Select-DomainController $DomainControllers
$cmdlets = 'Get-ADUser', 'Set-ADUser', 'Set-ADAccountPassword', 'Remove-ADobject'
New-ADSession -dc $dc -cmdlets $cmdlets -cred $ADCredential

$activeAD = Get-ActiveAD
$activeAeries = Get-ActiveAeries
$inactiveIDs = Get-InactiveIDs -activeAD $activeAD -activeAeries $activeAeries

$aDObjs = Get-InactiveADObj -activeAD $activeAD -inactiveIDs $inactiveIDs

Export-Report -ExportData (($aDObjs | Get-AssignedDeviceUsers).group)

# Disable inactive student accounts
$adObjs |
Disable-ADObjects |
Set-UserAccountControl |
Set-RandomPassword |
Set-GsuiteSuspended |
Remove-GsuiteLicense |
Get-AssignedDeviceUsers |
Update-Chromebooks |
Get-SecondaryStudents |
Format-Html |
Send-AlertEmail

# Remove old student accounts
Get-StaleAD | Remove-StaleAD | Remove-StaleGsuite

# for testing
# $adObjs | Get-AssignedDeviceUsers | Get-SecondaryStudents | Format-Html | Send-AlertEmail

Clear-SessionData
Show-TestRun