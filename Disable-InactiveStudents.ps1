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
#>
[cmdletbinding()]
param (
 [Parameter(Mandatory = $True)]
 [Alias('DC', 'Server')]
 [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 5 })]
 [string]$DomainController,
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
 [System.Management.Automation.PSCredential]$MailCredential,
 [Parameter(Mandatory = $True)]
 [string[]]$MailTarget,
 [string[]]$BccAddress,
 [string[]]$CCAddress,
 [Alias('wi')]
 [SWITCH]$WhatIf
)

# Script Functions =========================================================================
function Format-Html {
 begin {
  $html = Get-Content -Path .\html\return_chromebook_message.html -Raw
 }
 process {
  $data = $_.group[0]
  $stuName = $data.FirstName + ' ' + $data.LastName
  $output = @{html = $html; stuName = $stuName ; gmail = $data.mail }
  Write-Host ('{0},{1}' -f $data.mail, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
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
  Write-Host ('{0},{1}' -f $_.group[0].mail, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
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

 Get-ADUser @allStuParams | Where-Object {
  $_.samaccountname -match "^\b[a-zA-Z][a-zA-Z]\d{5,6}\b$" -and
  $_.employeeID -match "^\d{5,6}$" -and
  $_.title -notmatch 'test' -and
  $_.AccountExpirationDate -isnot [datetime] -and
  $_.Enabled -eq $True
 } | Sort-Object employeeId
}

function Get-ActiveAeries {
 Write-Host $MyInvocation.MyCommand.name
 $sqlParams = @{
  Server     = $SISServer
  Database   = $SISDatabase
  Credential = $SISCredential
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

filter Get-AssignedChromeBookUsers {
 $sqlParams = @{
  Server     = $SISServer
  Database   = $SISDatabase
  Credential = $SISCredential
 }
 Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
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
   Write-Host ('{0},{1},Grade: {2},Primary student detected. Skipping.' -f $data.Mail, $MyInvocation.MyCommand.name, $data.Grade) -ForegroundColor Yellow
  }
 }
 else {
  Write-Warning ('{0},{1},Grade: {2},Grade error.' -f $data.Mail, $MyInvocation.MyCommand.name, $data.Grade)
 }
}

function Disable-ADObjects {
 process {
  Write-Debug ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name)
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
  Set-ADUser -Identity $_.ObjectGUID -Enabled:$false -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Update-Chromebooks {
 begin {
  $gamExe = '.\lib\gam-64\gam.exe'
  $crosFields = 'serialNumber,orgUnitPath,deviceId,status'
 }
 process {
  $data = $_.group[0]
  Write-Host ('{0},{1}' -f $data.mail, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
  $sn = $data.serialNumber
  Write-Host ('{0},{1}' -f $sn, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
  # ' *>$null suppresses noisy output '
 ($crosDev = & $gamExe print cros query "id: $sn" fields $crosFields | ConvertFrom-CSV) *>$null
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
   Write-Host ('{0},{1}' -f $_.deviceId, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
   Write-Host "& $gamExe update cros $id ou /Chromebooks/Missing *>$null"
   if (-not$WhatIf) {
    & $gamExe update cros $id ou $targOu *>$null
   }
  }
  else { Write-Verbose "$id,Skipping. OrgUnitPath already $targOu" }
 }
}

function Disable-Chromebook {
 process {
  $id = $_.deviceId
  if ($crosDev.status -eq "ACTIVE") {
   Write-Host ('{0},{1}' -f $_.deviceId, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
   Write-Host "& $gamExe update cros $id action disable *>$null"
   if (-not$WhatIf) {
    & $gamExe update cros $id action disable *>$null
   }
  }
  else { Write-Verbose "$id,Skipping. Status already `'Disabled`'" }
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
   From       = 'Chico Unified Information Services'
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
  Write-Host ('Emails sent: {0}' -f $i) -ForegroundColor DarkGreen
 }
}

function Set-RandomPassword {
 Process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
  $randomPW = ConvertTo-SecureString -String (New-RandomPassword) -AsPlainText -Force
  Set-ADAccountPassword -Identity $_.ObjectGUID -NewPassword $randomPW -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Set-GsuiteSuspended {
 begin { $gamExe = '.\lib\gam-64\gam.exe' }
 process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
  if ($_.HomePage -and -not$WhatIf) { (& $gamExe update user $_.HomePage suspended on) *>$null }
  $_
 }
}

function Set-UserAccountControl {
 process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name) -ForegroundColor DarkCyan
  # Set uac to 514 (0x0202) to notify Bradford to stop access to network
  Set-ADUser -Identity $_.ObjectGUID -Replace @{UserAccountControl = 0x0202 } -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

# =========================================================================================
# Imported Functions
. .\lib\Clear-SessionData.ps1
. .\lib\Load-Module.ps1
. .\lib\New-RandomPassword.ps1
. .\lib\Show-TestRun.ps1

# Processing
Show-TestRun
Clear-SessionData

'SQLServer' | Load-Module

# AD Domain Controller Session
$adCmdLets = 'Get-ADUser', 'Set-ADUser', 'Set-ADAccountPassword'
$adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $adCmdLets -AllowClobber

$activeAD = Get-ActiveAD
$activeAeries = Get-ActiveAeries
$inactiveIDs = Get-InactiveIDs -activeAD $activeAD -activeAeries $activeAeries

Get-InactiveADObj -activeAD $activeAD -inactiveIDs $inactiveIDs | Disable-ADObjects | Set-UserAccountControl |
Set-RandomPassword | Set-GsuiteSuspended | Get-AssignedChromeBookUsers | Update-Chromebooks | Get-SecondaryStudents | Format-Html | Send-AlertEmail

# Get-InactiveADObj -activeAD $activeAD -inactiveIDs $inactiveIDs |
# Get-AssignedChromeBookUsers | Get-SecondaryStudents | Format-Html | Send-AlertEmail

Clear-SessionData
Show-TestRun