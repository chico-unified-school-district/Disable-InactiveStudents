#Requires -Version 5.0
<#
.SYNOPSIS
 Queries Aeries Student Inforamtion System and Active Directory and determines which AD accounts need to be disabled.
.DESCRIPTION
 EmployeeIDs Queried from Aeries and Active Directory student user objects are compared. If AD Object employeeID attribute
 is not present in Aeries results then the AD account is disabled,
 and if present will be used to determine if the account should remain active until the hold date expires.
.EXAMPLE
 .\Disable-InactiveStudents.ps1 -DC $dc -ADCredential $adCreds -SISConnection $sisConn -SISCredential $sisCreds
.EXAMPLE
 .\Disable-InactiveStudents.ps1 -DC $dc -ADCredential $adCreds -SISConnection $sisConn -SISCredential $sisCreds -WhatIf -Verbose
.INPUTS
 Active Directory Domain controller, AD account credentail object with server permissions and various user object permissions
.OUTPUTS
 Log entries are recorded for each operation
.NOTES
 In special cases an account can be held open until a set date. This is recorded in the 'info' attribute of the user object.
 the format to store the date could be for example:
 { "keepUntil": "9/1/2025" }
 or
 { "keepUntil": "Sep 1 2025" }
 or
 { "keepUntil": "Monday, September 1, 2025 12:00:00 AM" }

 These would be an examples of an invalid date reference:
 { "keepUntil": "Sept 1 2025" }
 and this is also invalid:
 { "keepUntil": "Sep 1st 2025" }
#>
[cmdletbinding()]
param (
 [Parameter(Mandatory = $True)]
 [Alias('DC', 'Server')]
 [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 5 })]
 [string]$DomainController,
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
 [string[]]$MailTarget,
 [string[]]$BccAddress,
 [Alias('wi')]
 [SWITCH]$WhatIf
)

# Script Functions
function Format-HTML {
 begin {
  $baseHtml = Get-Content -Path .\html\return_chromebook_message.html -Raw
 }
 process {
  $data = @(
   $_.School
   $_.SPermID
   $_.LasstName
   $_.FirstName
   $_.Parentname
   $_.ParentEMail
   $_.Fatherworkphone
   $_.Motherworkphone
   $_.ParentportalEmail
   $_.Barcode
   $_.serial
   $_.Code1
   $_.Condition
   $_.Comment
   $_.IssuedDate
   $_.Address
  )
  @{
   html = $baseHtml -f $data
   to   = $MailTarget
   cred = $MailCredential
   bcc  = $BccAddress
  }
 }
}
function Get-ActiveADStudents {
 $properties = 'AccountExpirationDate', 'EmployeeID', 'HomePage', 'info'
 $allStuParams = @{
  Filter     = { (homepage -like "*@*") -and (employeeID -like "*") }
  SearchBase = 'OU=Students,OU=Users,OU=Domain_Root,DC=chico,DC=usd'
  Properties = $properties
 }

 Get-ADUser @allStuParams | Where-Object {
  $_.samaccountname -match "^\b[a-zA-Z][a-zA-Z]\d{5,6}\b$" -and
  $_.employeeID -match "^\d{5,6}$" -and
  $_.Enabled -eq $True
 } | Sort-Object Surname
}

filter Get-AssignedChromeBookUsers {
 $sqlParams = @{
  Server     = $SISServer
  Database   = $SISDatabase
  Credential = $SISCredential
 }
 Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name)
 $sql = (Get-Content -Path .\sql\student_return_cb.sq.sql -Raw) -f $_.employeeId
 Invoke-SqlCmd @sqlParams -Query $sql
}

function Disable-ADObjects {
 process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name)
  Set-ADUser -Identity $_.ObjectGUID -Enabled:$false -Confirm:$false -WhatIf:$WhatIf
 }
}

function Select-InactiveADObj {
 begin {
  $sqlParams = @{
   Server     = $SISServer
   Database   = $SISDatabase
   Credential = $SISCredential
  }
  $query = Get-Content -Path '.\sql\active-students.sql' -Raw
  $aeriesActive = Invoke-SqlCmd @sqlParams -Query $query
 }
 process {
  if ($aeriesActive.employeeID -notcontains $_.employeeId) {
   Write-Host "$($_.employeeId), Inactive student found"
   $_
  }
 }
}

function Send-AlertEmail {
 process {
  $mailParams = @{
   To         = $_.to
   From       = $EmailCredential.Username
   Subject    = $_.subject
   bodyAsHTML = $true
   Body       = $_.body
   SMTPServer = 'smtp.office365.com'
   Cred       = $EmailCredential
   UseSSL     = $True
   Port       = 587
  }
  if ($BccAddress) { $mailParams += @{Bcc = $BccAddress } }

  if (-not$WhatIf) { Send-MailMessage @mailParams }
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, ($_.to -join ','), $_.subject)
 }
}

function Set-RandomPassword {
 Process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name)
  $randomPW = ConvertTo-SecureString -String (New-RandomPassword) -AsPlainText -Force
  Set-ADAccountPassword -Identity $_.ObjectGUID -NewPassword $randomPW -Confirm:$false -WhatIf:$WhatIf
 }
}

function Set-GsuiteSuspended {
 process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name)
  if ($_.HomePage -and -not$WhatIf) { (& $gamExe update user $_.HomePage suspended on) *>$null }
 }
}

function Set-UserAccountControl {
 process {
  Write-Host ('{0},{1}' -f $_.name, $MyInvocation.MyCommand.name)
  # Set uac to 514 to notify Bradford to stop access to network
  Set-ADUser -Identity $_.ObjectGUID -Replace @{UserAccountControl = 0x0202 } -Confirm:$false -WhatIf:$WhatIf
 }
}

# Variables
$gamExe = '.\lib\gam-64\gam.exe'

# Imported Functions
. .\lib\Clear-SessionData.ps1
. .\lib\Load-Module.ps1
. .\lib\New-RandomPassword.ps1
. .\lib\Show-TestRun.ps1

Show-TestRun

Clear-SessionData

'SQLServer' | Load-Module

# Processing

# AD Domain Controller Session
$adCmdLets = 'Get-ADUser', 'Set-ADUser', 'Set-ADAccountPassword'
$adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $adCmdLets -AllowClobber

# Get-ActiveADStudents | Select name | ft
Get-ActiveADStudents | Select-InactiveADObj | Disable-ADObjects 
# | Set-UserAccountControl | Set-RandomPassword
# | Set-GsuiteSuspended | Get-AssignedChromeBookUsers | Send-AlertEmail

Clear-SessionData
Show-TestRun