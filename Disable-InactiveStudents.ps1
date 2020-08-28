#Requires -Version 5.0
<# 
.SYNOPSIS
 Queries Aeries Student Inforamtion System and Active Directory and determines which AD accounts need to be disabled. 
.DESCRIPTION
 EmployeeIDs Queried from Aeries and Active Directory student user objects are compared. If AD Object employeeID attribute
 is not present in Aeries results then the AD account is disabled, moved to the 'Disabled student OU', and 
 if present the user's GSUite account is suspended as well. A hold date can be place in the user objects 'info' attribute
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
 In special cases an account can be held open until a set date. This is recorded in the info attribute of the user object.
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
 [Alias('wi')]
 [SWITCH]$WhatIf
)

# CLS;$error.clear() # Clear Screen and $error
Get-PSSession | Remove-PSSession -WhatIf:$false

# AD Domain Controller Session
$adCmdLets = 'Get-ADUser', 'Set-ADUser', 'Move-ADObject'
$adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $adCmdLets -AllowClobber

# Variables
$gamExe = '.\lib\gam-64\gam.exe'

# Imported Functions
. .\lib\Add-Log.ps1
. .\lib\Invoke-SqlCommand.ps1 # Useful function for querying SQL and returning results

# Processing
$properties = 'AccountExpirationDate', 'EmployeeID', 'HomePage', 'info'
$allStuParams = @{
 Filter     = { (homepage -like "*@*") -and (employeeID -like "*") }
 SearchBase = 'OU=Students,OU=Users,OU=Domain_Root,DC=chico,DC=usd'
 Properties = $properties
}

$allADStudents = Get-ADUser @allStuParams |
Where-Object { ($_.samaccountname -match "^\b[a-zA-Z][a-zA-Z]\d{5,6}\b$") -and ($_.employeeID -match "^\d{5,6}$") }

$query = Get-Content -Path '.\sql\active-students.sql' -Raw
$sqlParams = @{
 Server     = $SISServer
 Database   = $SISDatabase
 Credential = $SISCredential
 Query      = $query
}
$allSISActiveIds = Invoke-SqlCommand @sqlParams

Add-Log info ("Active AD Students :" + $allADStudents.count)
Add-Log info ("Aeries Active Records :" + $allSISActiveIds.count)

Write-Verbose "Computing difference..."
$inactiveEmpIds = Compare-Object -ReferenceObject $allSISActiveIds -DifferenceObject $allADStudents -Property employeeID |
Where-Object { $_.SideIndicator -eq '=>' }

Add-Log info ("AD Accounts Needing deactivation :" + $inactiveEmpIds.count)

foreach ( $empId in $inactiveEmpIds.employeeID ) {
 $user = $allADStudents.Where( { $_.employeeID -eq $empId })
 if ( !$user ) { continue } # Skip missign users
 $sam = $user.SamAccountName
 $guid = $user.ObjectGUID
 if ( $user.info ) {
  # BEGIN Skip if custom date set in User 'Info' Attrib via json format
  try {
   [datetime]$altExpireDate = Get-Date ($user.info | ConvertFrom-Json).keepUntil
   if ( (Get-Date) -le $altExpireDate ) {
    Add-Log info "$sam,Active until: $altExpireDate"
    # Read-Host 'LOOK!!!!!!!!!!!!!!!!===================================================='
    continue
   }
   else { Add-Log info "$sam,expired: $altExpireDate" }
  }
  catch { Add-Log warning "$sam,User.info missing date and/or json formating" }
 } # END Skip if custom date set in User 'Info' Attrib
 Add-Log disable $sam
 Set-ADUser -Identity $guid -Enabled $False -Whatif:$WhatIf # Disable the account
 Set-ADUser -Identity $guid -Replace @{UserAccountControl = 0x0202 } # Set uac to 514 to notify Bradford to stop access to network
 $disabledOU = 'OU=Disabled_Student_Objects,OU=Disabled_User_Objects,DC=chico,DC=usd'
 Move-ADObject -Identity $guid -TargetPath $disabledOU -Whatif:$WhatIf # move to disabled ou
 # Suspend Gsuite Account
 if ($user.HomePage -and !$WhatIf) { (& $gamExe update user $user.HomePage suspended on) *>$null }
}

Write-Verbose "Tearing down sessions"
Get-PSSession | Remove-PSSession -WhatIf:$false