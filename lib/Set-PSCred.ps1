function Set-PSCred {
 [CmdletBinding()]
param (
 # Username for credential object
 [Parameter(Position=0,Mandatory=$true)]
 [string]$Username,
 # password string for credential object
 [Parameter(Position=1,Mandatory=$true)]
 [string]$Password
)
$securePw = ConvertTo-SecureString $Password -AsPlainText -Force
New-Object System.Management.Automation.PSCredential ($Username,$securePw )
}