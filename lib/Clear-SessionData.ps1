function Clear-SessionData {
 Write-Host 'Clearing session and module data' -ForegroundColor Blue
 Get-Module -name *tmp* | Remove-Module -Confirm:$false -Force
 Get-PSSession | Remove-PSSession -Confirm:$false
}