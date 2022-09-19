function Select-DomainController([string[]]$DomainControllers) {
 foreach ($dc in $DomainControllers) {
  Write-Verbose ('Checking {0},{1}' -f $MyInvocation.MyCommand.Name, $dc)
  if (Test-Connection -ComputerName $dc -Count 1) {
   $dc
   return
  }
 }
 $msgVars = $MyInvocation.MyCommand.Name, ($DomainControllers -join ',')
 Write-Host ('{0},{1},No Controller is reachalbe at this time. Waiting a bit.' -f $msgVars)
 Start-Sleep 30
 Select-DomainController $DomainControllers
}