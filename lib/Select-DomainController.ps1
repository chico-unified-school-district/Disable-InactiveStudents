function Select-DomainController ([string[]]$DomainControllers) {
 $dc = Get-Random $DomainControllers
 if ( Test-Connection -ComputerName $dc -Count 1 -ErrorAction SilentlyContinue ) {
  Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.Name, $dc)
  return $dc
 }
 else {
  $msg = $MyInvocation.MyCommand.Name, $dc
  Write-Host ('{0},{1} Not responding. Trying random Domain Controller in 30 seconds...' -f $msg)
  Start-Sleep 30
  Select-DomainController $DomainControllers
 }
}