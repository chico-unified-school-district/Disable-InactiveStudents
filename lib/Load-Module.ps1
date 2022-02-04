function Load-Module {
 process {
  Write-Host ('{0}: {1}' -f $MyInvocation.MyCommand.Name, $_) -ForegroundColor Green
  if (-not(Get-Module -Name $_ -ListAvailable)) {
   Install-Module -Name $_ -Scope CurrentUser -AllowClobber -Confirm:$false -Force
  }
  Import-Module -Name $_ -Force -ErrorAction Stop -Verbose:$false | Out-Null
  # Get-Module -Name $_
 }
}