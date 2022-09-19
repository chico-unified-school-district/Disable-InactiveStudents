function New-ADSession ($dc, $cmdlets, $cred) {
 $msgVars = $MyInvocation.MyCommand.Name, $dc, ($cmdLets -join ',')
 Write-Host ('{0},{1}' -f $msgVars)
 $adSession = New-PSSession -ComputerName $dc -Credential $cred
 Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $cmdLets -AllowClobber | Out-Null
}