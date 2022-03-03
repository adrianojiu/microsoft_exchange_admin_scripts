<#
    .Description
    Remote powershell connection to Microsoft Exchange Server 2013/2016/2019 .
#>

$ExcMgt = "mail.eur.EXAMPLE.net"

$TestPortConnection = (Test-NetConnection -ComputerName $ExcMgt -Port 443 -WarningAction SilentlyContinue).TcpTestSucceeded

if ($TestPortConnection -like "False") {
  Write-Host " xxx Connection to server $ExcMgt failed. xxx" -ForegroundColor DarkRed
  break
}
else {
  Write-Host "# OK # Connection test to server $ExcMgt success." -ForegroundColor DarkGreen
}

write-host "# Type Exchange server admin, Eg. EXAMPLE\user. #" -ForegroundColor black -BackgroundColor DarkYellow

$credential = get-credential
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$ExcMgt/PowerShell" -credential $credential

Write-Host "# Importing and opening into Exchange server. #"  -ForegroundColor black -BackgroundColor DarkGreen
Import-PSSession $ExchangeSession -AllowClobber –Verbose

# Testing connection running a simple commando.
get-exchangeServer -identity $ExcMgt | Select-Object name, site, serverRole
  
