<#
Get count of messagens in Exchange Server queues.
#>


if ( (Get-PSSession).ConfigurationName -ne 'Microsoft.Exchange'  -and (Get-PSSession).ComputerName -ne 'mail.amer.example.com'  ) {
    Write-Host "There is no session for Microsoft Exchange, you must connect to the exchange server." -BackgroundColor Yellow -ForegroundColor Black
    break 
}

$ExServer = (Get-ExchangeServer).name
foreach ($iExServer in $ExServer) {
    [string]$iExServerStr = $iExServer
    Write-Host "Server --> "$iExServerStr -ForegroundColor Blue
    Get-Queue -Server $iExServerStr | Where-Object {$_.MessageCount -gt 10} | Format-Table -AutoSize
}

Write-Host "<...................................................>" -ForegroundColor Green
Write-Host "Only queues with more than 10 messages are displayed." -ForegroundColor Green
Write-Host "Column <MessageCount> indicates the number of messages in the queue." -ForegroundColor Green
