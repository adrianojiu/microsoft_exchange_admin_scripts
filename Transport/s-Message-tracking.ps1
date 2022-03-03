<#
.SYNOPSIS
    Script to search for traking log of messages sent and / or received in the Exchange server.
    Script to search for tracking log of messages sent and / or received on the Exchange server.
.DESCRIPTION
    Adjustment must be made to the variables for searching e-mail transactions.
    Variable you need without alterations:
    -> $ StartDate = "11/16/2020 00:00" Start date that you want to search for transactions. / Start date you want to search for transactions. Date in MM / DD / YYYY HH: MM format
    -> $ DataFim = "11/16/2020 23:59" End date that you want to search for transactions. / End date you want to search for transactions. Date in MM / DD / YYYYs format HH: MM
    -> $ SenderEmail = "adriano.ferreira@domain.com" Email address of the person sending the email. / E-mail address of the person sending the e-mail.
    -> $ RecipientEmail = "csc.vmi.pr@domain.com" Email address from which you should receive the email. / Email address from which you should receive the email.
.NOTES
    File Name: s-Message-tracking.ps1
    Author: Adriano Ferreira (adriano.ferreira@domain.com)
    Prerequisite: PowerShell session with Exchange server.
    Copyright 2020: Adriano de Oliveira Ferreira
.LINK
    Script posted over:
    https://github.com/adrianojiu/SharedAdminScripts/tree/main/ExchangeServer
.EXAMPLE
    . \ s-Message-tracking.ps1
    Note: the example above assumes that you are in the same folder where the script is.
         example above assumes that you are in the same folder where the script is.
         In case of execution policy issues you should disable or bypass execution policy.
.EXAMPLE
#>

$DataInicio = "03/31/2021 00:00"
$DataFim = "03/31/2021 23:59"
$SenderEmail = "fiscal01@gmail.com"
$RecipientEmail = "JuniodaSilva.Nunes@example.com"

$ExServer = Get-ExchangeServer | Where-Object {$_.isHubTransportServer -eq $true -or $_.isMailboxServer -eq $true}
foreach ($iExServer in $ExServer) {
    [string]$iExServerStr = $iExServer
    Write-Host "Server where the log was collected $iExServerStr."  -ForegroundColor DarkBlue
    Get-MessageTrackingLog -server $iExServerStr -Sender $SenderEmail -Recipients $RecipientEmail -Start $DataInicio -End $DataFim -ResultSize Unlimited | Where-Object {$_.EventId -inotlike "HA*"} |Select-Object Timestamp,'  ',Source,EventId,'   ',MessageSubject,RecipientStatus,Recipients,Sender | Format-Table -AutoSize
    #Get-MessageTrackingLog -server $iExServerStr -Sender $SenderEmail -Recipients $RecipientEmail -Start $DataInicio -End $DataFim -ResultSize Unlimited |Select-Object Timestamp,'  ',Source,EventId,'   ',MessageSubject,RecipientStatus,Recipients,Sender | Format-Table -AutoSize
    #Get-MessageTrackingLog -server $iExServerStr -Sender $SenderEmail -Start $DataInicio -End $DataFim -ResultSize Unlimited |Select-Object Timestamp,'  ',Source,EventId,'   ',MessageSubject,RecipientStatus,Recipients,Sender | Format-Table -AutoSize 
    #Get-MessageTrackingLog -server $iExServerStr -Sender $SenderEmail -Start $DataInicio -End $DataFim -ResultSize Unlimited | Export-Csv -Path c:\temp\track.csv -NoTypeInformation -Encoding utf8
}

Write-Host "To search for 'EventId' and 'Source' access:" -ForegroundColor DarkYellow
Write-Host "https://docs.microsoft.com/pt-br/exchange/mail-flow/transport-logs/message-tracking?view=exchserver-2019" -ForegroundColor DarkGreen
