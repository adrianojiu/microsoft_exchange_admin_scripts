# Get mailbox move.
# It gets password encoded and open a session for Microsoft Exchange.

$password = get-content .\Pass-t1-encoded\cred.txt | convertto-securestring
$cred= New-Object System.Management.Automation.PSCredential ("username", $password )

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://mail.amer.example.com/PowerShell/ -Credential $cred -AllowRedirection -Authentication Negotiate
Import-PSSession -Session $Session -AllowClobber

$GM = (Get-MoveRequest | Where-Object { $_.status -ne "Completed" }).displayname
$GMALL = @()
ForEach($i in $GM){
    $GMALL += Get-MoveRequestStatistics $i | Select-Object DisplayName,StatusDetail,PercentComplete  | Format-Table -HideTableHeaders| Out-String 
}

$MoveRequestCount = (Get-MoveRequest | Where-Object { $_.status -ne "Completed" }).count
if ($MoveRequestCount -gt 0) {

    Send-MailMessage -From 'adriano.ferreira@example.com' -To 'adriano.f@example.com ', 'daniel.o@example.com', 'marcelo.k@example.com'`
     -Subject 'Relatorio de move de caixas postais' -Body " ---> Total de caixas postais em move $MoveRequestCount `n `n *** Display name ***       *** Status ***       *** Precent Complete *** `n $GMALL"`
       -SmtpServer 'brspdsmtp.br.example.com'
    
}
return $GMALL
