# Send e-mails by powershell.

$From = "wsys@EXAMPLE.com"
$To = "cl.barros@EXAMPLE.com"
$Anexo = "C:\temp\temp.txt"
$Subject = "Teste PS script"
$Body = "WSYS 10.0.0.30 e-mail de teste relay"
$SMTPServer = "smtp-external.online"
$SMTPPort = "587"
$c = Get-Credential 

#Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -Credential $c -Attachments $Anexo

Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -Credential $c -Attachments $Anexo

#Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -Attachments $Anexo