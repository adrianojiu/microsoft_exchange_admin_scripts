<# 
.SYNOPSIS 
.\Get-mailbox-Statistic.ps1
 Return report about mailbox size.

1 .Display M

ailbox Statistics in screen.
2 .Export Mailbox Statistics to CSV file.
  
.Author 
Written By: Adriano Ferreira
 
Change Version
V1.0, 10/31/2019 - Initial version 
#> 
cd C:\Run-tasks\Get-Mailbox-Use-Report
Start-Transcript .\trascript.txt
$password = get-content .\Pass-t1-encoded\cred.txt | convertto-securestring
$cred= New-Object System.Management.Automation.PSCredential ("username", $password ) # Using credential encodade in file; You might change to get-credential if needed

$Session = New-PSSession -ConfigurationName Microsoft.Exchange.S-GetmailBox-Statistic-Sendemail -ConnectionUri https://mail.amer.example.com/PowerShell/ -Credential $cred -AllowRedirection -Authentication Negotiate
Import-PSSession -Session $Session -AllowClobber

$output = @() 

function Get-MBX-Fun {
    # Set OU to get mailboxes
    $filiais = "example.net/AMER/Brazil","examples.net/AMER/Uruguay","example.net/AMER/Argentina","ecamples.net/AMER/Colombia"
    $AllMailbox = foreach ($filial in $filiais){
                       
        Get-Mailbox -OrganizationalUnit $filial -ResultSize unlimited -WarningAction SilentlyContinue
        }
        return $AllMailbox
    }

##################### Export Mailbox Statistics to CSV file. #####################
  
    $CSVfile = ".\S-Get-mailBox-Statistic-teste.csv"
    $AllMailbox = Get-MBX-Fun

    Foreach($Mbx in $AllMailbox) { 
        $Stats = Get-mailboxStatistics $Mbx.PrimarySmtpAddress
        $GetDepartment = (Get-User $Mbx.Alias | Select-Object department | Format-Table -HideTableHeaders | Out-String).Trim()
       
        $userObj = New-Object PSObject 
        $userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname 
        $userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias 
        $userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientType 
        $userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit 
        $userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress 
        $userObj | Add-Member NoteProperty -Name "Database" -Value $Stats.Database 
        $userObj | Add-Member NoteProperty -Name "ServerName" -Value $Stats.ServerName 
        $userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
        $userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount 
        $userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount 
        $userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize 
        $userObj | Add-Member NoteProperty -Name "DatabaseProhibitSendReceiveQuota" -Value $Stats.DatabaseProhibitSendReceiveQuota 
        $userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime
        $userObj | Add-Member NoteProperty -Name "Department" -Value $GetDepartment
        $output += $UserObj   
    } 
    $output | Export-csv -Path $CSVfile -NoTypeInformation  -Encoding UTF8 -Delimiter ";"
            
    Send-MailMessage -From 'adriano.f@example.com' -To 'adriano.f@example.com ' -Subject 'Relatorio de uso de caixas postais' -Body "Relatorio de uso de caixas postais." -Attachments $CSVfile -SmtpServer 'brspdsmtp.br.example.com'
    
    Stop-Transcript
    #;Break

