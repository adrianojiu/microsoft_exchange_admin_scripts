<# 
.SYNOPSIS 
.\Get-mailbox-Statistic.ps1
 Return report about mailbox size.

1 .Display Mailbox Statistics in screen.
2 .Export Mailbox Statistics to CSV file.
  
.Author 
Written By: Adriano Ferreira
 

#> 
 
Write-host " 
=-=-=-=-=-=-=-=-=-=-=-=-= 
=- Mailbox Size Report -=
=-=-=-=-=-=-=-=-=-=-=-=-=
1 .Display Mailbox Statistics in screen.
2 .Export Mailbox Statistics to CSV file.
"-ForeGround "Blue" 

Write-Host "               " 
$number = Read-Host "Choose one of option above" 
$output = @() 

function Get-MBX-Fun {
    $filiais = Import-Csv -Path .\CSV-Get-mailBox-properties-SPO-CGB.csv
    $AllMailbox = foreach ($filial in $filiais){
        $filialOU = $filial.OU
        $filialC = $filial.name
       
        Get-Mailbox -OrganizationalUnit "OU=$filialOU,OU=South America,DC=EXAMPLE,DC=NET" -ResultSize unlimited -WarningAction silentlycontinue | Where-Object {$_.Name -like "*$filialC*"} 
                        
        }
        return $AllMailbox
    }


    switch ($number){ 
############ 1. Display Mailbox Statistics in screen ############
1 { 
    $AllMailbox = Get-MBX-Fun

    Foreach($Mbx in $AllMailbox){ 
        $Stats = Get-mailboxStatistics $Mbx.PrimarySmtpAddress
        $GetDepartment = (Get-User $Mbx.Alias | Select-Object department | Format-Table -HideTableHeaders | Out-String).Trim()
                
        $userObj = New-Object PSObject 
        $userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname 
        $userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $mbx.PrimarySmtpAddress
        $userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $stats.TotalItemSize
        $userObj | Add-Member NoteProperty -Name "Server" -Value $Stats.ServerName
        $userObj | Add-Member NoteProperty -Name "Department" -Value $GetDepartment
        Write-Output $Userobj
        } 
    ;Break} 

##################### 2. Export Mailbox Statistics to CSV file. #####################
2 { 
    $CSVfile = Read-Host "Enter the Path of CSV file (Eg. Report-MBX-stats.csv)'"
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
    ;Break} 
}
