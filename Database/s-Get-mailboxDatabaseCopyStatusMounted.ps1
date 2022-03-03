
#Returns number of Active databases that do not have "Mounted" status.
$Mb_Rpl_Status = (Get-MailboxDatabaseCopyStatus -Identity AMER* |  Where-Object {$_.ActiveCopy -eq $True -and $_.Status -ne "Mounted"}).count
   
    if ( $Mb_Rpl_Status -eq 0) {
        Write-Output "OK"
    }
    else {
        Write-Output "FAIL"
        Break
    }
