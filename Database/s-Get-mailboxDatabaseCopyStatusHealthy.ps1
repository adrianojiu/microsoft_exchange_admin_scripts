Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;


# Returns the number of copied databases that do not have the status "Healthy".
$Mb_Rpl_Status = (Get-MailboxDatabaseCopyStatus -Identity AMER* |  Where-Object {$_.ActiveCopy -eq $false -and $_.Status -ne "Healthy"}).count
   
    if ( $Mb_Rpl_Status -eq 0) {
        Write-Output "OK"
    }
    else {
        Write-Output "FAIL"
        Break
    }

