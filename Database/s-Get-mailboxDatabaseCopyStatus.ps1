# Get database copy health for any database which begins with AMER.
$Mb_Rpl_Status = (Get-MailboxDatabaseCopyStatus -Identity AMER*).status | Sort-Object -Unique

ForEach ($i in $Mb_Rpl_Status){
    
    $RPL_Status = @('Healthy','Mounted')

    if ($RPL_Status -contains $i ) {
        Write-Output "OK"
    }
    else {
        Write-Output "Failed"
        Break
    }
}
