<#
   Execute the command that is inside the DO until without interrupted (kill the process or execution.).
   In this case, the command is to show the percentage of import of pst for postal bounces that match the string "migtest".
   Leave 10 seconds to display the next command execution.
#>

$a = "ok"

DO
    {
 
        Get-MailboxImportRequest | Where-Object {$_.Mailbox -like "*migtest*"} | Get-MailboxImportRequestStatistics | Where-Object {$_.PercentComplete -ne "100"}
        Write-Host "Proximo get" -BackgroundColor Yellow -ForegroundColor Black
        Start-Sleep 10
  
    }   While ($a -like "ok")
