# Get mailbox using user arguments interactive.
# Used for helpdesk find a specific mailbox.


Write-host " 
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=
=-=-=-=-= Exchange Server User Search =-=-=-=-=
=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=
1 .To search if E-MAIL is being used.
2 .To search if USERNAME is being used.
"-ForeGround "Blue" 

Write-Host "               " 
$number = Read-Host "Choose one of the options above (1) ou (2) " 

switch ($number){ 
1 {
    $Email = Read-Host "Enter email address: " 
    Get-Mailbox $Email -ResultSize unlimited
    }

2 { 
    $userName = Read-Host "Enter username: " 
    Get-Mailbox $userName -ResultSize unlimited
    } 
}







