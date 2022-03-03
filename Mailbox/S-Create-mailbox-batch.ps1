<#
.NOTES
Create batch mailboxes based on the CSV CreateMailboxes.csv file.
   Adjust the "OrganizationalUnit" parameter as necessary.
   If it is necessary to add or remove parameters, it is necessary to change the command and the CSV file.

CSV file example:
MBXName,MBXFirstName,MBXLastName,PasswordMBX,DisplayName,UPN
joao,joao,Ferreira,P@ssw0rd,João Ferreira,joao.ferreira@example.com

#>

Import-CSV CreateMailboxes.csv | ForEach-Object {
    New-Mailbox -Name $_.MBXName  -OrganizationalUnit "OU=corp,DC=domain,DC=com" ´
     -FirstName $_.MBXFirstName -LastName $_.MBXLastName -Password (ConvertTo-SecureString $_.PasswordMBX -AsPlainText -Force) ´
     -DisplayName $_.DisplayName -ResetPasswordOnNextLogon $false -UserPrincipalName $_.UPN
    }
