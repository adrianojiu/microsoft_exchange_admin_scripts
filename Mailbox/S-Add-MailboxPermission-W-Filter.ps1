
<#
.NOTES  
  Assigns full access to the user svc-mcs-excadm full access permission for all accounts with the exception of Admin account.
   "Automapping" parameter is set to false so that the accounts are not mapped to the user's outlook svc-mcs-excadm (who will be allowed to access the boxes).
#>

Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'UserMailbox') -and (Get-Alias -ne 'Admin')} | ´
Add-MailboxPermission -User svc-mcs-excadm ´
-AccessRights fullaccess -InheritanceType all -AutoMapping:$false