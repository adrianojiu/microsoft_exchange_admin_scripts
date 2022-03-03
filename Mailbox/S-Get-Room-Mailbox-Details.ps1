<#
   Exports properties of meeting room boxes to CSV.
   To add more properties it is necessary to add after the Select-Object.
 #>

Get-Mailbox -RecipientTypeDetails RoomMailbox | `
 Select-Object Displayname,PrimarySmtpAddress,RecipientTypeDetails, ModerationEnabled, samAccountName, legacyExchangeDN, Identity | `
  Export-Csv -Path .\RoomMailbox.csv -Encoding utf8

