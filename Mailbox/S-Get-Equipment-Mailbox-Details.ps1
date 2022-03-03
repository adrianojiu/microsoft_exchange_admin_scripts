<#
Exports equipment box properties to CSV.
   To add more properties it is necessary to add after the Select-Object.
 #>

Get-Mailbox -RecipientTypeDetails EquipmentMailbox | `
Select-Object Displayname,PrimarySmtpAddress,RecipientTypeDetails, ModerationEnabled, samAccountName, legacyExchangeDN, Identity | `
  Export-Csv -Path .\EquipmentMailbox.csv -Encoding utf8

