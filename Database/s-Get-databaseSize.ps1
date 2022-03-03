
# Get databse size for all DBs which begins with DBEX-.
Get-MailboxDatabase | Where-Object {$_.name -match "^DBEX-"} | Get-MailboxDatabase -status  | Select-Object Name,DatabaseSize,AvailableNewMailboxSpace

# Formated Get databse size for all DBs which begins with DBEX-.
Get-MailboxDatabase | Where-Object {$_.name -match "^DBEX-"} | Get-MailboxDatabase -Status | Sort-Object name | Select-Object name,@{Name='DB Size (Gb)';Expression={$_.DatabaseSize.ToGb()}},@{Name='Available New Mbx Space Gb)';Expression={$_.AvailableNewMailboxSpace.ToGb()}}
