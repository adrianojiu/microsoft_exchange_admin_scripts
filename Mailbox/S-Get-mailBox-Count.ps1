# Get and sumarize mailboxes.

# CSV with branch information.
$filiais = Import-Csv -Path C:\temp\Get-mailBox-properties.csv

# Create result header for visual reference only.
Write-Host "Data, Filial, Total, Migradas, Restantes"

foreach ($filial in $filiais) {            
    $filialC = $filial.name
    $filialOU = $filial.OU
    $filialCD = $filial.cidade
    $nowDate = Get-Date -Format "dd-MM-yyyy"
    
    #  Account number of boxes per branch, TOTAL.
    $MBXCountFilial = 0
    $MBXCountFilial = (Get-Mailbox -OrganizationalUnit "OU=$filialOU,OU=South America,DC=EXMPLE,DC=NET" -ResultSize unlimited -WarningAction silentlycontinue | Where-Object {$_.Name -like "*$filialC*"}).count
    if ($null -eq $MBXCountFilial){ 
        $MBXCountFilial = 1 }
        
    #  Account number of boxes not migrated, REMAINING.
    $MBXCountNoMig = (Get-Mailbox -OrganizationalUnit "OU=$filialOU,OU=South America,DC=EXMPLE,DC=NET" -ResultSize unlimited -WarningAction silentlycontinue | Where-Object {$_.Name -like "*$filialC*" -and $_.ServerName -like "clstg*"}).count
    if ($null -eq $MBXCountNoMig){ 
        $MBXCountNoMig = 1 }
        
    #  Account number of migrated boxes, MIGRATED.
    $MBXCountMig = (Get-Mailbox -OrganizationalUnit "OU=$filialOU,OU=South America,DC=EXMPLE,DC=NET" -ResultSize unlimited -WarningAction silentlycontinue | Where-Object {$_.Name -like "*$filialC*" -and $_.ServerName -like "nlrtd*"}).count
    if ($null -eq $MBXCountMig){ 
        $MBXCountMig = 1 }

    Write-Host "$nowDate,$filialCD,$MBXCountFilial,$MBXCountMig,$MBXCountNoMig"
        
    $report = New-Object psobject
    $report | Add-Member -MemberType NoteProperty -name Data -Value $nowDate
    $report | Add-Member -MemberType NoteProperty -name Site -Value $filialC
    $report | Add-Member -MemberType NoteProperty -name Total -Value $MBXCountFilial
    $report | Add-Member -MemberType NoteProperty -name Migradas -Value $MBXCountMig
    $report | Add-Member -MemberType NoteProperty -name Restantes -Value $MBXCountNoMig

    $report | Export-Csv -Path C:\Temp\TestTemp\$filialC-$nowDate.csv -NoTypeInformation

}
