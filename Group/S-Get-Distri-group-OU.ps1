
# Get group based in OU and export to CSV file.

$arrayOU = 'London','Tokio','Sao Paulo'

$GetGroups = foreach ($OUBase in $arrayOU) {            
    Get-DistributionGroup  -OrganizationalUnit "OU=$OUBase,OU=South America,DC=EXAMPLE,DC=NET" -ResultSize unlimited -WarningAction silentlycontinue
}

foreach ($DistGrp in $GetGroups) {            
    
    $GrpName = $DistGrp.name
    Get-DistributionGroupMember $DistGrp.PrimarySmtpAddress -ResultSize unlimited -WarningAction silentlycontinue
    $report = New-Object psobject
    $report | Add-Member -MemberType NoteProperty -name Grupo -Value $GrpName
    $report | Export-Csv -Path C:\Temp\grupos.csv -NoTypeInformation -Append
}
