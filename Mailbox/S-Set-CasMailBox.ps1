

#$arrayOU = 'Cuiaba','Guararapes','Sao Paulo'

$CASMBX = Import-Csv -Path C:\temp\CSV-Set-CASMailbox.csv

foreach ($CASMBXUser in $CASMBX) {            
    
    Get-CASMailbox $CASMBXUser.MBX
    #Get-CASMailbox $CASMBXUser.MBX | Set-CASMailbox -PopEnabled $true
    
}
