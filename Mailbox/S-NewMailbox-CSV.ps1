<#
.SYNOPSIS
 Create mailbox based on the parameters below.

.EXAMPLE
-> Example of script execution, execution within the powershel connected to the Exchange Server.
    . \ S-NewMailbox-CSV.ps1

-> Example of CSV file content
    Index; DisplayName; FirstName; LastName; PrimarySmtpAddress; userName; OrganizationalUnit; Password; state; Street; PostalCode; City; telephoneNumber; Department; Company; Office
    2; Mailbox Test18 (Test); Mailbox18; Test18; test.18EXAMPLE1.com; mbteste18; OU = Tests, OU = Users, OU = Sao Paulo, OU = South America, DC = EXAMPLE, DC = NET; P @ EE215651a; SP; Brig Faria Lima 201 151 and; 05416-020; Sao Paulo; +5515 350415000; TI; EXAMPLE1 Imp Exp SA; Sao Paulo

------------------------ Script information ---------------------- -
These scripts are used for mailbox creation on Microsoft Exchange.

Prerequisites:
 - You must be connected to the Microsoft Exchange Powershell.
 - The files must be in the same folder.
 - To simplify, copy the script to your computer (always do this, as it may have been updated).
 - You must always fill in the CSV file in all fields. Creation is always based on the S-NewMailbox-CSV.CSV file, which uses a ";" separator.
 - Always execute the script "S-NewMailbox-CSV-ValidateCSV.ps1", to validate the CSV file before executing "S-NewMailbox-CSV.ps1".
 - The first field of the CSV index file must be filled in sequentially and the first line with account information must be 2.

* Mailboxes are created in Rotteram databases, with bsae in the filter when the database has the name "NLRT".

* The domain controller used is "'NLRTDSRV01001.EXAMPLE.NET'" if it is not operational it must be changed to another operational one.

* It is possible to create a single mailbox or several with just one execution of the script.

* After creating the CSV file and validating, use "S-NewMailbox-CSV.ps1" to create the mailbox.

* It may be necessary to change the execution policy to execute the script (Set-ExecutionPolicy Unrestricted).

* To execute the Powershell validation script, access the folder where the script is located and execute. \ S-NewMailbox-CSV-ValidateCSV.ps1
- The script accesses the "S-NewMailbox-CSV.CSV" file and validates it, based on the information contained in the file.

* To execute the Powershell account creation script, access the folder where the script is located and execute. \ S-NewMailbox-CSV.ps1
- The script accesses the file "S-NewMailbox-CSV.CSV" and creates it, based on the information contained in the file.

#>

$ImpCSV = Import-Csv -Path .\S-NewMailbox-CSV.CSV -Delimiter ";" -Encoding utf8 -Verbose
$ImpCSV | ForEach-Object {
$Email =        $_.PrimarySmtpAddress
$UserName =     $_.username
$Pword =        $_.Password
$DisplayName =  $_.'DisplayName'
$FirstName =    $_.'FirstName'
$LastName =     $_.'LastName'
$OU =           $_.'OrganizationalUnit'
$State =        $_.'state'
$Street =       $_.'street'
$PostalCode =   $_.'PostalCode'
$City =         $_.'City'
$PhoneNumber =  $_.'telephoneNumber'
$Department =   $_.'Department'
$Company =      $_.'Company'
$Office =       $_.'Office'
$nowTimeIni =   Get-Date -Format "dd/MM/yyyy HH:mm K"
$NLRTdb =       (Get-MailboxDatabase | where-object {$_.name -like "*NLRT*"}).name | Get-Random
$NLRTdc =       'DCServer.EXAMPLE.NET'

Write-Host "start"
$nowTimeIni

#################################################################
#  Checking if there is any value and blank in the CSV list..  #
#################################################################
$AttributesArray = @($Email,$UserName,$Pword,$DisplayName,$FirstName,$LastName,$OU,$State,$Street,$PostalCode,$City,$PhoneNumber,$Department,$Company,$Office)
ForEach($iAttributesArray in $AttributesArray){
    if (!$iAttributesArray) {
        $AttributesArray
        Write-Host "---------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
        Write-Host "All fields must contain value, validate the list above for fields with no value and enter the value in the CSV file." -ForegroundColor Yellow
        Write-Host "---------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
    break
    }
}

####################################################
# Checking remote Microsoft Exchange session. #
####################################################
try { Get-Mailbox -ResultSize 1 -WarningAction SilentlyContinue | Out-Null   }
catch { 
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "You are not connected to Microsoft Exchange, please connect before continuing." -ForegroundColor Yellow 
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor Yellow
    break
}

################################################
#  Checking if the email address has @.  #
################################################
if ($AttributesArray[0] -notlike '*@*') {
    Write-Host "-----------------------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "The format of the email address appears to be incorrect, check the value in the CSV file." -ForegroundColor Yellow 
    Write-Host "-----------------------------------------------------------------------------------------" -ForegroundColor Yellow
    break
}

#############################################
#  Checking for mailbox existence.  #
#############################################
$GetMailbox = Get-Mailbox $Email -EA SilentlyContinue -DomainController $NLRTdc
$GetMailboxIdentity = Get-Mailbox $Email -ErrorAction SilentlyContinue -DomainController $NLRTdc | Select-Object Identity -ExpandProperty Identity
if($GetMailbox.PrimarySmtpAddress -eq $Email){
    Write-Host "------------------------------------------------------------------------------------------"
    Write-Host "ERROR - The email address " -NoNewline -ForegroundColor Yellow; Write-Host "$Email" -NoNewline -ForegroundColor DarkRed; Write-Host " already in use, try again with another email, object in use: " -NoNewline -ForegroundColor Yellow; Write-Host $GetMailboxIdentity"." -ForegroundColor DarkRed
    Write-Host "------------------------------------------------------------------------------------------"
    break
}

#########################################
#  Checking for username.  #
#########################################
try {  $GetADuser = Get-ADUser $userName -Server $NLRTdc}
catch { }
Finally{  $GetADuser | Out-Null  }

if($GetADuser.samaccountname -eq $userName){
    Write-Host "------------------------------------------------------------------------------------------"
    Write-Host "User name " -ForegroundColor Yellow -NoNewline; Write-Host  $username -ForegroundColor DarkRed -NoNewline; Write-Host " already exists in AD use another username." -ForegroundColor Yellow
    Write-Host "------------------------------------------------------------------------------------------"
    break    
}

##################################################
#  Validating password number of characters.  #
##################################################
if ($Pword.Length -lt 8) {
    Write-Host "------------------------------------------------------------------------------------------------------"
    Write-Host "The password must have at least 8 characters, with lowercase letters, numbers and special characters." -ForegroundColor Yellow
    Write-Host "Try again with the correct password pattern." -ForegroundColor DarkGreen
    Write-Host "------------------------------------------------------------------------------------------------------"
    break
}

#######################################
#  Checking for existence of contact. #
#######################################
$GetMailContact = Get-MailContact $Email -EA SilentlyContinue -DomainController $NLRTdc
if($GetMailContact.PrimarySmtpAddress -eq $Email){
    Write-Host "--------------------------------------------------------------------------------------------"
    Write-Host "An email contact already exists using this email address, use another email address." -ForegroundColor Yellow
    Write-Host "--------------------------------------------------------------------------------------------"
    Get-MailContact $Email | Select-Object OrganizationalUnit | Format-List
    break
}

################################################################
#  Check if there is an email group with an email address. #
################################################################
$GetDistributionGroup = Get-DistributionGroup $Email -EA SilentlyContinue -DomainController $NLRTdc
if($GetDistributionGroup.PrimarySmtpAddress -eq $Email){
    Write-Host "------------------------------------------------------------------------------------------"
    Write-Host "There is already an email group using this email address, use another email address." -ForegroundColor Yellow
    Write-Host "------------------------------------------------------------------------------------------"
    Get-DistributionGroup $Email | Select-Object OrganizationalUnit | Format-List
    break
}

#######################################################
# Creating Mailbox and setting values in the AD user. # 
#######################################################
New-Mailbox -Name "$DisplayName" -DisplayName "$DisplayName" -FirstName $FirstName -LastName $LastName -PrimarySmtpAddress $Email -UserPrincipalName "$userName@EXAMPLE.net" -OrganizationalUnit $OU -Password (ConvertTo-SecureString $Pword -AsPlainText -Force) -ResetPasswordOnNextLogon $false -RetentionPolicy "Global-Agriculture Standard Mailbox Management and Retention" -DomainController $NLRTdc -Database $NLRTdb
Set-ADUser -Identity $UserName -Replace @{c="BR"; co="Brazil"} -State $State -StreetAddress $Street -PostalCode $PostalCode -City $City -OfficePhone $PhoneNumber -Department $Department -Company $Company -Office $Office -Server $NLRTdc

#######################################
#  Mailbox creation summary. #
#######################################
Write-Host "-----------------------------------"
Write-Host "Mailbox creation summary." -ForegroundColor DarkGreen
Write-Host "-----------------------------------"
Get-Mailbox $Email -DomainController $NLRTdc | Select-Object Name,FirstName,LastName,DisplayName,PrimarySmtpAddress,Database,RetentionPolicy,ServerName,OrganizationalUnit
Get-ADUser -Identity $UserName -Properties * -Server $NLRTdc | Select-Object samaccountname,State,StreetAddress,PostalCode,City,OfficePhone,Department,Company,Office,c,co

Write-Host "End"
$nowTimefim = Get-Date -Format "dd/MM/yyyy HH:mm K"
$nowTimefim

}
