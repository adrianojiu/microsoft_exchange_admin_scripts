<#
.SYNOPSIS 
Cria caixa postal baseado nos parametros abaixo.

.EXAMPLE
Write-Host 'Exemplo de execuçao do script - .\S-NewMailbox-Param.ps1 -userName teste.mbx3 -Email 3teste.mailbox@EXAMPLE.com -Pword a123@456P -OU "OU=Tests,OU=Users,OU=Sao Paulo,OU=South America,DC=EXAMPLE,DC=NET" -DisplayName Teste3 script mailbox3 -FirstName Mailbox3 -LastName Teste3'.
#>

<#
Param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]$userName,

    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]$Email,
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]$Pword,
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]$OU,
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]$FirstName,
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]$LastName,
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNull()]
    [string]$DisplayName
)
#>

#############################################
#  Checking for mailbox existence.  #
#############################################
$GetMailbox = Get-Mailbox $Email -EA SilentlyContinue
$GetMailboxIdentity = Get-Mailbox $Email -ErrorAction SilentlyContinue | Select-Object Identity -ExpandProperty Identity

if($GetMailbox.PrimarySmtpAddress -eq $Email){
    Write-Host "-----------------------------"
    Write-Host "ERRO - The email address " -NoNewline -ForegroundColor Yellow; Write-Host "$Email" -NoNewline -ForegroundColor DarkRed; Write-Host " already in use, try again with another email, object in use: " -NoNewline -ForegroundColor Yellow; Write-Host $GetMailboxIdentity"." -ForegroundColor DarkRed
    Write-Host "-----------------------------"
    break
}

#########################################
#  Checking for username.  #
#########################################
try {  $GetADuser = Get-ADUser $userName  }
catch { }
Finally{  $GetADuser | Out-Null  }

if($GetADuser.samaccountname -eq $userName){
    Write-Host "-----------------------------"
    Write-Host "User name " -ForegroundColor Yellow -NoNewline; Write-Host  $username -ForegroundColor DarkRed -NoNewline; Write-Host " already exists in AD use another username." -ForegroundColor Yellow
    Write-Host "-----------------------------"
    break    
}

##################################################
#  Validating password number of characters.  #
##################################################
if ($Pword.Length -lt 8) {
    Write-Host "-----------------------------"
    Write-Host "The password must have at least 8 characters, with lowercase letters, numbers and special characters." -ForegroundColor Yellow
    Write-Host "Try again with the correct password pattern." -ForegroundColor DarkGreen
    Write-Host "-----------------------------"
    break
}

######################################
#  Checking for existence of contact.#
######################################
$GetMailContact = Get-MailContact $Email -EA SilentlyContinue

if($GetMailContact.PrimarySmtpAddress -eq $Email){
    Write-Host "-----------------------------"
    Write-Host "An email contact already exists using this email address, use another email address." -ForegroundColor Yellow
    Write-Host "-----------------------------"
    Get-MailContact $Email | Select-Object OrganizationalUnit | Format-List
    break
}

################################################################
#  Check if there is an email group with an email address. #
################################################################
$GetDistributionGroup = Get-DistributionGroup $Email -EA SilentlyContinue

if($GetDistributionGroup.PrimarySmtpAddress -eq $Email){
    Write-Host "-----------------------------"
    Write-Host "There is already an email group using this email address, use another email address." -ForegroundColor Yellow
    Write-Host "-----------------------------"
    Get-DistributionGroup $Email | Select-Object OrganizationalUnit | Format-List
    break
}

###################################
# Creating Mailbox, if it does not exist.
#if($GetMailbox.Name -ne $Email){
New-Mailbox -Name "$DisplayName" -DisplayName "$DisplayName" -FirstName $FirstName -LastName $LastName -PrimarySmtpAddress $Email -UserPrincipalName "$userName@EXAMPLE.net" -OrganizationalUnit $OU -Password (ConvertTo-SecureString $Pword -AsPlainText -Force) -ResetPasswordOnNextLogon $false -RetentionPolicy "Global-Agriculture Standard Mailbox Management and Retention"

#}

#######################################
#  Mailbox creation summary. #
#######################################
Write-Host "Mailbox creation summary." -ForegroundColor DarkGreen
Get-Mailbox $Email | Select-Object Name,FirstName,LastName,DisplayName,PrimarySmtpAddress,Database,RetentionPolicy,ServerName
