﻿$OUList = Get-ADOrganizationalUnit -Filter *
$DesiredOU = NULL
$found = $false
$O365URL = "https://ps.outlook.com/powershell"
$O365Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"
$O365Session = Connect-MsolService -Credential $O365Credentials

$ExchangeCredentials = Get-Credential -Message "Enter your Domain Admin credentials"

#Populate variables
$FirstName = Read-Host -Prompt "Enter the first name"
$LastName = Read-Host -Prompt "Enter the last name"
$UserName = Read-Host -Prompt "Enter the user name"
$Password = Read-Host -Prompt "Enter the password" -AsSecureString
$Title = Read-Host -Prompt "Enter the title of the user. Leave blank if not needed"
$name = $FirstName + ' ' + $LastName

#Get list of all OUs and prompt for desired one
#To do: convert to function to allow parametrisation
$oulist | select name | ft
$ANS = (Read-Host "Choose from these OUs")

foreach ($ou in $OUList)
{
    if ($ou.DistinguishedName -like "*$ANS*")
    {
        Write-Host "Found"
        Write-host $ou
        $found = $true
        $DesiredOU = $ou
        break
    }
}

if (-not $found)
{
    Write-Host "NO GOOD"
}

#Create new AD user
New-ADUser -SamAccountName $UserName -Name $name -GivenName $FirstName -Surname $LastName -Path $DesiredOU -Enabled $true

#Prompt for 356 or local exchange
$location = Read-Host -Prompt "Office 365 or exchange"
if(!($location -like "365"))
{

    #Detect Exchange Version
    $ExchangeVersion = Get-Command  Exsetup.exe | ForEach-Object {$_.FileversionInfo}

    if ($ExchangeVersion -like "14")
    {
        #Exchange 2010 commands
        Add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
        $ExchangeServer = Get-ExchangeServer
    }
    elseif ($ExchangeVersion -like "15")
    {
        #Exchange 2013 commands
        $Exchangeserver = (Get-ADDomainController).HostName
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangeserver/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredentials
    }
    elseif ($ExchangeVersion -like "16")
    {
        #Exchange 2016 commands
        $Exchangeserver = (Get-ADDomainController).HostName
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangeserver/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredentials
    }
    else
    {
        Write-Host "Unsupported Exchange version"
        Exit
    }
}
elseif($location -like "365")
{
    Write-host "Office 365"
    New-MsolUser
}


#Write-host "Connected"

<#switch ($a) 
    { 
        1 {$ou} 
        2 {"The color is blue."} 
        3 {"The color is green."} 
        4 {"The color is yellow."} 
        5 {"The color is orange."} 
        6 {"The color is purple."} 
        7 {"The color is pink."}
        8 {"The color is brown."} 
    }
#>



