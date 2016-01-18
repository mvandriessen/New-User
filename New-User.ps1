$OUList = Get-ADOrganizationalUnit -Filter *
$DesiredOU = $NULL
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
$ADDomain = Read-Host -Prompt "Enter the AD Domain name"
$Title = Read-Host -Prompt "Enter the title of the user. Leave blank if not needed"
$Domain = Read-Host -Prompt "Enter your email domain."
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
$NewUser = New-ADUser -SamAccountName $UserName -UserPrincipalName $UserName@$ADDomain -Name $name -GivenName $FirstName -Surname $LastName -Path $DesiredOU -AccountPassword $Password -Enabled $true

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

        Enable-Mailbox -Identity $NewUser.SamAccountName
    }
    elseif ($ExchangeVersion -like "15")
    {
        #Exchange 2013 commands
        $Exchangeserver = (Get-ADDomainController).HostName
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangeserver/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredentials
        Import-PSSession -Session $ExchangeSession

        Enable-Mailbox -Identity $NewUser.SamAccountName
    }
    elseif ($ExchangeVersion -like "16")
    {
        #Exchange 2016 commands
        $Exchangeserver = (Get-ADDomainController).HostName
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangeserver/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredentials
        Import-PSSession -Session $ExchangeSession

        Enable-Mailbox -Identity $NewUser.SamAccountName
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