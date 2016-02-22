$OUList = Get-ADOrganizationalUnit -Filter *
$DesiredOU = $NULL
$found = $false
$O365URL = "https://ps.outlook.com/powershell"
$O365Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"
$O365Session = Connect-MsolService -Credential $O365Credentials

$ExchangeCredentials = Get-Credential -Message "Enter your Domain Admin credentials"

#Populate variables
#will add error handling for wrong input
$FirstName = Read-Host -Prompt "Enter the first name"
$LastName = Read-Host -Prompt "Enter the last name"
$UserName = Read-Host -Prompt "Enter the user name"
$Password = Read-Host -Prompt "Enter the password" -AsSecureString
$ADDomain = Read-Host -Prompt "Enter the AD Domain name"
$MailDomain = Read-Host -Prompt "Enter the e-mail domain name"
$Title = Read-Host -Prompt "Enter the title of the user. Leave blank if not needed"
$Domain = Read-Host -Prompt "Enter your email domain."
$name = "$FirstName $LastName"

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
        #This assumes Exchange is on the DC. Will add error handling, allowing to specify exchange server.
        $Exchangeserver = (Get-ADDomainController).HostName
        $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangeserver/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredentials
        Import-PSSession -Session $ExchangeSession

        Enable-Mailbox -Identity $NewUser.SamAccountName
    }
    elseif ($ExchangeVersion -like "16")
    {
        #Exchange 2016 commands
        #This assumes Exchange is on the DC. Will add error handling, allowing to specify exchange server.
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
    Write-host "Creating user in Office 365"

    #Detect Office 365 licenses and let user choose 
    Write-Host "Detecting Office 365 licenses" -ForegroundColor Yellow
    Connect-MsolService -Credential $O365Credentials

    $O365Licenses = Get-MsolAccountSku
    Write-Host $O365Licenses.AccountSkuId
    $DesiredLicense = Read-Host -Prompt "What license do you want to assign to this user?"
    foreach ($O365License in $O365Licenses)
    {
        #Calculate ammount of free licenses
        $O365FreeLicenses = $O365Licenses.ActiveUnits - $O365Licenses.WarningUnits - $O365Licenses.ConsumedUnits

        #If there are free licenses, create the user. If not throw an error
        if($O365FreeLicenses -gt 0)
        {
            if($O365License.AccountSkuID -like "*$DesiredLicense*")
            {
               <#
                    author: MatthewG
                    URL: http://stackoverflow.com/questions/28352141/convert-a-secure-string-to-plain-text

                    Convert secure string password to plain text
               #>
               $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
               $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
               
               #Create new user in Office 365 with predefined variables
               New-MsolUser -UserPrincipalName "$UserName@$MailDomain" -Password $PlainPassword -DisplayName $name -FirstName $FirstName -LastName $LastName -UsageLocation BE -LicenseAssignment $O365License.AccountSkuId -StrongPasswordRequired $false
               $PlainPassword = $null
               Write-Host "User $name created with "$O365Licenses.AccountSkuID" license." -ForegroundColor Yellow
               $O365FreeLicenses -= 1
               Write-Host "You now have $O365FreeLicenses licenses remaining" -ForegroundColor Yellow
               
               break
            }
            else{Write-Host "License not found! User was not created." -ForegroundColor Red}
        }
        else 
        {
            Write-Host "No more Office 365 licenses available! Order more in the portal" -ForegroundColor Red
            break
        }
    }
    
}