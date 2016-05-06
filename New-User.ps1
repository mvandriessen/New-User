$O365Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"
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
$Loginscript = Read-Host -Prompt "Enter the login script for the user. Leave blank if not needed"
$Homefolder = Read-Host -Prompt "Enter the home folder of the user. Leave blank if not needed"
$name = "$FirstName $LastName"

#Get list of all OUs and prompt for desired one
#To do: convert to function to allow parametrisation
$DesiredOU = $NULL
$found = $false
$OUList = Get-ADOrganizationalUnit -Filter *
$oulist | select name | ft
$ANS = $NULL

#Loop through OU list untill one is found
do
{
    $ANS = (Read-Host "Choose from these OUs")
    foreach ($ou in $OUList)
    {
        if ($ou.Name -like "*$ANS*")
        {
            Write-Host "OU Found" -ForegroundColor Yellow
            Write-host $ou
            $found = $true
            $DesiredOU = $ou
        }
    }

    if (-not $found)
    {
        Write-Host "OU not found! Please try again" -ForegroundColor Red
    }
}while($found -eq $false)


#Create new AD user
try 
{
    $NewUser = New-ADUser -SamAccountName $UserName -UserPrincipalName $UserName@$ADDomain -Name $name -GivenName $FirstName -Surname $LastName `
        -Path $DesiredOU -Title $Title -HomeDirectory $Homefolder -ScriptPath $Loginscript -AccountPassword $Password -Enabled $true
}
catch [System.Object] 
{
    Write-Output "Could not create user $name."
    break
}


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
        Write-Host "Mailbox created" -ForegroundColor Yellow
    }
    elseif ($ExchangeVersion -like "15")
    {
        #Exchange 2013 commands
        $Exchangeserver = (Get-ADDomainController).HostName
        $success = $false
        while($success -eq $false)
        {
            $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangeserver/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredentials -ErrorAction Ignore
            if($?)
            {
                $success = $true
                Import-PSSession -Session $ExchangeSession
                Write-host "Session to $ExchangeServer has been opened" -ForegroundColor Yellow 
            }
            else 
            {
                Write-Host "Please enter a valid Exchange server: " -ForegroundColor Red -NoNewline
                $Exchangeserver = Read-Host
            }
        }
        
        Enable-Mailbox -Identity $NewUser.SamAccountName
        Write-Host "Mailbox created" -ForegroundColor Yellow
    }
    elseif ($ExchangeVersion -like "16")
    {
        #Exchange 2016 commands
        $success = $false
        while($success -eq $false)
        {
            $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangeserver/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredentials -ErrorAction Ignore
            if($?)
            {
                $success = $true
                Import-PSSession -Session $ExchangeSession
                Write-host "Session to $ExchangeServer has been opened" -ForegroundColor Yellow 
            }
            else 
            {
                Write-Host "Please enter a valid Exchange server: " -ForegroundColor -NoNewline
                $Exchangeserver = Read-Host
            }
        }
        
        Enable-Mailbox -Identity $NewUser.SamAccountName
        Write-Host "Mailbox created" -ForegroundColor Yellow
    }
    else
    {
        Write-Host "Unsupported Exchange version" -ForegroundColor Red
        Exit
    }
}
elseif($location -like "365")
{
    Write-host "Creating user in Office 365" -ForegroundColor Yellow

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
            else
            {
                Write-Host "License not found! User was not created." -ForegroundColor Red 
                break
            }
        }
        else 
        {
            Write-Host "No more Office 365 licenses available! Order more in the portal" -ForegroundColor Red
            break
        }
    }
    
}