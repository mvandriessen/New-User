$OUList = Get-ADOrganizationalUnit -Filter *
$DesiredOU = NULL
$found = $false
$O365URL = "https://ps.outlook.com/powershell"
$O365Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"
$O365Session = Connect-MsolService -Credential $O365Credentials

$ExchangeCredentials = Get-Credential -Message "Enter your Domain Admin credentials"

#Detect Exchange Version
$ExchangeVersion = Get-Command  Exsetup.exe | ForEach-Object {$_.FileversionInfo}

if ($ExchangeVersion -like "14")
{
    #Exchange 2010 commands
    Add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
    get-exchangeserver
}
elseif ($ExchangeVersion -like "15")
{
    #Exchange 2013 commands
    $Exchangeserver = (Get-ADDomainController).HostName
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchangeserver/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredentialss
}
elseif ($ExchangeVersion -like "16")
{
    #Exchange 2016 commands
}
else
{
    Write-Host "Unsupported Exchange version"
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


