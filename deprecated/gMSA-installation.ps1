#  Thomas Verheyden - 09/04/24

#GMSA Credentials
$GMSA = Read-Host "Enter GMSA account name"

$DomainRole = Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty DomainRole
Write-Host "Script is running on a" -ForegroundColor Gray
 switch ($DomainRole)
   {
   0 {Write-Host "Stand alone workstation" -ForegroundColor Red}
   1 {Write-Host "Member workstation" -ForegroundColor Red}
   2 {Write-Host "Stand alone server" -ForegroundColor Red}
   3 {Write-Host "Member server" -ForegroundColor Red}
   4 {Write-Host "Back-up domain controller" -ForegroundColor Green}
   5 {Write-Host "Primary domain controller" -ForegroundColor Green}
   default {Write-Host "The role can not be determined"}
   }


If (($DomainRole -eq 4) -or ($DomainRole -eq 5))
    {
    Write-Host "Prerequisites are good to go!" -ForegroundColor Green
    }
Else 
    {
    Write-Host "You need to run this script on a Domain Controller!" -ForegroundColor Red
    Exit
    }

Write-Host "Importing Active Directory module" -ForegroundColor Gray
Try {
    Get-Module -listavailable | Where-Object {$_.name -like "*ActiveDirectory*"} | Import-Module -Force -Erroraction stop
    Write-Host "Active Directory module loaded..." -ForegroundColor Green
}
Catch {
    Write-Host "Active Directory module not found!" -ForegroundColor Red
    Exit
}

#Create KdsRootKey if not exist
$KDSRootKeyID  = @((Get-KdsRootKey).KeyID)
$TestKDSRootKey = Test-KdsRootKey -KeyId $KDSRootKeyID[0]
If ($TestKDSRootKey -eq "True") 
    {
        Write-host "KDSRootKey already exist!" -ForegroundColor Green
    }
Else
    {
        Write-host "KDSRootKey does not exist. Creating KDSRootKey..." -ForegroundColor Red
        Add-KdsRootKey –EffectiveTime ((get-date).addhours(-10)) 
        Start-Sleep 5
        Write-host "KDSRootKey created..."  -ForegroundColor Green
        
    }
        

$DomainController = @((Get-ADComputer -Filter 'primarygroupid -eq "516"').Name) | ForEach-Object {$_ + "$"}
Write-Host "Found the following domain controllers:" $DomainController -ForegroundColor Green
$Domain = (Get-ADDomain).DnsRoot
Write-Host "DNS Root Domain:" $Domain -ForegroundColor Green
$DomainShort = (Get-WmiObject Win32_NTDomain).DomainName
Write-Host "Domain Name:" $DomainShort -ForegroundColor Green
$DNSHostname = $GMSA+"."+$Domain
Write-Host "DNS Hostname:" $DNSHostname -ForegroundColor Green


# Only use to include all domains and trusted domain in a forest
#read: https://techcommunity.microsoft.com/t5/microsoft-365-defender-blog/using-gmsa-account-in-microsoft-defender-for-identity-in-multi/ba-p/2942690
$DCGroup = get-adgroup "name of the universal group containing all dc's"
New-ADServiceAccount $GMSA -DNSHostName $DNSHostname -PrincipalsAllowedToRetrieveManagedPassword $DCGroup -KerberosEncryptionType RC4, AES128, AES256 -ServicePrincipalNames http/$GMSA.$Domain/$Domain, http/$GMSA.$Domain/$DomainShort, http/$GMSA/$Domain, http/$GMSA/$DomainShort


#Write-Host "Creating GMSA" $GMSA -ForegroundColor Green
#New-ADServiceAccount $GMSA -DNSHostName $DNSHostname -PrincipalsAllowedToRetrieveManagedPassword $DomainController -KerberosEncryptionType RC4, AES128, AES256 -ServicePrincipalNames http/$GMSA.$Domain/$Domain, http/$GMSA.$Domain/$DomainShort, http/$GMSA/$Domain, http/$GMSA/$DomainShort
#Set-ADServiceAccount $GMSA -PrincipalsAllowedToDelegateToAccount $DCGroup
#Set-ADServiceAccount $GMSA -PrincipalsAllowedToRetrieveManagedPassword $DCGroup

#Review Properties
Write-Host "Get properties" $GMSA -ForegroundColor Green
Get-ADServiceAccount -Identity $GMSA -Properties PrincipalsAllowedToRetrieveManagedPassword

#Install managemed service account on every domain controller
# If problems with kerberos do this: klist purge -li 0x3e7
Write-Host "Install and test" "$GMSA" -ForegroundColor Green
Install-ADServiceAccount $GMSA
Test-ADServiceAccount $GMSA

