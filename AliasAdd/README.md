# DomainChanges
Methods to change aliases or UPNs in bulk.

AliasAdd.ps1 - Use this add an alias (and make it primary) to everything: Users, M365 Groups, Distribution Lists"  
UserUPNChange.ps1 - Use this update UPN of users (if needed) based on a CSV file  

# AliasAdd.ps1
Use this add an alias (and make it primary) to everything: Users, M365 Groups, Distribution Lists  

This script will prompt for the new domain name.  e.g. newdomain.com  
You can run it in *view only* mode to see what changes it will make without touching anything.  

It will then loop through these objects in Entra to add the new domain (if needed) and make it primary (if needed).  
*Note: If nothing is needed, the object will be left alone and you will see a confirmation message for that object.* 
- Users  
- M365 Groups  
- Distribution Lists  


# How it works
- Checks for some modules and installs them if needed
``` PowerShell
Install-Module Microsoft.Graph
Install-Module ExchangeOnlineManagement
```

- Connects to org and confirms it's the right org
``` PowerShell
Connect-MgGraph
Connect-ExchangeOnline
``` 
- Asks if you want to make changes or just view them  
`Ready to go live or just checking. Yes: Make changes, No: Just checking?`  

- Runs through these commands  
``` PowerShell
Get-MgUser
   UserPrincipalName
   ProxyAddresses
Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')"
   ProxyAddresses
   MailNickname
Get-DistributionGroup
   PrimarySmtpAddress
```