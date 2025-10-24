# DomainChanges
Methods to change aliases or UPNs in bulk.

AliasAdd.ps1 - Use this add an alias (and make it primary) to everything: Users, M365 Groups, Distribution Lists"  
UserUPNChange.ps1 - Use this update UPN of users (if needed) based on a CSV file  

# UserUPNChange.ps1
Use this update UPN of users (if needed) based on a CSV file    

On first run, this script will create an input CSV for the users to adjust.  
` DisplayName,AliasAndUPNToEnforce`  
` Business Owner,owner@contoso.com`  


Users are looked up based on DisplayName.  
Then the AliasAndUPNToEnforce will be added and made primary (if needed).  
*Note: If nothing is needed, the object will be left alone and you will see a confirmation message.* 


# How it works
- Checks for some modules and installs them if needed
``` PowerShell
Install-Module Microsoft.Graph
```

- Connects to org and confirms it's the right org
``` PowerShell
Connect-MgGraph
``` 
- For each row in the CSV, asks if you want to process that row, or just continue processing all rows without confirmation    
`Process Entry? (Yes or Yes to All or Exit)`  

- Runs through these commands  
``` PowerShell
Get-MgUser -Filter "displayName eq '$displayName'"  
   UserPrincipalName
   ProxyAddresses
Update-MgUser -UserId $user.Id -UserPrincipalName $newUPN
```


