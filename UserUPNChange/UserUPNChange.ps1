#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####

#################### Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
#################### Transcript Open

### Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
############
if (!(Test-Path $scriptCSV))
{
    ######### Template
    "DisplayName,AliasAndUPNToEnforce" | Add-Content $scriptCSV
    "Business Owner,owner@contoso.com" | Add-Content $scriptCSV
    ######### Template
	Write-Host "Couldn't find '$(Split-Path $scriptCSV -leaf)'. Template CSV created. It will open for editing now."
    PressEnterToContinue
    Invoke-Item $scriptCSV
    Write-Host "When done editing the CSV file."
    PressEnterToContinue
}
## ----------Fill $entries with contents of file or something
$entries=@(import-csv $scriptCSV)
#$entries
$entriescount = $entries.count
##
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in M365"
Write-Host ""
Write-Host "Use this update UPN of users (if needed) based on a CSV file."
Write-Host ""
Write-Host ""
Write-Host "CSV: $(Split-Path $scriptCSV -leaf) ($($entriescount) entries)"
$entries | Format-Table
Write-Host "-----------------------------------------------------------------------------"
PressEnterToContinue
$no_errors = $true
$error_txt = ""
$results = @()
#region modules
$modules=@()
$modules+="Microsoft.Graph.Authentication"
$modules+="Microsoft.Graph.Groups"
$modules+="Microsoft.Graph.Identity.DirectoryManagement"
ForEach ($module in $modules)
{ 
    Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module -checkver $false; Write-Host $lm_result
    if ($lm_result.startswith("ERR")) {
        Write-Host "ERR: Load-Module $($module) failed. Suggestion: Open PowerShell $($PSVersionTable.PSVersion.Major) as admin and run: Install-Module $($module)";Start-sleep  3; Return $false
    }
}
#endregion modules
#region Connections
if ($true) { # Connect-MgGraph
    $mgScopes = @() 
    $mgScopes += "User.ReadWrite.All"
    $mgScopes += "Mail.ReadWrite"
    $mgScopes += "Directory.ReadWrite.All"
    if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
        Write-Host "Connect-MgGraph is NOT available. You may need to install the Microsoft Graph module:"
        Write-Host 'Install-Module Microsoft.Graph -Scope CurrentUser'
        PressEnterToContinue
        exit
    }
    # Check if we are already connected
    while ($true) {
        # Check if already connected to Microsoft Graph
        $mgContext = Get-MgContext
        if ($mgContext -and $mgContext.Account -and $mgContext.TenantId) {
            $tenantDomain = (Get-MgDomain | Where-Object { $_.IsDefault }).Id
            Write-Host "Connect-MgGraph is connected to Account $($mgContext.Account) Tenant Domain: " -NoNewline
            Write-Host $tenantDomain -ForegroundColor Green
            $response = AskForChoice "Choice: " -Choices @("&Use this connection","&Disconnect and try again","E&xit") -ReturnString
            # If the user types 'exit', break out of the loop
            if ($response -eq 'Disconnect and try again') {
                Write-Host "Disconnect-MgGraph..."
                Disconnect-MgGraph | Out-Null
                PressEnterToContinue "Done. Press <Enter> to connect again."
                Continue # loop again
            }
            elseif ($response -eq 'exit') {
                return
            }
            else { # on to next step
                break
            }
        } else {
            Write-Host "Connect-MgGraph not connected. Connecting now..."
            PressEnterToContinue "Open a browser to an admin session on the desired tenant and press Enter."
            Connect-MgGraph -Scopes $mgScopes -NoWelcome
            # Confirm connection
            $mgContext = Get-MgContext
            if ($mgContext) {
                Write-Host "Now connected to Microsoft Graph as $($mgContext.Account)"
                $tenantDomain = (Get-MgDomain | Where-Object { $_.IsDefault }).Id
                Write-Host "Tenant Domain: $tenantDomain" -ForegroundColor Green
            } else {
                Write-Error "Connect-MgGraph: Failed"
            }
        }
    } # while true forever loop
    Write-Host
} # Connect-MgGraph
if ($false) { # Connect-ExchangeOnline
    # Check if Connect-ExchangeOnline is available
    if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: 'Connect-ExchangeOnline' command was not found."
        Write-Host "Please install the ExchangeOnlineManagement module using:"
        Write-Host "   Install-Module ExchangeOnlineManagement"
        Write-Host "Or load the module if it is already installed, then try again."
        Write-Host "Press any key to exit..."
        PressEnterToContinue
        exit
    }
    # Check if we are already connected
    while ($true) {
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            $connected_eol=$true
        }
        catch {
            $connected_eol=$false
        }
        if ($connected_eol)
        { # already connected
            # The Identity property typically shows your tenant's name or domain
            $tenantNameOrDomain = $orgConfig.Identity
            Write-Host "Connect-ExchangeOnline is connected to tenant: " -NoNewline
            Write-host $tenantNameOrDomain -ForegroundColor Green
            $response = AskForChoice "Choice: " -Choices @("&Use this connection","&Disconnect and try again","E&xit") -ReturnString
            # If the user types 'exit', break out of the loop
            if ($response -eq 'Disconnect and try again') {
                Write-Host "Disconnect-ExchangeOnline..."
                $null = Disconnect-ExchangeOnline -Confirm:$false
                PressEnterToContinue "Done. Press <Enter> to connect again."
                Continue # loop again
            }
            elseif ($response -eq 'exit') {
                return
            }
            else { # on to next step
                break
            }
        } # already connected
        else
        { # not connected
            Write-Host "Connect-ExchangeOnline is not connected."
            # check powershell version
            if ($PSVersionTable.PSVersion.Major -lt 7)
            { # PS 5
                Write-Host "We will try 'Connect-ExchangeOnline'. Use admin creds to authenticate."
                PressEnterToContinue
                Write-Host "Connect-ExchangeOnline ... " -ForegroundColor Yellow
                Connect-ExchangeOnline -ShowBanner:$false
            } # PS 5
            else 
            { # PS 7
                $choice = AskForChoice "Connect method (PS7)" -Choices @("&Browser code","&Password","E&xit") -Default 0 -ReturnString
                if ($choice -eq "Exit") { exit }
                if ($choice -eq "Browser code")
                { # Browser code
                    Write-Host "1. Open a browser to an admin session on the desired tenant"
                    Write-Host "2. Copy the code below"
                    Write-Host "3. Click the link and paste the code to authenticate"
                    Write-Host "Connect-ExchangeOnline -Device ... " -ForegroundColor Yellow
                    Connect-ExchangeOnline -ShowBanner:$false -Device
                } # Browser code
                else { # Password
                    Write-Host "We will try 'Connect-ExchangeOnline' to authenticate. Use admin creds to authenticate."
                    PressEnterToContinue
                    Write-Host "Connect-ExchangeOnline ... " -ForegroundColor Yellow
                    Connect-ExchangeOnline -ShowBanner:$false
                } # Password
            } # PS 7
            Write-Host "Done" -ForegroundColor Yellow
            Continue # loop again
        } # not connected
    } # while true forever loop
    Write-Host
} # Connect-ExchangeOnline
#endregion Connections
# continue?
$processed=0
$message="$entriescount Entries. Continue?"
$choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","&No")
[int]$defaultChoice = 0
$choiceRTN = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
if ($choiceRTN -eq 1)
{ "Aborting" }
else 
{ ## continue choices
    $choiceLoop=0
    $i=0        
    foreach ($x in $entries)
    { # each entry
        $i++
        write-host "-----" $i of $entriescount $x
        if ($choiceLoop -ne 1)
        {
            $message="Process entry "+$i+"?"
            $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","Yes to &All","&No","No and E&xit")
            [int]$defaultChoice = 1
            $choiceLoop = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
        }
        if (($choiceLoop -eq 0) -or ($choiceLoop -eq 1))
        { # choiceloop
            $processed++
            ####### Start code for object $x
            # CSV column names
            $displayName = $x.DisplayName
            $newDomain = $x.AliasAndUPNToEnforce.Split("@")[1]   # domain part
            # Step 1: Find user by display name
            $user = Get-MgUser -Filter "displayName eq '$displayName'"
            if (-not $user) {
                Write-Host "User with display name '$displayName' not found." -ForegroundColor Yellow
                PressEnterToContinue
                return
            }
            # Step 2: Check for alias in new domain
            $localPart = $user.UserPrincipalName.Split("@")[0]
            $newUPN = "$localPart@$newDomain"
            # Fetch full user object including proxyAddresses
            $userDetail = Get-MgUser -UserId $user.Id -Property "proxyAddresses,userPrincipalName"
            $hasAlias = $userDetail.ProxyAddresses -match ("smtp:$newUPN")
            if (-not $hasAlias) {
                # Add the alias to the proxyAddresses list
                $updatedProxies = $userDetail.ProxyAddresses + @("smtp:$newUPN")
                # Update-MgUser -UserId $user.Id -ProxyAddresses $updatedProxies  (This isn't needed and errors)
                Write-Host "[UPDATE] Alias smtp:$newUPN added." -ForegroundColor Yellow
            } else {
                Write-Host "[OK] Alias smtp:$newUPN already exists." -ForegroundColor Green
            }
            # Step 3: Update UPN if needed
            if ($userDetail.UserPrincipalName -ne $newUPN) {
                Update-MgUser -UserId $user.Id -UserPrincipalName $newUPN
                Write-Host "[UPDATE] UPN updated to $newUPN" -ForegroundColor Yellow
            } else {
                Write-Host "[OK] UPN already set to $newUPN" -ForegroundColor Green
            }
            ####### End code for object $x
        } # choiceloop
        if ($choiceLoop -eq 2)
        {
            write-host ("Entry "+$i+" skipped.")
        }
        if ($choiceLoop -eq 3)
        {
            write-host "Aborting."
            break
        }
    } # each entry
} ## continue choices
WriteText "Removing any open sessions..."
Get-PSSession 
Get-PSSession | Remove-PSSession
WriteText "------------------------------------------------------------------------------------"
$message ="Done. $($processed) of $($entriescount) entries processed. Press [Enter] to exit."
WriteText $message
WriteText "------------------------------------------------------------------------------------"
#################### Transcript Save
Stop-Transcript | Out-Null
$date = get-date -format "yyyy-MM-dd_HH-mm-ss"
New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
$TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
Move-Item $Transcript $TranscriptTarget -Force
#################### Transcript Save

PressEnterToContinue