Function MakePrimaryinProxy {
    param (
        [string]$address,
        [string[]]$proxyAddresses
    )
    $proxyWithPrimary = @()
    # Set $proxyStatus to "notfound" "primary" "nonprimary"
    # Check if the alias already exists
    $proxyTest = $proxyAddresses | Where-Object {$_.ToLower().EndsWith($address.ToLower())} | Select-Object -First 1
    if ($proxyTest) {
        if ($proxyTest.StartsWith("SMTP")) { # already primary (do nothing)
            $proxyStatus = "primary"
            $proxyWithPrimary += $proxyAddresses
        }
        else { # found but not primary (make it uppercase)
            $proxyStatus = "nonprimary" 
            $proxyWithPrimary = $proxyAddresses
            foreach ($proxyAddress in $proxyAddresses) {
                $prxy = $proxyAddress
                if ($prxy.StartsWith("SMTP")) { # demote primary
                    $prxy = $prxy.Replace("SMTP","smtp") 
                }
                if ($prxy.ToLower().EndsWith($address)) { # promote primary
                    $prxy = $prxy.Replace("smtp","SMTP") 
                }
                $proxyWithPrimary += $prxy # appeend
            }
        }
    }
    else { # not found, add it as primary
        $proxyStatus = "notfound"
        foreach ($proxyAddress in $proxyAddresses) {
            $prxy = $proxyAddress
            if ($prxy.StartsWith("SMTP")) { # demote primary
                $prxy = $prxy.Replace("SMTP","smtp") 
            }
            $proxyWithPrimary += $prxy # appeend
        }
        $proxyWithPrimary += "SMTP:$address" # add primary
    }
    return @($proxyWithPrimary, $proxyStatus)
}
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
##
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in M365"
Write-Host ""
Write-Host "Use this add an alias (and make it primary) to everything: Users, M365 Groups, Distribution Lists"
Write-Host "Caution: This will change the UserPrincipalName of users.  They will need to use the new UPN to log in."
Write-Host ""
Write-Host "-----------------------------------------------------------------------------"
# Load settings
$csvFile = "$($scriptDir )\$($scriptBase) Settings.csv"
$settings = CSVSettingsLoad $csvFile
# Defaults
$settings_updated = $false
if ($null -eq $settings.NewDomain) {$settings.NewDomain = "newdomain.com"; $settings_updated = $true}
if ($settings_updated) {$retVal = CSVSettingsSave $settings $csvFile; Write-Host "Initialized - $($retVal)"}
do {
    Write-Host "Domain to add (to everything that needs it): " -NoNewline
    Write-Host $settings.NewDomain -ForegroundColor Yellow
    $domainok = AskForChoice "Is this domain correct?"
    if (-not $domainok) {
        Write-Host "Enter the desired domain:"
        $settings.NewDomain = Read-Host "NewDomain (e.g. newdomain.com) (blank to abort) [$($settings.NewDomain)]"
        if ($settings.NewDomain -eq "") {Write-Host "Aborting"; Start-sleep -Seconds 2; exit}
        # Save Settings
        $retVal = CSVSettingsSave $settings $csvFile
        Write-Host $retVal
    }
} until ($domainok)
$NewDomain = $settings.NewDomain
Write-Host "NewDomain: $($NewDomain)"
#region modules
$modules=@()
$modules+="Microsoft.Graph.Authentication"
$modules+="Microsoft.Graph.Groups"
$modules+="Microsoft.Graph.Identity.DirectoryManagement"
$modules+="Microsoft.Graph.Users"
$modules+="ExchangeOnlineManagement"
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
    $mgScopes += "Group.ReadWrite.All"
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
if ($true) { # Connect-ExchangeOnline
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
$makechanges = AskForChoice "Ready to go live or just checking. Yes: Make changes, No: Just checking?"
if ($makechanges) {
    Write-Host "You chose to MAKE CHANGES" -ForegroundColor Red
} else {
    Write-Host "You chose to just check (no changes will be made)" -ForegroundColor Yellow
}
PressEnterToContinue
# === USERS ===
$objs = Get-MgUser -All -Property @("Id", "DisplayName", "UserPrincipalName", "MailNickname", "ProxyAddresses","UserType")
$objs = $objs | Where-Object { $_.UserType -eq "Member" -and $_.ProxyAddresses -ne $null }
$i_total = $objs.Count
$i_count = 0
ForEach ($obj in $objs | Sort-Object DisplayName) {
    $primary = $obj.ProxyAddresses | Where-Object {$_.StartsWith("SMTP")}
    Write-Host "$((++$i_count)) of $($i_total): $($obj.DisplayName) ($primary) " -NoNewline
    $user = $obj
    $nick = $user.UserPrincipalName.Split("@")[0]
    if (-not $nick) {
        Write-Host " [Skipping user $($user.DisplayName) no usable alias]"
        PressEnterToContinue
        return
    }
    $alias = "$($nick)@$($NewDomain)"
    if ($user.UserPrincipalName -ne $alias) {
        if (-not $makechanges) {
            Write-Host " [Would change UPN to $alias]" -ForegroundColor Yellow
        }
        else {
            Update-MgUser -UserId $user.Id -UserPrincipalName $alias
            Write-Host " [Changed UPN to $alias]" -ForegroundColor Yellow
        }
    } else {
        Write-Host " [UPN already set to $alias]" -ForegroundColor Green
    }
}
# === MICROSOFT 365 GROUPS ===
$objs = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')"  -Property @("Id", "DisplayName", "ProxyAddresses","groupType","MailNickname")
$i_total = $objs.Count
$i_count = 0
ForEach ($obj in $objs) {
    $primary = $obj.ProxyAddresses | Where-Object {$_.StartsWith("SMTP")}
    Write-Host "$((++$i_count)) of $($i_total): $($obj.DisplayName) ($primary) " -NoNewline
    $group = $obj
    $alias = "$($group.MailNickname)@$NewDomain".ToLower()
    $ProxyWithPrimary, $ProxyStatus = MakePrimaryinProxy $alias $obj.ProxyAddresses
    if ($ProxyStatus -eq "primary") {
        Write-Host " [$($alias) Alias already primary]" -ForegroundColor Green
    } else {
        if (-not $makechanges) {
            $wouldclause = "Would"
        } else {
            $wouldclause = ""
        }   
        if ($ProxyStatus -eq "nonprimary") {
            Write-Host " [$($alias) $($wouldclause) Promote existing alias to primary]" -ForegroundColor Yellow
        } elseif ($ProxyStatus -eq "notfound") {
            Write-Host " [$($alias) $($wouldclause) Add new primary alias]" -ForegroundColor Yellow
        }
        if ($makechanges) {
            Set-UnifiedGroup -Identity $group.MailNickname -PrimarySmtpAddress $alias
        }     
    }
}
# === MAIL-ENABLED SECURITY GROUPS / DISTRIBUTION LISTS ===
$objs = Get-DistributionGroup
$i_total = $objs.Count
$i_count = 0
ForEach ($obj in $objs) {
    Write-Host "$((++$i_count)) of $($i_total): $($obj.DisplayName) ($($obj.PrimarySmtpAddress)) " -NoNewline
    $group = $obj
    $alias = $group.PrimarySmtpAddress.Split("@")[0]
    $alias += "@$NewDomain"
    $ProxyWithPrimary, $ProxyStatus = MakePrimaryinProxy $alias $obj.EmailAddresses
    if ($ProxyStatus -eq "primary") {
        Write-Host " [$($alias) Alias already primary]" -ForegroundColor Green
    } else {
        if (-not $makechanges) {
            $wouldclause = "Would"
        } else {
            $wouldclause = ""
        }   
        if ($ProxyStatus -eq "nonprimary") {
            Write-Host " [$($alias) $($wouldclause) Promote existing alias to primary]" -ForegroundColor Yellow
        } elseif ($ProxyStatus -eq "notfound") {
            Write-Host " [$($alias) $($wouldclause) Add new primary alias]" -ForegroundColor Yellow
        }
        if ($makechanges) {
            Set-DistributionGroup -Identity $group.id -EmailAddresses $ProxyWithPrimary
        }
    }
}
Write-Host "Done"
PressEnterToContinue
