<#
    .SYNOPSIS
        This script returns the Microsoft Graph API reports on lastactivitydetails of users for Office 365 services like ExchangeOnline, SharePointOnline, OneDrive for Business etc.
        The reporting is made based on a Native Application registered in Azure AD. Please follow the article https://blogs.technet.microsoft.com/dawiese/2017/04/15/get-office365-usage-reports-from-the-microsoft-graph-using-windows-powershell/
        to dig deeper.

    .DESCRIPTION
        Using an native App registered in Azure AD along with a valid O365 Administrator Account, Invoke a Graph API Reporting Endpoint and returns the report in a .CSV file extension.
        For more details refer: 
            > https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report 
            > https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/reportroot_getoffice365activeuserdetail

    .EXAMPLE
        PS C:\Users\MyPOSH> .\lastActivityStats.ps1

    .NOTES
        Author:         Noble K Varghese
        Version:        3.0.1
            Creation Date:  15-May-2018
            Purpose/Change: Reference to the article https://www.petri.com/get-mailboxstatistics-cmdlet-wrong, Last Login Date Reported by the Get-MailboxStatistics Cmdlet was not accurate.
                            Re-designed the script to use Microsoft GRAPH API to return lastActivityDate of Users.
        
        Version:        3.2
            Creation Date:  12-April-2019
            Purpose/Change: Reference to the issue reported https://github.com/noblevarghese/Office365-Users-LastActivityDetails/issues/1, redesigned the script to use OAuth & ADAL based Modern Authentication.
                            Earlier the script was using Basic Authentication using Get-Credential

    .LINK
    https://blogs.technet.microsoft.com/dawiese/2017/04/15/get-office365-usage-reports-from-the-microsoft-graph-using-windows-powershell/

#>

####################################################

#region AuthToken
function Get-AuthToken {
    <#
    .SYNOPSIS
    This function is used to authenticate with the Graph API REST interface
    .DESCRIPTION
    The function authenticate with the Graph API Interface with the tenant name
    .EXAMPLE
    Get-AuthToken
    Authenticates you with the Graph API interface
    .NOTES
    NAME: Get-AuthToken
    #>
        [cmdletbinding()]
        param
        (
            [Parameter(Mandatory=$true)]
            $User
        )
    
        $userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User
        $tenant = $userUpn.Host
    
        Write-Host "Checking for MSOnline module..."
        $AadModule = Get-Module -Name "MSOnline" -ListAvailable
    
        if ($AadModule -eq $null) {
            write-host
            write-host "MSOnline Powershell module not installed..." -f Red
            write-host "Install by running 'Install-Module MSOnline' or 'Install-Module MSOnline' from an elevated PowerShell prompt" -f Yellow
            write-host "Script can't continue..." -f Red
            write-host
            exit
        }
        # Getting path to ActiveDirectory Assemblies
        # If the module count is greater than 1 find the latest version
    
        if($AadModule.count -gt 1){
            $Latest_Version = ($AadModule | select version | Sort-Object)[-1]
            $aadModule = $AadModule | ? { $_.version -eq $Latest_Version.version }
            # Checking if there are multiple versions of the same module found
    
            if($AadModule.count -gt 1){
                $aadModule = $AadModule | select -Unique
            }
    
            $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
            $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
        else {
            $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
            $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
    
        [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
        [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    
        $clientId = "c2a5cdb6-3f67-4aa8-8c8f-672a65608721"
        $redirectUri = "urn:myApp"
        $resourceAppIdURI = "https://graph.microsoft.com"
        $authority = "https://login.microsoftonline.com/$Tenant"
    
        try {
            $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
            # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
            # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
            $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
            $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
            $MethodArguments = [Type[]]@("System.String", "System.String", "System.Uri", "Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior", "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier")
            $NonAsync = $AuthContext.GetType().GetMethod("AcquireToken", $MethodArguments)
    
            if ($NonAsync -ne $null){
                $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId, [Uri]$redirectUri, [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto, $userId)
            }
            else {
                $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId, [Uri]$redirectUri, $platformParameters, $userId).Result 
                }
            # If the accesstoken is valid then create the authentication header
            
            if($authResult.AccessToken){
                # Creating header for Authorization token
                $authHeader = @{
                    'Content-Type'='application/json'
                    'Authorization'="Bearer " + $authResult.AccessToken
                    'ExpiresOn'=$authResult.ExpiresOn
                    }
                return $authHeader
            }
            else {
                Write-Host
                Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
                Write-Host
                break
            }
        }
        catch {
            write-host $_.Exception.Message -f Red
            write-host $_.Exception.ItemName -f Red
            write-host
            break
        }
    }
#endregion
    
####################################################
    
#region Authentication
write-host
# Checking if authToken exists before running authentication

if($global:authToken){
    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()
    # If the authToken exists checking when it expires
    $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

    if($TokenExpires -le 0){
        write-host "Authentication Token expired" $TokenExpires "minutes ago" -ForegroundColor Yellow
        write-host
        # Defining User Principal Name if not present
        
        if($User -eq $null -or $User -eq ""){
            $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
            Write-Host
            }
        $global:authToken = Get-AuthToken -User $User
    }
}
# Authentication doesn't exist, calling Get-AuthToken function
else {
    if($User -eq $null -or $User -eq ""){
    $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
    Write-Host
    }

    # Getting the authorization token
    $global:authToken = Get-AuthToken -User $User
}
#endregion

####################################################

#region GRAPH Call

$uri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')"
$result = @(Invoke-RestMethod -Uri $uri -Headers $global:authToken -Method Get)
$report = ConvertFrom-Csv -InputObject $result

$report | Export-Csv -Path "LastActivityStats_$((Get-Date -uformat %Y%m%d%H%M%S).ToString()).csv" -NoTypeInformation -Encoding UTF8

#endregion

####################################################