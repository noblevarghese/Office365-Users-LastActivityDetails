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
        lastActivityStats.ps1 -TenantName "contoso.onmicrosoft.com" -clientId "74f0e6c8-0a8e-4a9c-9e0e-4c8223013eb9" -redirecturi "urn:ietf:wg:oauth:2.0:oob"

    .PARAMETER TentantName
        Tenant name in the format <tenantname>.onmicrosoft.com

    .PARAMETER clientID
        The clientID or AppID of the native app created in AzureAD to grant access to the reporting API. This is the application ID of the App registered in Azure AD.

    .Parameter redirecturi
        The replyURL of the native app created in AzureAD to grant access to the reporting API. This is the redirectURI of the App registered in Azure AD.

    .Parameter resourceAppIDURI
        Protocol and Hostname for the endpoint you are accessing. For the Graph API enter "https://graph.microsoft.com" This is hardcoded in the script.
        Hence you needn't pass it while running the script.

    .NOTES
        Author:         Noble K Varghese
        Version:        3.0.1
            Creation Date:  15-May-2018
            Purpose/Change: Reference to the article https://www.petri.com/get-mailboxstatistics-cmdlet-wrong, Last Login Date Reported by the Get-MailboxStatistics Cmdlet was not accurate.
                            Re-designed the script to use Microsoft GRAPH API to return lastActivityDate of Users.

    .LINK
    https://blogs.technet.microsoft.com/dawiese/2017/04/15/get-office365-usage-reports-from-the-microsoft-graph-using-windows-powershell/

#>

[cmdletbinding()]
param (
    [Parameter(Mandatory=$true)]
    $TenantName,

    [Parameter(Mandatory=$true)]
    $clientId,

    [Parameter(Mandatory=$true)]
    $redirecturi,

    [Parameter(Mandatory=$false)]
    $resourceAppIdURI
)

####################################################

function Get-AuthToken ([string]$TenantName, [string]$clientId, [string]$redirecturi,[string]$resourceAppIdURI,[System.Management.Automation.PSCredential]$Credential) {

    <#
        .SYNOPSIS
        Gets an OAuth token for use with the Microsoft Graph API
    
        .DESCRIPTION
        Gets an OAuth token for use with the Microsoft Graph API

        .EXAMPLE
        Get-AuthToken -TenantName "contoso.onmicrosoft.com" -clientId "74f0e6c8-0a8e-4a9c-9e0e-4c8223013eb9" -redirecturi "urn:ietf:wg:oauth:2.0:oob" -resourceAppIdURI "https://graph.microsoft.com"
    
        .PARAMETER TentantName
        Tenant name in the format <tenantname>.onmicrosoft.com

        .PARAMETER clientID
        The clientID or AppID of the native app created in AzureAD to grant access to the reporting API

        .Parameter redirecturi
        The replyURL of the native app created in AzureAD to grant access to the reporting API

        .Parameter resourceAppIDURI
        protocol and hostname for the endpoint you are accessing. For the Graph API enter "https://graph.microsoft.com"
    
        .NOTES
        Inital authentication sample from:
        https://blogs.technet.microsoft.com/paulomarques/2016/03/21/working-with-azure-active-directory-graph-api-from-powershell/

    #>
    
    #Import the MSOnline module so we can lookup the directory for Microsoft.IdentityModel.Clients.ActiveDirectory.dll and Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll
    #MSOnline module documentation: https://www.powershellgallery.com/packages/MSOnline/1.1.166.0
    Try {

        Write-Debug "Importing MSONline Module for ADAL assemblies"
        Import-Module MSOnline -ErrorAction Stop
    }
    Catch [System.IO.FileNotFoundException] {

        Write-Warning "The module MSOnline is not installed.`nPlease run Install-Module MSOnline from an elevated window to install it from the PowerShell Gallery"
        Throw "MSOnline module not installed"
    }
    #Get the module folder so we can load the DLLs we want
    $modulebase = (Get-Module MSONline | Sort Version -Descending | Select -First 1).ModuleBase
    $adal = "{0}\Microsoft.IdentityModel.Clients.ActiveDirectory.dll" -f $modulebase
    $adalforms = "{0}\Microsoft.IdentityModel.Clients.ActiveDirectory.WindowsForms.dll" -f $modulebase

    #Attempt to load the assemblies. Without these we cannot continue so we need the user to stop and take an action
    Try {

        [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
        [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    }
    Catch {

        #MSOnline Version 1.0 does not contain the DLLs that we need, a minimum version of 1.1.166.0 is required
        Write-Warning "Unable to load ADAL assemblies.`nUpdate the MSOnline module by running Install-Module MSOnline -Force -AllowClobber"
        Throw $error[0]
    }

    #Build the logon URL with the tenant name
    $authority = "https://login.windows.net/$TenantName"
    Write-Verbose "Logon Authority: $authority"

    #Build the auth context and get the result
    Write-Verbose "Creating AuthContext"
    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    Write-Verbose "Creating AD UserCredential Object"
    #$Credential = Get-Credential -Credential "admin@M365x615832.onmicrosoft.com"
    $AdUserCred = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential" -ArgumentList $Credential.username, $Credential.Password
    Try {

        Write-Verbose "Attempting passive authentication"
        $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$AdUserCred)
    }
    Catch [System.Management.Automation.MethodInvocationException] {

        #The first that the the user runs this, they must open an interactive window to grant permissions to the app
        If ($error[0].Exception.Message -like "*Send an interactive authorization request for this user and resource*") {

                Write-Warning "The app has not been granted permissions by the user. Opening an interactive prompt to grant permissions"
                $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId,$redirectUri, "Always") #Always prompt for user credentials so we don't use Windows Integrated Auth
        }
        Else {
            
            Throw
        }
    }
    

    #Return the authentication token
    return $authResult
}

####################################################

#region Authentication

if($global:token) {

    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()

    # If the authToken exists checking when it expires
    $TokenExpires = ($token.ExpiresOn.datetime - $DateTime).Minutes

    if($TokenExpires -le 0) {

        write-host "Authentication Token expired" $TokenExpires "minutes ago" -ForegroundColor Yellow
        write-host

        # Defining User Principal Name if not present

        if($Cred -eq $null -or $Cred -eq "") {

            $Cred = Get-Credential -Credential "user@domain.onmicrosoft.com"
            Write-Host
        }

        $global:token = Get-AuthToken -TenantName $TenantName -clientId $clientId -redirecturi $redirecturi -resourceAppIdURI "https://graph.microsoft.com" -Credential $cred

    }
}
    
# Authentication doesn't exist, calling Get-AuthToken function

else {

    if($Cred -eq $null -or $Cred -eq ""){

        $Cred = Get-Credential -Credential "user@domain.onmicrosoft.com"
        Write-Host
    }
    #Getting the authorization token
    $global:token = Get-AuthToken Get-AuthToken -TenantName $TenantName -clientId $clientId -redirecturi $redirecturi -resourceAppIdURI "https://graph.microsoft.com" -Credential $cred

}


#Build REST API header with authorization token
$authHeader = @{

    'Content-Type'='application\json'
    'Authorization'=$token.CreateAuthorizationHeader()
}

#endregion

####################################################

#region GRAPH Call

$uri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')"
$result = @(Invoke-RestMethod -Uri $uri -Headers $authHeader -Method Get)
$report = ConvertFrom-Csv -InputObject $result

$report | Export-Csv -Path "LastActivityStats_$((Get-Date -uformat %Y%m%d%H%M%S).ToString()).csv" -NoTypeInformation -Encoding UTF8

#endregion

####################################################