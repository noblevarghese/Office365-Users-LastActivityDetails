# Office365 Users LastActivityDetails

This script returns the Microsoft Graph API reports on lastactivitydetails of users for Office 365 services like ExchangeOnline, SharePointOnline, OneDrive for Business etc.
The reporting is made based on a Native Application registered in Azure AD. Please follow the article https://blogs.technet.microsoft.com/dawiese/2017/04/15/get-office365-usage-reports-from-the-microsoft-graph-using-windows-powershell/ to dig deeper.

Using an native App registered in Azure AD along with a valid O365 Administrator Account, Invoke a Graph API Reporting Endpoint and returns the report in a .CSV file extension.
For more details refer: 
    * https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report [https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report]
    * https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/reportroot_getoffice365activeuserdetail [https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/reportroot_getoffice365activeuserdetail]

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

https://blogs.technet.microsoft.com/dawiese/2017/04/15/get-office365-usage-reports-from-the-microsoft-graph-using-windows-powershell/