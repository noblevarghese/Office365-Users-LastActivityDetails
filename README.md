# Office365 Users LastActivityDetails

This script returns the Microsoft Graph API reports on lastactivitydetails of users for Office 365 services like ExchangeOnline, SharePointOnline, OneDrive for Business etc.
The reporting is made based on a Native Application registered in Azure AD. Please follow the article https://blogs.technet.microsoft.com/dawiese/2017/04/15/get-office365-usage-reports-from-the-microsoft-graph-using-windows-powershell/ to dig deeper.

Using an native App registered in Azure AD along with a valid O365 Administrator Account, Invoke a Graph API Reporting Endpoint and returns the report in a .CSV file extension.
For more details refer:

    * https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/report
    * https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/reportroot_getoffice365activeuserdetail


> EXAMPLE

.\lastActivityStats.ps1

> NOTES
        
        Author:         Noble K Varghese
        Version:        3.0.1
            Creation Date:  15-May-2018
            Purpose/Change: Reference to the article https://www.petri.com/get-mailboxstatistics-cmdlet-wrong, Last Login Date Reported by the Get-MailboxStatistics Cmdlet was not accurate.
                            Re-designed the script to use Microsoft GRAPH API to return lastActivityDate of Users.
        
        Version:        3.2
            Creation Date:  12-April-2019
            Purpose/Change: Reference to the issue reported https://github.com/noblevarghese/Office365-Users-LastActivityDetails/issues/1, redesigned the script to use OAuth & ADAL based Modern Authentication.
                            Earlier the script was using Basic Authentication using Get-Credential

> Read More:
https://blogs.technet.microsoft.com/dawiese/2017/04/15/get-office365-usage-reports-from-the-microsoft-graph-using-windows-powershell/
