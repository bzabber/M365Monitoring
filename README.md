
# M365 Monitoring

The M365Monitoring module has two cmdlets:
Get-M365ServiceHealth is used to query the M365 Service Health API (O365 Service Communications API). The output can be then pipelined to something else or simple used dump the status to the console. You must provide the M365 TenantID, ClientID (AppID) and Client Secret.

Get-M365Messages.ps1 is used to query the M365 Message Center API found at https://manage.office.com/api/v1.0/contoso.com/ServiceComms/Messages. 
Get-M365ServiceHealth.ps1 is used to query the M365 Service Health API found at https://manage.office.com/api/v1.0/contoso.com/ServiceComms/CurrentStatus.

The M365Monitoring module must be located in your module path in a folder called M365Monitoring.

You must provide the M365 TenantID, ClientID (AppID) and Client Secret.

Use example:
get-M365Messages.ps1 -TenantID <TenantID GUID> -ClientID <App ID GUID> -ClientSecret <Client Secret>
get-M365ServiceHealth.ps1 -TenantID <TenantID GUID> -ClientID <App ID GUID> -ClientSecret <Client Secret>
