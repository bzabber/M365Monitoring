function Get-OrgSyncHealth {
    <#
	.SYNOPSIS
		Retrieve AAD Sync status at the Azure AD Orgnaization level via Microsoft Graph.
	
	.DESCRIPTION
		Retrieve AAD Sync Status at the Azure AD Organization level using Microsoft Graph.
	
	.PARAMETER TenantID
		ID of the tenant to manage.
	
	.PARAMETER ClientID
		Client ID (or Application ID) of the application set up for the authentication workflow.
	
	.PARAMETER ClientSecret
		Client secret of the application to use for the authentication workflow
	.PARAMETER TenantName
		Friendly name of the TenantID, user input determines this name. (ex: Contoso)
	
	.EXAMPLE
		PS C:\> Get-OrgSyncHealth -TenantID $tenant -ClientID $clientID -ClientSecret $secret -TenantName $TenantName
	
		Retrieves messages for the tenant stored in $tenant
	
	.EXAMPLE
		PS C:\> Import-Csv .\tenants.csv | Get-OrgSyncHealth
	
		Retrieves messages for all tenants stored in the tenants.csv
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $TenantID,
		
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias('ApplicationID', 'AppID')]
        [string]
        $ClientID,
		
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $ClientSecret,
		
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [string]
        $TenantName
    )
	
    begin {
        $loginURL = "https://login.windows.net/"
    }
    process {
        $tenantLoginURL = $loginURL + $TenantID
        #May need to add /v2.0 after oauth2 in URI below
        $oauth = New-OAuthToken -Uri "$tenantLoginURL/oauth2/token" -Resource 'https://graph.microsoft.com/' -ClientID $ClientID -ClientSecret $ClientSecret
        $headerParams = @{ 'Authorization' = "$($oauth.token_type) $($oauth.access_token)" }
		
        $paramInvokeRestMethod = @{
            Headers   = $headerParams
            Uri       = 'https://graph.microsoft.com/v1.0/domains'
            UserAgent = 'application/json'
            Method    = 'Get'
        }
        $domain = ((Invoke-RestMethod @paramInvokeRestMethod).value | Where-Object isDefault).id
		
        $oauth = New-OAuthToken -Uri "$tenantLoginURL/oauth2/token" -Resource 'https://graph.microsoft.com/' -ClientID $ClientID -ClientSecret $ClientSecret
        $header = @{ "Authorization" = "Bearer $($oauth.access_token)"; "Content-Type" = "application/json" }
		
        #--- Get the data ---#
        $OrgSyncHealth = Invoke-RestMethod -Method GET -Headers $header -Uri "https://graph.microsoft.com/beta/organization"
		
        foreach ($OrgSyncHealthState in $OrgSyncHealth.value) {
            #if (($FeatureMessage.LastUpdatedTime -as [DateTime]).AddDays(1) -lt (Get-Date)) { continue }
            #foreach ($MessageHistory in $FeatureMessage.Messages) {
            [pscustomobject][ordered]@{
                Computer                          = $env:COMPUTERNAME
                OrgOnPremSyncEnabled              = $OrgSyncHealthState.onPremisesSyncEnabled
                O365TenantName                    = $TenantName
                O365DefaultId                     = $domain
                OrgDisplayName                    = $OrgSyncHealthState.displayName
                O365Tenant                        = $TenantID
                OrgOnPremLastSyncDateTime         = $OrgSyncHealthState.onPremisesLastSyncDateTime
                OrgOnPremLastPasswordSyncDateTime = $OrgSyncHealthState.onPremisesLastPasswordSyncDateTime
            }
        }
    }
}
