function Get-M365ServiceHealth {
	<#
	.SYNOPSIS
		Retrieve M365 Service Health from O365 Service Communications API.
	
	.DESCRIPTION
		Retrieve Query M365 Service health from )365 Service Communications API from the Message Center.
	
	.PARAMETER TenantID
		ID of the tenant to manage.
	
	.PARAMETER ClientID
		Client ID (or Application ID) of the application set up for the authentication workflow.
	
	.PARAMETER ClientSecret
		Client secret of the application to use for the authentication workflow
	
	.EXAMPLE
		PS C:\> Get-M365ServiceHealth -TenantID $tenant -ClientID $clientID -ClientSecret $secret
	
		Retrieves messages for the tenant stored in $tenant
	
	.EXAMPLE
		PS C:\> Import-Csv .\tenants.csv | Get-MmMessage
	
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

		[Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $false)]
		[string]
		$Tenant
	)
	
	begin {
		$loginURL = "https://login.windows.net/"
	
		Get-Process {
			$tenantLoginURL = $loginURL + $TenantID
		
			$oauth = New-OAuthToken -Uri "$tenantLoginURL/oauth2/token" -Resource 'https://graph.microsoft.com/' -ClientID $ClientID -ClientSecret $ClientSecret
			$headerParams = @{ 'Authorization' = "$($oauth.token_type) $($oauth.access_token)" }
		
			$paramInvokeRestMethod = @{
				Headers   = $headerParams
				Uri       = 'https://graph.microsoft.com/v1.0/domains'
				UserAgent = 'application/json'
				Method    = 'Get'
			}
			$domain = ((Invoke-RestMethod @paramInvokeRestMethod).value | Where-Object isDefault).id
		
			$oauth = New-OAuthToken -Uri "$tenantLoginURL/oauth2/token" -Resource 'https://manage.office.com/' -ClientID $ClientID -ClientSecret $ClientSecret
			$header = @{ "Authorization" = "Bearer $($oauth.access_token)"; "Content-Type" = "application/json" }
		
			#--- Get the data ---#
			$healthData = Invoke-RestMethod -Method GET -Headers $header -Uri "https://manage.office.com/api/v1.0/$TenantID/ServiceComms/CurrentStatus"
		
			Foreach ($O365Workload in $healthData.value ) {
				Foreach ($O365Feature in $O365Workload.FeatureStatus ) {
					[PSCustomObject][ordered]@{
						Computer                = $env:COMPUTERNAME
						O365StatusTime          = $O365Workload.StatusTime
						O365TenantName          = $Tenant
						O365DefaultId           = $domain
						O365TenantID            = $TenantID
						O365WorkloadId          = $O365Workload.Id
						O365WorkloadDisplayName = $O365Workload.WorkloadDisplayName
						O365WorkloadStatus      = $O365Workload.Status
						O365WorkloadIncidentID  = $O365Workload.IncidentIds -join ","
						O365FeatureName         = $O365Feature.FeatureDisplayName
						O365FeatureStatus       = $O365Feature.FeatureServiceStatus
						O365ImpactDesc          = $O365Feature.ImpactDescription
					}
				
				}
			}
		}
	}
}
