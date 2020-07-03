﻿function Get-M365Messages {
	<#
	.SYNOPSIS
		Retrieve messages from the Message Center.
	
	.DESCRIPTION
		Retrieve messages from the Message Center.
	
	.PARAMETER TenantID
		ID of the tenant to manage.
	
	.PARAMETER ClientID
		Client ID (or Application ID) of the application set up for the authentication workflow.
	
	.PARAMETER ClientSecret
		Client secret of the application to use for the authentication workflow
	
	.EXAMPLE
		PS C:\> Get-MmMessage -TenantID $tenant -ClientID $clientID -ClientSecret $secret
	
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
		$ClientSecret
	)
	
	begin {
		$loginURL = "https://login.windows.net/"
	}
	process {
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
		$messageCenterMessages = Invoke-RestMethod -Method GET -Headers $header -Uri "https://manage.office.com/api/v1.0/$TenantID/ServiceComms/Messages"
		
		foreach ($FeatureMessage in $messageCenterMessages.value) {
			if (($FeatureMessage.LastUpdatedTime -as [DateTime]).AddDays(1) -lt (Get-Date)) { continue }
			foreach ($MessageHistory in $FeatureMessage.Messages) {
				[pscustomobject][ordered]@{
					Computer                = $env:COMPUTERNAME
					O365MsgId               = $FeatureMessage.Id
					O365DefaultId           = $domain
					O365Tenant              = $TenantID
					O365WorkloadId          = $FeatureMessage.Workload
					O365WorkloadDisplayName = $FeatureMessage.WorkloadDisplayName
					O365FeatureDisplayName  = $FeatureMessage.FeatureDisplayName
					O365MsgStatus           = $FeatureMessage.Status
					O365MsgSeverity         = $FeatureMessage.Severity
					O365MsgType             = $FeatureMessage.MessageType
					O365MsgClass            = $FeatureMessage.Classification
					O365MsgTitle            = $FeatureMessage.Title
					O365MsgImpactDesc       = $FeatureMessage.ImpactDescription
					O365MsgUpdatedTime      = $FeatureMessage.LastUpdatedTime
					O365MsgHistPubTime      = $MessageHistory.PublishedTime
					O365MessageHistText     = $MessageHistory.MessageText
				}
			}
		}
	}
}