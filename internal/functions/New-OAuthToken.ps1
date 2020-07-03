function New-OAuthToken
{
	[CmdletBinding()]
	param (
		[string]
		$Uri,
		
		[string]
		$Resource,
		
		[string]
		$ClientID,
		
		[string]
		$ClientSecret
	)
	
	process
	{
		$body = @{
			grant_type    = "client_credentials"
			resource	  = $Resource
			client_id	  = $ClientID
			client_secret = $ClientSecret
		}
		Invoke-RestMethod -Method Post -Uri $Uri -Body $body
	}
}