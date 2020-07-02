#--- Login URL used by each tenant to create a OAuth2 token ---#
$loginURL = "https://login.windows.net/"
###############################################################
#- Array of tenants and the required parameters to authenticate with M365 tenant -#
#- You need to add a new array for each tenant that will be queried.             -#
#- Each Hashtable value will reference associated value in Azure Automation      -#
#- Variables.                                                                    -#
#- Each Azure Variable will use the following naming convention:                 -#
#- Dept Abbreviation<VariableName>                                               -#
#- Ex: SSCTenantID                                                               -#
#-     SSCClientID                                                               -#
#-     SSCClientSecret                                                           -#
###################################################################################
$tenants = @(
    @{
        "TenantID"     = "<ENTER TENANTID HERE>";
        "ClientID"     = "<ENTER CLIENTID HERE>";
        "ClientSecret" = "<ENTER CLIENT SECRET HERE>"
    }<#,
    @{
        "TenantID"     = "<ENTER 2nd TENANTID HERE>";
        "ClientID"     = "<ENTER 2nd CLIENTID HERE>";
        "ClientSecret" = "<ENTER 2nd CLIENT SECRET HERE>";
    }#>
)
foreach ($tenant in $tenants.GetEnumerator()) {
    #generate tenant specfic login URL
    $tenantLoginURL = $loginURL + $tenant.TenantID

    #-- Get an Oauth 2 access token based on client id, secret and tenant id ---#
    $body = @{grant_type = "client_credentials"; resource = 'https://graph.microsoft.com/'; client_id = $tenant.ClientID; client_secret = $tenant.ClientSecret }
    #$oauth = Invoke-RestMethod -Method Post -Uri $loginURL$tenant.TenantID/oauth2/token -Body $body
    $oauth = Invoke-RestMethod -Method Post -Uri $tenantLoginURL/oauth2/token -Body $body
    #--- Extract token and add to HTTP Header
    $headerParams = @{'Authorization' = "$($oauth.token_type) $($oauth.access_token)" }

    #$domain = (Invoke-RestMethod -Header $headerParams -Method GET -URI 'https://graph.microsoft.com/v1.0/domains' -UserAgent 'application/json')
    $domain = (Invoke-RestMethod -Header $headerParams -URI 'https://graph.microsoft.com/v1.0/domains' -UserAgent 'application/json' -Method Get).value `
    | Where-Object isDefault -eq $true `
    | Select-Object -expandproperty id
     
    #--Get M365 SHD data for each tenant. --#
    #-- Get an Oauth 2 access token based on client id, secret and tenant domain ---#
    $body = @{grant_type = "client_credentials"; resource = 'https://manage.office.com/'; client_id = $tenant.ClientID; client_secret = $tenant.ClientSecret }
    $oauth = Invoke-RestMethod -Method Post -Uri $tenantLoginURL/oauth2/token -Body $body

    #--- Extract token and add to HTTP Header
    $accessToken = $oauth.access_token
    $header = @{"Authorization" = "Bearer $accessToken"; "Content-Type" = "application/json" };
    <# Commented this piece out to test MEssage center retrieval
    $healthData = Invoke-RestMethod -Method GET -Headers $header -Uri https://manage.office.com/api/v1.0/$tenant.TenantID/ServiceComms/CurrentStatus
    ##################################################################################################
    #--- This section will get details of current O365 service health state and upload to Log analytics
    ##################################################################################################
    # Specify the name of the record type that you'll be creating
    $O365TenantID = $tenant.TenantID
    $LogType = "O365ServiceHealth"

    #--- Start JSON ---#
    #$JSON = '['
    #$JSONCount = 0
    #--- Extract key values from each record and build JSON ---#
    $JSON = Foreach ($O365Workload in $healthData.value ) {
        Foreach ($O365Feature in $O365Workload.FeatureStatus ) {
            $O365WorkloadId = $O365Workload.Id
            $O365WorkloadDisplayName = $O365Workload.WorkloadDisplayName
            $O365FeatureName = $O365Feature.FeatureDisplayName
            $O365FeatureStatus = $O365Feature.FeatureServiceStatus
            $0365FeatureIncidents = $0365Feature.IncidentIds
            $O365ImpactDesc = $0365Feature.ImpactDescription   
            $O365StatusTime = $O365Workload.StatusTime
            ForEach-Object {
                @{
                    "O365DefaultId"           = "$domain"
                    "O365StatusTime"          = "$O365StatusTime"
                    "O365TenantID"            = "$O365TenantID"
                    "O365WorkloadId"          = "$O365WorkloadId"
                    "O365WorkloadDisplayName" = "$O365WorkloadDisplayName"
                    "Computer"                = "$env:ComputerName"
                    "O365FeatureName"         = "$O365FeatureName"
                    "O365FeatureStatus"       = "$O365FeatureStatus"
                    "0365FeatureIncidents"    = "$0365FeatureIncidents"
                    "O365ImpactDesc"          = "$O365ImpactDesc"
                } 
            } 
        } 
        
        
    }
    ConvertTo-Json -InputObject $JSON
    #$JSON
    #Send-OMSAPIIngestionFile -customerId $CustomerID -sharedKey $SharedKey -body $JSON -logType $LogType -Verbose 

}
Commented this piece out to test MEssage center retrieval#>
    ##################################################################################################
    #--- This section will get details of current O365 service messages and upload to Log analytics
    ##################################################################################################
    # Specify the name of the record type that you'll be creating
    $O365TenantID = $tenant.TenantID
    $LogType = "O365MessageCenter"

    #--- Start JSON ---#
    #$JSON = '['
    #$JSONCount = 0

    #--- Get the data ---#
    $messageCenterMessages = Invoke-RestMethod -Method GET -Headers $header -Uri https://manage.office.com/api/v1.0/$tenant.TenantID/ServiceComms/Messages
    #Export-CLixml -InputObject $messageCenterMessages -Path .\Messagesobject.clixml -Depth 5

    #--- Extract key values from each record and build JSON ---#
    $JSON = Foreach ($FeatureMessage in $messageCenterMessages.value) {
        $O365MsgUpdatedTime = $FeatureMessage.LastUpdatedTime
        $MsgHoursOld = (New-TimeSpan –Start $O365MsgUpdatedTime).totalhours

        if ( $MsgHoursOld -lt 24 ) {	
            $O365MsgId = $FeatureMessage.Id
            $O365WorkloadId = $FeatureMessage.Workload
            $O365WorkloadDisplayName = $FeatureMessage.WorkloadDisplayName
            $O365FeatureDisplayName = $FeatureMessage.FeatureDisplayName
            $O365MsgStatus = $FeatureMessage.Status
            $O365MsgSeverity = $FeatureMessage.Severity
            $O365MsgType = $FeatureMessage.MessageType
            $O365MsgClass = $FeatureMessage.Classification
            $O365MsgTitle = $FeatureMessage.Title
            $O365MsgImpactDesc = $FeatureMessage.ImpactDescription
            Foreach ($MessageHistory in $FeatureMessage.MessageHistory) {
                $O365MsgHistPubTime = $MessageHistory.PublishedTime
                $O365MsgHistText = $MessageHistory.MessageText
            
                Foreach-object {
                    @{
                        "Computer"                = "$env:COMPUTERNAME"
                        "O365MsgId"               = "$O365MsgId"
                        "O365DefaultId"           = "$domain"
                        "O365Tenant"              = "$O365TenantID"
                        "O365WorkloadId"          = "$O365WorkloadId"
                        "O365WorkloadDisplayName" = "$O365WorkloadDisplayName"
                        "O365FeatureDisplayName"  = "$O365FeatureDisplayName"
                        "O365MsgStatus"           = "$O365MsgStatus"
                        "O365MsgSeverity"         = "$O365MsgSeverity" 
                        "O365MsgType"             = "$O365MsgType"
                        "O365MsgClass"            = "$O365MsgClass"
                        "O365MsgTitle"            = "$O365MsgTitle"
                        "O365MsgImpactDesc"       = "$O365MsgImpactDesc"
                        "O365MsgUpdatedTime"      = "$O365MsgUpdatedTime"
                        "O365MsgHistPubTime"      = "$O365MsgHistPubTime"
                        "O365MessageHistText"     = "$O365MsgHistText"

                    }

                }
               
            }
          

            #write-host "Feature Message -> $O365MsgId,$O365WorkloadId,$O365WorkloadName,$O365FeatureName,$O365MsgTitle, $O365MsgStatus, $365MsgSeverity,$O365MsgType,$O365MsgClass,$O365MsgUpdatedTime,$O365MsgImpactDesc"

            $JSON
            #--- Submit the data to the API endpoint
            #Send-OMSAPIIngestionFile -customerId $CustomerID -sharedKey $SharedKey -body $JSON -logType $LogType -Verbose
        }#>
    }
          
}
   



exit


