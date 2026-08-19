# Status code to readable name mapping
$statusMap = @{
    "SOURCE_STATE_HEALTHY" = "Healthy"
    "SOURCE_STATE_UNCHECKED_SOURCE_NO_ACCOUNTS" = "Config Incomplete"
    # Add more as needed
}

# Connector name to readable name mapping
# For the Other types
$connectorMap = @{
    "delimited-file-angularsc" = "Delimited File"
    "salesforce-saas" = "Salesforce"
    "oktasaas" = "Okta"
    "cyberarkpcloudsharedservices-saas" = "CyberArk"
    # Add more as needed
}

function Get-AD-Source-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        #Filtering on type "Active Directory - Direct"
		# $adSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources?filters=type%20eq%20%22Active%20Directory%20-%20Direct%22" -Method 'GET' -Headers $headers
		$adSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources?filters=type%20eq%20%22Active%20Directory%20-%20Direct%22" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($source in $adSourceResponse) {

            $forestAuth = "Not Configured"
            $domainAuth = "Not Configured"

            if ($source.connectorAttributes.forestSettings) {
                if ($source.connectorAttributes.forestSettings.user -and  $null -ne $source.connectorAttributes.forestSettings.user) {
                    $forestAuth = "Basic"
                    if ($source.connectorAttributes.forestSettings.useSSL){
                        $forestAuth = "SSL"
                    }
                    if ($source.connectorAttributes.forestSettings.enablePasswordLessAuthenticationForForest) {
                        $forestAuth = "gMSA"
                    }
                }
            }

            if ($source.connectorAttributes.domainSettings) {
                if ($source.connectorAttributes.domainSettings.user -and  $null -ne $source.connectorAttributes.domainSettings.user) {
                    $domainAuth = "Basic"
                    if ($source.connectorAttributes.domainSettings.useSSL){
                        $domainAuth = "SSL"
                    }
                    if ($source.connectorAttributes.domainSettings.enablePasswordLessAuthenticationForForest) {
                        $domainAuth = "gMSA"
                    }
                }
            }

            if ($domainAuth -eq $forestAuth ) { $sourceAuth = $domainAuth}
            elseif ($forestAuth -eq "Not Configured") {  $sourceAuth = $domainAuth }
            elseif ($domainAuth -eq "Not Configured") {  $sourceAuth = $forestAuth }
            else {  $sourceAuth = $forestAuth + "/" + $domainAuth }

            $results += [PSCustomObject]@{
                "Source Name" = $source.name
                "Owner"       = $source.owner.name
                "Auth"        = $sourceAuth
                "Status"      = $statusMap[$source.status]
            }
        }

        # If no results, add a default "no data" object
        if ($results.Count -eq 0) {
            $results += [PSCustomObject]@{
                "Source Name" = "--NO DATA--"
                "Owner"       = "--NO DATA--"
                "Auth"        = "--NO DATA--"
                "Status"      = "--NO DATA--"
            }
        }

		return $results | Sort-Object "Source Name"

    } catch {
        Write-Error "Failed to retrieve AD Source Data: $_"
        return $null
    }
}
function Get-IQ-Service-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        #Filtering on type "Active Directory - Direct"
		# $iqServiceResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources?filters=type%20eq%20%22Active%20Directory%20-%20Direct%22" -Method 'GET' -Headers $headers
		$iqServiceResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources?filters=type%20eq%20%22Active%20Directory%20-%20Direct%22" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($source in $iqServiceResponse) {

            if ($source.connectorAttributes.IQServiceHost) {

                # $iqServiceVersionResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources/$($source.id)/source-health" -Method 'GET' -Headers $headers
                $iqServiceVersionResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources/$($source.id)/source-health" -Method 'GET' -Headers $headers

                $results += [PSCustomObject]@{
                    "Source Name"         = $source.name
                    "IQService Host:Port" = $source.connectorAttributes.IQServiceHost + ":" + $source.connectorAttributes.IQServicePort
                    "TLS"                 = $source.connectorAttributes.useTLSForIQService
                    "Version"             = $iqServiceVersionResponse.iqServiceVersion
                    "Status"              = $statusMap[$source.status]
                }
            }
        }

        # If no results, add a default "no data" object
        if ($results.Count -eq 0) {
            $results += [PSCustomObject]@{
                "Source Name"         = "--NO DATA--"
                "IQService Host:Port" = "--NO DATA--"
                "TLS"                 = "--NO DATA--"
                "Version"             = "--NO DATA--"
                "Status"              = "--NO DATA--"
            }
        }

		return $results | Sort-Object "Source Name"

    } catch {
        Write-Error "Failed to retrieve IQ Service Data: $_"
        return $null
    }
}
function Get-Entra-Source-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        #Entra mapped to type "2631a78f-e9b0-4a59-859f-f37f3e549d9f"
        #Filtering on type "2631a78f-e9b0-4a59-859f-f37f3e549d9f"
		# $entraSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources?filters=type%20eq%20%222631a78f-e9b0-4a59-859f-f37f3e549d9f%22%20or%20type%20eq%20%22Azure%20Active%20Directory%22" -Method 'GET' -Headers $headers
		$entraSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources?filters=type%20eq%20%222631a78f-e9b0-4a59-859f-f37f3e549d9f%22%20or%20type%20eq%20%22Azure%20Active%20Directory%22" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($source in $entraSourceResponse) {


            $results += [PSCustomObject]@{
                "Source Name" = $source.name
                "Owner"       = $source.owner.name
                "Auth"        = $source.connectorAttributes.grantType
                "Status"      = $statusMap[$source.status]
            }
        }

        # If no results, add a default "no data" object
        if ($results.Count -eq 0) {
            $results += [PSCustomObject]@{
                "Source Name" = "--NO DATA--"
                "Owner"       = "--NO DATA--"
                "Auth"        = "--NO DATA--"
                "Status"      = "--NO DATA--"
            }
        }

		return $results | Sort-Object "Source Name"

    } catch {
        Write-Error "Failed to retrieve Entra Source Data: $_"
        return $null
    }
}
function Get-JDBC-Source-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        #Filtering on type "JDBC"
		# $jdbcSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources?filters=type%20eq%20%22JDBC%22" -Method 'GET' -Headers $headers
		$jdbcSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources?filters=type%20eq%20%22JDBC%22" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($source in $jdbcSourceResponse) {


            $results += [PSCustomObject]@{
                "Source Name"         = $source.name
                "Owner"               = $source.owner.name
                "Connector files"     = $source.connectorAttributes.connector_files
                "Credential Provider" = $source.credentialProviderEnabled
                "Status"              = $statusMap[$source.status]
            }
        }

        # If no results, add a default "no data" object
        if ($results.Count -eq 0) {
            $results += [PSCustomObject]@{
                "Source Name"         = "--NO DATA--"
                "Owner"               = "--NO DATA--"
                "Connector files"     = "--NO DATA--"
                "Credential Provider" = "--NO DATA--"
                "Status"              = "--NO DATA--"
            }
        }

		return $results | Sort-Object "Source Name"

    } catch {
        Write-Error "Failed to retrieve JDBC Source Data: $_"
        return $null
    }
}
function Get-WebService-Source-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        #Filtering on type "Web Services"
		# $webSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources?filters=type%20eq%20%22Web%20Services%22" -Method 'GET' -Headers $headers
		$webSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources?filters=type%20eq%20%22Web%20Services%22" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($source in $webSourceResponse) {


            $results += [PSCustomObject]@{
                "Source Name"         = $source.name
                "Owner"               = $source.owner.name
                "Auth"                = $source.connectorAttributes.authenticationMethod
                "Credential Provider" = $source.credentialProviderEnabled
                "Status"              = $statusMap[$source.status]
            }
        }

        # If no results, add a default "no data" object
        if ($results.Count -eq 0) {
            $results += [PSCustomObject]@{
                "Source Name"         = "--NO DATA--"
                "Owner"               = "--NO DATA--"
                "Auth"                = "--NO DATA--"
                "Credential Provider" = "--NO DATA--"
                "Status"              = "--NO DATA--"
            }
        }

		return $results | Sort-Object "Source Name"

    } catch {
        Write-Error "Failed to retrieve WebService Source Data: $_"
        return $null
    }
}
function Get-Delimited-Source-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        #Filtering on type "Web Services"
		# $delimitedSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources?filters=type%20eq%20%22DelimitedFile%22" -Method 'GET' -Headers $headers
		$delimitedSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources?filters=type%20eq%20%22DelimitedFile%22" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($source in $delimitedSourceResponse) {
            if ($statusMap[$source.status] -eq "Healthy") {
                $colourCode = $script:colourCodeGreen
            } elseif ($statusMap[$source.status] -eq "Config Incomplete") {
                $colourCode = $script:colourCodeYellow
            } else {
                $colourCode = $script:colourCodeTransparent
            }

            $results += [PSCustomObject]@{
                "Source Name" = $source.name
                "Owner"       = $source.owner.name
                "Status"      = $statusMap[$source.status]
                "Colour"      = $colourCode
            }
        }

        # If no results, add a default "no data" object
        if ($results.Count -eq 0) {
            $results += [PSCustomObject]@{
                "Source Name" = "--NO DATA--"
                "Owner"       = "--NO DATA--"
                "Status"      = "--NO DATA--"
            }
        }

		return $results | Sort-Object "Source Name"

    } catch {
        Write-Error "Failed to retrieve WebService Source Data: $_"
        return $null
    }
}
function Get-Other-Source-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        #Entra mapped to type "2631a78f-e9b0-4a59-859f-f37f3e549d9f"
        #Filtering on type ne "Active Directory - Direct" and type ne "Microsoft-Entra" and type ne "JDBC" and type ne "Web Services" and type ne "DelimitedFile"
		# $otherSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources?filters=type%20ne%20%22Active%20Directory%20-%20Direct%22%20and%20type%20ne%20%222631a78f-e9b0-4a59-859f-f37f3e549d9f%22%20and%20type%20ne%20%22JDBC%22%20and%20type%20ne%20%22Web%20Services%22%20and%20type%20ne%20%22DelimitedFile%22and%20type%20ne%20%22Azure%20Active%20Directory%22" -Method 'GET' -Headers $headers
		$otherSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources?filters=type%20ne%20%22Active%20Directory%20-%20Direct%22%20and%20type%20ne%20%222631a78f-e9b0-4a59-859f-f37f3e549d9f%22%20and%20type%20ne%20%22JDBC%22%20and%20type%20ne%20%22Web%20Services%22%20and%20type%20ne%20%22DelimitedFile%22and%20type%20ne%20%22Azure%20Active%20Directory%22" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($source in $otherSourceResponse) {
            switch ($statusMap[$source.status]) {
                "Healthy" {
                    $colourCode = $script:colourCodeGreen
                    break
                }
                "Config Incomplete" {
                    $colourCode = $script:colourCodeYellow
                    break
                }
                default {
                    $colourCode = $script:colourCodeTransparent
                }
            }

            $results += [PSCustomObject]@{
                "Source Name" = $source.name
                "Type"        = $connectorMap[$source.connector]
                "Owner"       = $source.owner.name
                "Auth"        = "--MANUAL--"
                "Status"      = $statusMap[$source.status]
                "Colour"      = $colourCode
            }
        }

        # ----------- SPLIT DATA INTO DUPLICATE TYPES AND OTHERS -----------
        # grouping results by Type
        $grouped = $results | Group-Object Type

        # Duplicate Types (>= 3 results)
        $duplicateTypeGroups = $grouped | Where-Object { 
            $_.Count -ge 2 -and 
            -not [string]::IsNullOrWhiteSpace($_.Name)
        }

        # Misc = everything else
        $miscSources = $grouped |
            Where-Object { 
                $_.Count -lt 3 -or
                [string]::IsNullOrWhiteSpace($_.Name)
            } |
            ForEach-Object { $_.Group } |
            Sort-Object "Source Name"

        # If no results, add a default "no data" object
        # if ($results.Count -eq 0) {
        if ($miscSources.Count -eq 0) {
            $results += [PSCustomObject]@{
                "Source Name" = "--NO DATA--"
                "Type"        = "--NO DATA--"
                "Owner"       = "--NO DATA--"
                "Auth"        = "--NO DATA--"
                "Status"      = "--NO DATA--"
            }
        }

		# return $results | Sort-Object "Source Name"
		return [PSCustomObject]@{
            DuplicateTypeGroups = $duplicateTypeGroups
            MiscSources = $miscSources
        }

    } catch {
        Write-Error "Failed to retrieve Other Source Data: $_"
        return $null
    }
}

