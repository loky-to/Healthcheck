
function Get-VA-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        $limit = 50
        $offset = 0
        $allClusters = @()

        do {
            # $uri = "https://$script:tenant.api.identitynow-demo.com/v2024/managed-clusters"
            $uri = "https://$script:tenant.api.identitynow.com/v2024/managed-clusters"
            $response = Invoke-WebRequest -Uri $uri -Method 'GET' -Headers $headers

            $clusters = $response.Content | ConvertFrom-Json
            $totalCount = [int]$response.Headers["X-Total-Count"]

            foreach ($cluster in $clusters) {
                
                $allClusters += [PSCustomObject]@{
                    "Cluster Name"     = $cluster.name
                    "# of VAs"         = $cluster.clientIds.Count
                    "CCG Version"      = $cluster.ccgVersion
                    "Status"           = $cluster.consolidatedHealthIndicatorsStatus
                    "# of Connections" = $cluster.serviceCount
                    "Recommendations"  = ""
                }
            }
            $offset += $limit
        } while ($offset -lt $totalCount)

        # If no results, add a default "no data" object
        if ($allClusters.Count -eq 0) {
            $allClusters += [PSCustomObject]@{
                "Cluster Name"     = "--NO DATA--"
                "# of VAs"         = "--NO DATA--"
                "CCG Version"      = "--NO DATA--"
                "Status"           = "--NO DATA--"
                "# of Connections" = "--NO DATA--"
                "Recommendations"  = "--NO DATA--"
            }
        }

        return $allClusters | Sort-Object "Cluster Name"
    }
    catch {
        Write-Error "Failed to retrieve cluster details: $_"
        return $null
    }
}