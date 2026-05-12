function Get-DocCount {
    param (
        $buckets,
        [string]$key
    )
    $item = $buckets | Where-Object { $_.key -eq $key }
    if ($item) { return $item.doc_count } else { return 0 }
}

# Helper function to calculate average time
function Get-AverageTime {
    param ($tasks)
    $durations = @()
    foreach ($task in $tasks) {
        if ($task.completionStatus -eq "SUCCESS") {
            $launched = [datetime]::Parse($task.launched)
            $completed = [datetime]::Parse($task.completed)
            $duration = ($completed - $launched).TotalSeconds
            $durations += $duration
        }
    }
    if ($durations.Count -gt 0) {
        $avg = ($durations | Measure-Object -Average).Average
        $mins = [math]::Floor($avg / 60)
        $secs = [math]::Round($avg % 60)
        return "$mins mins, $secs secs"
    } else {
        return "N/A"
    }
}

function Get-OperationalData() {

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Content-Type", "application/json")
    $headers.Add("Authorization", "Bearer $accessToken")

    $daysInPast = "now-30d"

    try {

$jsonBody = @"
{
    "aggregationType": "SAILPOINT",
    "aggregations": {
        "bucket": {
            "field": "type"
        },
        "subAggregation": {
            "bucket": {
                "field": "technicalName"
            }
        }
    },
    "query": {
        "query": "status:FAILED AND created:[$daysInPast TO now]"
    },
    "indices": ["events"]
}
"@

        # $initialDataResponse = Invoke-RestMethod -Uri "https://$script:tenant.api.identitynow-demo.com/v3/search/aggregate" -Method Post -Headers $headers -Body $jsonBody
        $initialDataResponse = Invoke-RestMethod -Uri "https://$script:tenant.api.identitynow.com/v3/search/aggregate" -Method Post -Headers $headers -Body $jsonBody

        # Provisioning Summary
        $provisioningBuckets = $initialDataResponse.aggregations.type.buckets | Where-Object { $_.key -eq "provisioning" }
        $provTechBuckets = $provisioningBuckets.technicalName.buckets

        $accountProvSummaryData = [PSCustomObject]@{
            "AccModFail" = Get-DocCount $provTechBuckets "account_modify_failed"
            "AccCrtFail" = Get-DocCount $provTechBuckets "account_create_failed"
            "AccDisFail" = Get-DocCount $provTechBuckets "account_disable_failed"
            "AccEnaFail" = Get-DocCount $provTechBuckets "account_enable_failed"
        }

        # Access Item Summary
        $accessItemBuckets = $initialDataResponse.aggregations.type.buckets | Where-Object { $_.key -eq "access_item" }
        $accessTechBuckets = $accessItemBuckets.technicalName.buckets

        $accessProvSummaryData = [PSCustomObject]@{
            "EntAddFail" = Get-DocCount $accessTechBuckets "entitlement_add_failed"
            "EntRevFail" = Get-DocCount $accessTechBuckets "entitlement_remove_failed"
        }

        $summaryData = @()

        $buckets = $initialDataResponse.aggregations.type.buckets

        foreach ($category in $buckets) {
            $failure_type = $category.key
            foreach ($failure in $category.technicalName.buckets) {
                $technical_name = $failure.key
                $count30 = $failure.doc_count

                $summaryData += [PSCustomObject]@{
                    "failure type"     = $failure_type
                    "technical name"   = $technical_name
                    "last 30 days"     = $count30
                }
            }
        }

        $provisioningComboCounts = @{}
        $accessComboCounts = @{}
        $authComboCounts = @{}
        $aggComboCounts = @{}
        $passwordComboCounts = @{}
        $userPasswordComboCounts = @{}

        foreach ($failure_combo in $summaryData) {
            
            $limit = 250
            $offset = 0
            $allEvents = @()

            do {   
                $jsonBody = @"
{
    "query": {
        "query": "status:FAILED AND created:[$daysInPast TO now] AND type:$($failure_combo.'failure type') AND technicalName:$($failure_combo.'technical name')"
    },
    "indices": ["events"]
}
"@

                # $searchResponse = Invoke-WebRequest -Uri "https://$script:tenant.api.identitynow-demo.com/v3/search?offset=$offset&limit=$limit&count=true" -Method Post -Headers $headers -Body $jsonBody
                $searchResponse = Invoke-WebRequest -Uri "https://$script:tenant.api.identitynow.com/v3/search?offset=$offset&limit=$limit&count=true" -Method Post -Headers $headers -Body $jsonBody
                $data = $searchResponse.Content | ConvertFrom-Json

                $totalCount = [int]$searchResponse.Headers["X-Total-Count"]
                if ($totalCount -gt 10000) {$totalCount = 10000}

                $allEvents += $data
                $offset += $limit
            } while ($offset -lt $totalCount)

            # Analyze combinations
            switch ($failure_combo.'failure type') {
                "provisioning" {
                    foreach ($event in $allEvents) {
                        $provisioningComboKey = "$($event.attributes.sourceName)|$($event.attributes.interface)|$($event.technicalName)"
                        
                        if (-not $provisioningComboCounts.ContainsKey($provisioningComboKey)) {
                            $provisioningComboCounts[$provisioningComboKey] = @{ count = 0; error = $event.attributes.errors }
                        }
                        $provisioningComboCounts[$provisioningComboKey].count++
                    }
                    break
                }
                "access_item" {
                    foreach ($event in $allEvents) {
                        $accessComboKey = "$($event.attributes.sourceName)|$($event.attributes.interface)|$($event.technicalName)"
                        
                        if (-not $accessComboCounts.ContainsKey($accessComboKey)) {
                            $accessComboCounts[$accessComboKey] = @{ count = 0; error = $event.attributes.errors }
                        }
                        $accessComboCounts[$accessComboKey].count++
                    }
                    break
                }
                "auth" {
                    foreach ($event in $allEvents) {
                        $authComboKey = "$($event.actor.name)|$($event.attributes.info)|$($event.technicalName)"

                        if (-not $authComboCounts.ContainsKey($authComboKey)) {
                            $authComboCounts[$authComboKey] = @{ count = 0 }
                        }
                        $authComboCounts[$authComboKey].count++
                    }
                    break
                }
                "source_management" {
                    foreach ($event in $allEvents) {
                        $sourceComboKey = "$($event.attributes.sourceName)|$($event.technicalName)"

                        if (-not $aggComboCounts.ContainsKey($sourceComboKey)) {
                            $aggComboCounts[$sourceComboKey] = @{ count = 0 }
                        }
                        $aggComboCounts[$sourceComboKey].count++
                    }
                    break
                }
                "password_activity" {
                    foreach ($event in $allEvents) {

                        if ($event.technicalName -eq "password_change_failed") {
                        
                            if ($event.attributes.errors.Contains("Timeout waiting for response")) {
                                $shortError = "Timeout waiting for response"
                            } else {
                                $shortError = $event.attributes.errors
                            }
                
                            $passwordComboKey = "$($event.attributes.sourceName)|$shortError"

                            if (-not $passwordComboCounts.ContainsKey($passwordComboKey)) {
                                $passwordComboCounts[$passwordComboKey] = @{ count = 0 }
                            }
                            $passwordComboCounts[$passwordComboKey].count++
                        } else {
                            $userPasswordComboKey1 = "Repeat User|$($event.attributes.sourceName)|$($event.actor.name)"
                            $userPasswordComboKey2 = "Repeat Issue|$($event.attributes.sourceName)|$($event.attributes.info)"

                            if (-not $userPasswordComboCounts.ContainsKey($userPasswordComboKey1)) {
                                $userPasswordComboCounts[$userPasswordComboKey1] = @{ count = 0 }
                            }
                            $userPasswordComboCounts[$userPasswordComboKey1].count++

                            if (-not $userPasswordComboCounts.ContainsKey($userPasswordComboKey2)) {
                                $userPasswordComboCounts[$userPasswordComboKey2] = @{ count = 0 }
                            }
                            $userPasswordComboCounts[$userPasswordComboKey2].count++
                        }
                    }
                    break
                }
                default {
                    Write-Host "Unknown failure type: $($_)"
                }
            }

        }

        $accountProvIssues = @()
        foreach ($key in $provisioningComboCounts.Keys) {
            $parts = $key -split '\|'
            $accountProvIssues += [PSCustomObject]@{
                "Source"     = $parts[0]
                "Error Type" = $parts[2]
                "Trigger"    = $parts[1]
                "Count"      = $provisioningComboCounts[$key].count
            }
        }

        $accessProvIssues = @()
        foreach ($key in $accessComboCounts.Keys) {
            $parts = $key -split '\|'
            $accessProvIssues += [PSCustomObject]@{
                "Source"     = $parts[0]
                "Error Type" = $parts[2]
                "Trigger"    = $parts[1]
                "Count"      = $accessComboCounts[$key].count
            }
        }

        $authIssues = @()
        foreach ($key in $authComboCounts.Keys) {
            #Only going be be concerned if more than 5 attempts in the last 30 days
            if ($authComboCounts[$key].count -ge 5) {
                $parts = $key -split '\|'
                $authIssues += [PSCustomObject]@{
                    "Actor"       = $parts[0]
                    "Error Type"  = $parts[2]
                    "Information" = $parts[1]
                    "Count"       = $authComboCounts[$key].count
                }
            }
        }

        $aggIssues = @()
        foreach ($key in $aggComboCounts.Keys) {
            $parts = $key -split '\|'
            $sourceName = $parts[0]
            $technicalName = $parts[1]
            $count = $aggComboCounts[$key].count

            $technicalNamePassed = $technicalName -replace "FAILED", "PASSED"

            $jsonBody = @"
{
    "query": {
        "query": "created:[$daysInPast TO now] AND target.name:\`"$sourceName\`" AND (technicalName:$technicalName OR technicalName:$technicalNamePassed)"
    },
    "indices": ["events"]
}
"@
        
            # $totalAggResponse = Invoke-WebRequest -Uri "https://$script:tenant.api.identitynow-demo.com/v3/search?offset=0&limit=1&count=true" -Method Post -Headers $headers -Body $jsonBody
            $totalAggResponse = Invoke-WebRequest -Uri "https://$script:tenant.api.identitynow.com/v3/search?offset=0&limit=1&count=true" -Method Post -Headers $headers -Body $jsonBody
            $totalAggregations = [int]$totalAggResponse.Headers["X-Total-Count"]

            # Calculate percentage
            $percentage = ([double]$count / [double]$totalAggregations) * 100
            $formattedPercent = "{0:N1}" -f $percentage

            $aggIssues += [PSCustomObject]@{
                "Source"              = $sourceName
                "Error Type"          = $technicalName
                "30 Day Count"        = $count
                "30 Day Failure Rate" = $formattedPercent + "%"
            }
        }

        $passwordIssues = @()
        foreach ($key in $passwordComboCounts.Keys) {
            $parts = $key -split '\|'
            $passwordIssues += [PSCustomObject]@{
                "Source"     = $parts[0]
                "Error Type" = $parts[1]
                "Count"      = $passwordComboCounts[$key].count
            }
        }

        # If no results, add a default "no data" object
        if ($authIssues.Count -eq 0) {
            $authIssues += [PSCustomObject]@{
                "Actor"       = "--NO DATA--"
                "Error Type"  = "--NO DATA--"
                "Information" = "--NO DATA--"
                "Count"       = "--NO DATA--"
            }
        }

        # If no results, add a default "no data" object
        if ($aggIssues.Count -eq 0) {
            $aggIssues += [PSCustomObject]@{
                "Source"              = "--NO DATA--"
                "Error Type"          = "--NO DATA--"
                "30 Day Count"        = "--NO DATA--"
                "30 Day Failure Rate" = "--NO DATA--"
            }
        }

        # If no results, add a default "no data" object
        if ($passwordIssues.Count -eq 0) {
            $passwordIssues += [PSCustomObject]@{
                "Source"     = "--NO DATA--"
                "Error Type" = "--NO DATA--"
                "Count"      = "--NO DATA--"
            }
        }

        # If no results, add a default "no data" object
        if ($accountProvIssues.Count -eq 0) {
            $accountProvIssues += [PSCustomObject]@{
                "Source"     = "--NO DATA--"
                "Error Type" = "--NO DATA--"
                "Trigger"    = "--NO DATA--"
                "Count"      = "--NO DATA--"
            }
        }

        # If no results, add a default "no data" object
        if ($accessProvIssues.Count -eq 0) {
            $accessProvIssues += [PSCustomObject]@{
                "Source"     = "--NO DATA--"
                "Error Type" = "--NO DATA--"
                "Trigger"    = "--NO DATA--"
                "Count"      = "--NO DATA--"
            }
        }

        $authIssues = $authIssues | Sort-Object "Count" -Descending
        $aggIssues = $aggIssues | Sort-Object @{ Expression = { [double] } } -Descending
        $passwordIssues = $passwordIssues | Sort-Object "Count" -Descending
        $accountProvIssues = $accountProvIssues | Sort-Object "Count" -Descending
        $accessProvIssues = $accessProvIssues | Sort-Object "Count" -Descending

        return @{
            "authIssues"         = $authIssues
            "aggIssues"          = $aggIssues
            "passwordIssues"     = $passwordIssues
            "accountProvSummary" = $accountProvSummaryData
            "accountProvIssues"  = $accountProvIssues
            "accessProvSummary"  = $accessProvSummaryData
            "accessProvIssues"   = $accessProvIssues
        }

    } catch {
        Write-Error "Failed to retrieve operational data: $_"
        return $null
    }

}

function Get-WorkflowData() {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")

    try {
        #Get all workflows"
		# $workflowResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/workflows" -Method 'GET' -Headers $headers
		$workflowResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/workflows" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($workflow in $workflowResponse) {
            
            # getting workflow object for each workflow ID (workflowResponse returns 0 for executionCount)
            $id = $workflow.id
            # $workflowDetails = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/workflows/$id" -Method 'GET' -Headers $headers
            $workflowDetails = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/workflows/$id" -Method 'GET' -Headers $headers

            if ($workflowDetails.executionCount -gt 0) {
                $errorRate = ($workflowDetails.failureCount / $workflowDetails.executionCount) * 100
                $formattedErrorRate = "{0:N1}" -f $errorRate
            } else {
                $formattedErrorRate = "0"
            }

            $results += [PSCustomObject]@{
                "Name"            = $workflow.name
                "Owner"           = $workflow.owner.name
                "Enabled"         = $workflow.enabled
                "Execution Count" = $workflowDetails.executionCount
                "Error Rate"      = $formattedErrorRate + "%"
            }
        }

		return $results | Sort-Object @{ Expression = { [double] } } -Descending
        
    } catch {
        Write-Error "Failed to retrieve workflow data: $_"
        return $null
    }


    $headers = @("Name", "Owner", "Enabled", "Error Rate")


}



function Get-AggregationData() { 
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {
        #Filtering on type "Active Directory - Direct"
		# $adSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources" -Method 'GET' -Headers $headers
		$adSourceResponse = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources" -Method 'GET' -Headers $headers

		# Prepare output array
		$results = @()

		foreach ($source in $adSourceResponse) {
            $sourceId = $source.id
            $sourceName = $source.name

            # Account Aggregation
            # $accountUrl = "https://$tenant.api.identitynow-demo.com/v2024/task-status?limit=10&offset=0&filters=sourceId eq `"$sourceId`" and type in (`"CLOUD_ACCOUNT_AGGREGATION`")&sorters=-created"
            $accountUrl = "https://$tenant.api.identitynow.com/v2024/task-status?limit=10&offset=0&filters=sourceId eq `"$sourceId`" and type in (`"CLOUD_ACCOUNT_AGGREGATION`")&sorters=-created"
            $accountTasks = Invoke-RestMethod -Uri $accountUrl -Method 'GET' -Headers $headers
            $accountAvg = if ($accountTasks.Count -gt 0) { Get-AverageTime $accountTasks } else { "N/A" }

            # Entitlement Aggregation
            # $entitlementUrl = "https://$tenant.api.identitynow-demo.com/v2024/task-status?limit=10&offset=0&filters=sourceId eq `"$sourceId`" and type in (`"CLOUD_GROUP_AGGREGATION`")&sorters=-created"
            $entitlementUrl = "https://$tenant.api.identitynow.com/v2024/task-status?limit=10&offset=0&filters=sourceId eq `"$sourceId`" and type in (`"CLOUD_GROUP_AGGREGATION`")&sorters=-created"
            $entitlementTasks = Invoke-RestMethod -Uri $entitlementUrl -Method 'GET' -Headers $headers
            $entitlementAvg = if ($entitlementTasks.Count -gt 0) { Get-AverageTime $entitlementTasks } else { "N/A" }

            $results += [PSCustomObject]@{
                "Source Name"          = $sourceName
                "Account Schedule"     = "--MANUAL--"
                "Account Avg. Time"    = $accountAvg
                "Entitlement Schedule" = "--MANUAL--"
                "EntitlementAvg. Time" = $entitlementAvg
            }
        }

		return $results | Sort-Object "Source Name"
        
    } catch {
        Write-Error "Failed to retrieve aggregation data: $_"
        return $null
    }
}