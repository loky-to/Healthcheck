function Get-AD-Uncorrelated-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    try {

        $limit = 50
        $offset = 0
        $allaccounts = @()

        do {
            # $uri = "https://$script:tenant.api.identitynow-demo.com/beta/accounts?limit=$limit&offset=$offset&count=true&filters=uncorrelated eq true"
            $uri = "https://$script:tenant.api.identitynow.com/beta/accounts?limit=$limit&offset=$offset&count=true&filters=uncorrelated eq true"
            $response = Invoke-WebRequest -Uri $uri -Method 'GET' -Headers $headers

            #$accounts = $response.Content | ConvertFrom-Json
            Add-Type -AssemblyName System.Web.Extensions
            $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
            $accounts = $serializer.DeserializeObject($response.Content)

            $totalCount = [int]$response.Headers["X-Total-Count"]

            $allAccounts += $accounts
            $offset += $limit
        } while ($offset -lt $totalCount)

        # Group and count by source name
        $groupedResults = @{}

        foreach ($account in $allAccounts) {
            $source = $account.sourceName
            $disabled = $account.disabled

            if (-not $groupedResults.ContainsKey($source)) {
                $groupedResults[$source] = @{
                    "uncorrelated disabled" = 0
                    "uncorrelated disabled with access" = 0
                    "uncorrelated enabled" = 0
                    "uncorrelated enabled with access" = 0
                    "uncorrelated enabled total" = 0
                    "uncorrelated disabled total" = 0
                }
            }

            if ($disabled) {
                $groupedResults[$source]["uncorrelated disabled total"]++
            } else {
                $groupedResults[$source]["uncorrelated enabled total"]++
            }
        }

        $finalData = @()
        foreach ($source in $groupedResults.Keys) {
            $finalData += [PSCustomObject]@{
                "Source Name"  = $source
                "Disabled"     = $groupedResults[$source]["uncorrelated disabled total"]
                "Enabled"      = $groupedResults[$source]["uncorrelated enabled total"]
            }
        }

        
        # If no results, add a default "no data" object
        if ($finalData.Count -eq 0) {
            $finalData += [PSCustomObject]@{
                "Source Name"  = "--NO DATA--"
                "Disabled"     = "--NO DATA--"
                "Enabled"      = "--NO DATA--"
            }
        }

        return $finalData | Sort-Object "Source Name"


    }
    catch {
        Write-Error "Failed to retrieve uncorrelated data: $_"
        return $null
    }
}


function Get-Role-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    $maxRetries = 5
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        try {

            $limit = 50
            $offset = 0
            $allRoles = @()
            $roleNoAccess = 0
            $roleNoUsers = 0
            $roleNoAccUser = 0
            $roleDisable = 0
            $roleGood = 0

            do {
                # $uri = "https://$script:tenant.api.identitynow-demo.com/beta/roles?limit=$limit&offset=$offset&count=true"
                $uri = "https://$script:tenant.api.identitynow.com/beta/roles?limit=$limit&offset=$offset&count=true"
                $response = Invoke-WebRequest -Uri $uri -Method 'GET' -Headers $headers

                $roles = $response.Content | ConvertFrom-Json
                $totalCount = [int]$response.Headers["X-Total-Count"]

                foreach ($role in $roles) {
                    $roleId = $role.id
                    # $identityUri = "https://$script:tenant.api.identitynow-demo.com/beta/roles/$roleId/assigned-identities?count=true"
                    $identityUri = "https://$script:tenant.api.identitynow.com/beta/roles/$roleId/assigned-identities?count=true"
                    $identityResponse = Invoke-WebRequest -Uri $identityUri -Method 'GET' -Headers $headers
                    $identityCount = [int]$identityResponse.Headers["X-Total-Count"]

                    if (!$role.enabled) {
                        $summary = "Disabled"
                        $colourCode = $script:colourCodeYellow
                        $roleDisable++
                    } else {
                        if($identityCount -eq 0 -and $role.entitlements.Count -eq 0 -and $role.accessProfiles.Count -eq 0) {
                            $summary = "No Users or Access"
                            $colourCode = $script:colourCodeRed
                            $roleNoAccUser++
                        } elseif ($identityCount -eq 0) {
                            $summary = "No Users"
                            $colourCode = $script:colourCodeRed
                            $roleNoUsers++
                        } elseif ($role.entitlements.Count -eq 0 -and $role.accessProfiles.Count -eq 0) {
                            $summary = "No Access"
                            $colourCode = $script:colourCodeRed
                            $roleNoAccess++
                        } else {
                            $summary = "Good"
                            $colourCode = $script:colourCodeGreen
                            $roleGood++
                        }
                    }

                    $allRoles += [PSCustomObject]@{
                        "Role Name" = $role.name
                        "Status"    = if ($role.enabled) { "Enabled" } else { "Disabled" }
                        "Request"   = $role.requestable
                        "Auto"      = if ($role.membership) { $true } else { $false }
                        "Ents"      = $role.entitlements.Count
                        "APs"       = $role.accessProfiles.Count
                        "IDs"       = $identityCount
                        "Summary"   = $summary
                        "Colour"    = $colourCode
                    }
                }
                $offset += $limit
            } while ($offset -lt $totalCount)

            # If no results, add a default "no data" object
            if ($allRoles.Count -eq 0) {
                $allRoles += [PSCustomObject]@{
                    "Role Name" = "--NO DATA--"
                    "Status"    = "--NO DATA--"
                    "Request"   = "--NO DATA--"
                    "Auto"      = "--NO DATA--"
                    "Ents"      = "--NO DATA--"
                    "APs"       = "--NO DATA--"
                    "IDs"       = "--NO DATA--"
                    "Summary"   = "--NO DATA--"
                }
            }

            $allRoles = $allRoles | Sort-Object "Role Name"

            $roleSummary += [PSCustomObject]@{
                "RoleNoAccess"  = $roleNoAccess
                "RoleNoUsers"   = $roleNoUsers
                "RoleNoAccUser" = $roleNoAccUser
                "RoleDisable"   = $roleDisable
                "RoleGood"      = $roleGood
            }

            return @{
                "RoleSummary" = $roleSummary
                "AllRoles"    = $allRoles
            }

        }
        catch {
           if ($_ -match '"Rate Limit Exceeded"') {
                Write-Host "Rate limit exceeded. Retrying in 5 seconds... ($($retryCount + 1)/$maxRetries)"
                Start-Sleep -Seconds 5
                $retryCount++
            } else {
                Write-Error "Failed to retrieve role data: $_"
             return $null
            }
        }
    }
}


function Get-Workgroup-Data {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    
    $maxRetries = 5
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        try {

            $limit = 50
            $offset = 0
            $noMembers = @()
            $noAssociations = @()

            do {
                # $uri = "https://$script:tenant.api.identitynow-demo.com/beta/workgroups?offset=$offset&limit=$limit&count=true"
                $uri = "https://$script:tenant.api.identitynow.com/beta/workgroups?offset=$offset&limit=$limit&count=true"
                $response = Invoke-WebRequest -Uri $uri -Method 'GET' -Headers $headers

                $groups = $response.Content | ConvertFrom-Json
                $totalCount = [int]$response.Headers["X-Total-Count"]

                foreach ($item in $groups) {
                    if ($item.memberCount -eq 0) {
                        $noMembers += [PSCustomObject]@{
                            "Group Name"        = $item.name
                            "Owner"             = $item.owner.displayName
                            "Member Count"      = $item.memberCount
                            "Association Count" = $item.connectionCount
                        }
                    }
                    if ($item.connectionCount -eq 0) {
                        $noAssociations += [PSCustomObject]@{
                            "Group Name"        = $item.name
                            "Owner"             = $item.owner.displayName
                            "Member Count"      = $item.memberCount
                            "Association Count" = $item.connectionCount
                        }
                    }
                }
                $offset += $limit
            } while ($offset -lt $totalCount)

            # If no results, add a default "no data" object
            if ($noMembers.Count -eq 0) {
                $noMembers += [PSCustomObject]@{
                    "Group Name"        = "--NO DATA--"
                    "Owner"             = "--NO DATA--"
                    "Member Count"      = "--NO DATA--"
                    "Association Count" = "--NO DATA--"
                }
            }

            # If no results, add a default "no data" object
            if ($noAssociations.Count -eq 0) {
                $noAssociations += [PSCustomObject]@{
                    "Group Name"        = "--NO DATA--"
                    "Owner"             = "--NO DATA--"
                    "Member Count"      = "--NO DATA--"
                    "Association Count" = "--NO DATA--"
                }
            }

            $noMembers = $noMembers | Sort-Object "Group Name"
            $noAssociations = $noAssociations | Sort-Object "Group Name"

            return @{
                "NoMembers"      = $noMembers
                "NoAssociations" = $noAssociations
            }

        }
        catch {
           if ($_ -match '"Rate Limit Exceeded"') {
                Write-Host "Rate limit exceeded. Retrying in 5 seconds... ($($retryCount + 1)/$maxRetries)"
                Start-Sleep -Seconds 5
                $retryCount++
            } else {
                Write-Error "Failed to retrieve workgroup data: $_"
             return $null
            }
        }
    }
}