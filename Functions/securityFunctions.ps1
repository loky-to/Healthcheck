# Country code to full name mapping
$countryMap = @{
    "AU" = "Australia"
    "NZ" = "New Zealand"
    "US" = "United States"
    "GB" = "United Kingdom"
    "CA" = "Canada"
    # Add more as needed
}

function Convert-MinutesToReadableTime {
    param (
        [string]$minutesString
    )

    $minutes = [int]$minutesString
    $hours = [math]::Floor($minutes / 60)
    $remainingMinutes = $minutes % 60

    if ($minutes -eq 60) {
        return "1 Hour"
    } elseif ($hours -ge 1 -and $remainingMinutes -gt 0) {
        $hourLabel = if ($hours -eq 1) { "Hour" } else { "Hours" }
        $minuteLabel = if ($remainingMinutes -eq 1) { "Minute" } else { "Minutes" }
        return "$hours $hourLabel $remainingMinutes $minuteLabel"
    } elseif ($hours -ge 1) {
        $hourLabel = if ($hours -eq 1) { "Hour" } else { "Hours" }
        return "$hours $hourLabel"
    } else {
        $minuteLabel = if ($minutes -eq 1) { "Minute" } else { "Minutes" }
        return "$minutes $minuteLabel"
    }
}

function Get-Auth-Org-Lockout-Config {
	
	$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	$headers.Add("Accept", "application/json")
	$headers.Add("Authorization", "Bearer $script:accessToken")
	
    try {
		# $lockoutResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow-demo.com/v3/auth-org/lockout-config" -Method 'GET' -Headers $headers
		$lockoutResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow.com/v3/auth-org/lockout-config" -Method 'GET' -Headers $headers

        return [PSCustomObject]@{
            MaximumAttempts = $lockoutResponse.maximumAttempts.ToString()
            LockoutDuration = Convert-MinutesToReadableTime -minutesString $lockoutResponse.lockoutDuration
            LockoutWindow   = Convert-MinutesToReadableTime -minutesString $lockoutResponse.lockoutWindow
        }
    }
    catch {
        Write-Error "Failed to retrieve lockout config: $_"
        return $null
    }
}

function Get-Auth-Org-Session-Config {
	
	$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	$headers.Add("Accept", "application/json")
	$headers.Add("Authorization", "Bearer $script:accessToken")
	
    try {
		# $sessionResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow-demo.com/v3/auth-org/session-config" -Method 'GET' -Headers $headers
		$sessionResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow.com/v3/auth-org/session-config" -Method 'GET' -Headers $headers

        return [PSCustomObject]@{
            maxSessionTime = Convert-MinutesToReadableTime -minutesString $sessionResponse.maxSessionTime
            maxIdleTime    = Convert-MinutesToReadableTime -minutesString $sessionResponse.maxIdleTime
            rememberMe     = (-not $sessionResponse.rememberMe).ToString()
        }
    }
    catch {
        Write-Error "Failed to retrieve session config: $_"
        return $null
    }
}

function Get-Auth-Org-SSO-Config {
	
	$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	$headers.Add("Accept", "application/json")
	$headers.Add("Authorization", "Bearer $script:accessToken")
	
    try {
		# $ssoResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow-demo.com/v3/auth-org/service-provider-config" -Method 'GET' -Headers $headers
		$ssoResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow.com/v3/auth-org/service-provider-config" -Method 'GET' -Headers $headers

        return [PSCustomObject]@{
            enabled   = ($ssoResponse.enabled).ToString()
            bypassIDP = ($ssoResponse.bypassIDP).ToString()
        }
    }
    catch {
        Write-Error "Failed to retrieve sso config: $_"
        return $null
    }
}

function Get-Auth-Org-Network-Config {
	
	$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	$headers.Add("Accept", "application/json")
	$headers.Add("Authorization", "Bearer $script:accessToken")
	
    try {
		# $networkResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow-demo.com/v3/auth-org/network-config" -Method 'GET' -Headers $headers
		$networkResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow.com/v3/auth-org/network-config" -Method 'GET' -Headers $headers
		
		# Convert geolocation codes to full names or return "None" if null
		if ($null -eq $networkResponse.geolocation) {
			$countryString = "None"
		} else {
			$fullNames = $networkResponse.geolocation | ForEach-Object { $countryMap[$_] }
			$countryString = $fullNames -join ", "
		}

        return [PSCustomObject]@{
            countryString = $countryString
        }
    }
    catch {
        Write-Error "Failed to retrieve network config: $_"
        return $null
    }
}

function Get-Admin-Access {

	$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	$headers.Add("Accept", "application/json")
	$headers.Add("Content-Type", "application/json")
	$headers.Add("Authorization", "Bearer $script:accessToken")

	$identityBody = @"
{
  "indices": [
    "identities"
  ],
  "query": {
    "query": "@access(source.name:\"IdentityNow\")"
  }
}
"@

    try {
		# $identityResponse = Invoke-RestMethod -Uri "https://$script:tenant.api.identitynow-demo.com/v3/search" -Method Post -Headers $headers -Body $identityBody
		$identityResponse = Invoke-RestMethod -Uri "https://$script:tenant.api.identitynow.com/v3/search" -Method Post -Headers $headers -Body $identityBody

		# Prepare output array
		$results = @()

		foreach ($identity in $identityResponse) {
			$name = $identity.name
			$displayName = $identity.displayName
			$identitySource = $identity.source.name

			$account = $identity.accounts | Where-Object { $_.source.type -eq "IdentityNowConnector" }
			$assignedGroups = $account.entitlementAttributes.assignedGroups -join " "

			# Search for last login event
			$eventBody = @"
{
  "indices": [
	"events"
  ],
  "query": {
	"query": "name:\"Request Authentication Passed\" AND attributes.info:LOGIN_SUCCESS* AND actor.name:\"$name\""
  },
  "sort":["-created"]
}
"@

			# $eventResponse = Invoke-RestMethod -Uri "https://$script:tenant.api.identitynow-demo.com/v3/search" -Method Post -Headers $headers -Body $eventBody
			$eventResponse = Invoke-RestMethod -Uri "https://$script:tenant.api.identitynow.com/v3/search" -Method Post -Headers $headers -Body $eventBody

			# Process last login
			if ($eventResponse.Count -gt 0) {
                $lastEvent = $eventResponse[0]
				$createdUtc = [datetime]::Parse($lastEvent.created).ToUniversalTime()
				$melbourneTime = $createdUtc.AddHours(10)  # Adjust for Melbourne time (UTC+10)
				$formattedDate = $melbourneTime.ToString("dd/MM/yyyy HH:mm")
				$loginType = "N/A"
                
                # Calculate days since last login
                $daysSinceLogin = (Get-Date) - $melbourneTime

                # setting cell to yellow if last login was more than 90 days ago
                if ($daysSinceLogin.TotalDays -gt 90) {
                    $colourCode = $script:colourCodeYellow
                } else {
                    $colourCode = $script:colourCodeGreen
                }

				switch ($lastEvent.attributes.info) {
					"LOGIN_SUCCESS_SAML" { $loginType = "SSO" }
					"LOGIN_SUCCESS"      { $loginType = "Password" }
					default              { $loginType = "N/A" }
				}
			} else {
				$formattedDate = "Never Logged In"
                $colourCode = $script:colourCodeRed
			}

			# Add to results
			$results += [pscustomobject]@{
				"Display Name"   = $displayName
				"Identity Source"= $identitySource
				"Assigned Groups"= $assignedGroups
				"Last Login"     = $formattedDate
				"Login Type"     = $loginType
                "Colour"         = $colourCode
			}
		}
        
		return $results | Sort-Object "Display Name"
		
    }
    catch {
        Write-Error "Failed to retrieve admin data: $_"
        return $null
    }
}

function Get-PAT-Data {
	
	$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
	$headers.Add("Accept", "application/json")
	$headers.Add("Authorization", "Bearer $script:accessToken")
	
    try {
		# $patResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow-demo.com/v3/personal-access-tokens" -Method 'GET' -Headers $headers
		$patResponse = Invoke-RestMethod "https://$script:tenant.api.identitynow.com/v3/personal-access-tokens" -Method 'GET' -Headers $headers
		
		$scopeAllCount = 0
        $inactiveCount = 0
        $neverCount = 0

        foreach ($token in $patResponse) {
            if ($token.scope -contains "sp:scopes:all") {
                $scopeAllCount++
            }

            if ($null -eq $token.lastUsed) {  
                $neverCount++
                continue
            }

            $trimmedDate = $($token.lastUsed).Substring(0, $($token.lastUsed).LastIndexOf('.'))

            $lastUsedDateTime = [datetime]::ParseExact($trimmedDate, "yyyy-MM-ddTHH:mm:ss", $null)
            $currentDateTime = Get-Date
            $daysDifference = ($currentDateTime - $lastUsedDateTime).Days
            if ($daysDifference -gt 90) {
                $inactiveCount++
            }
        }

        # Add to results
        $results += [pscustomobject]@{
            "PATTotal"      = $patResponse.Length
            "PATAdminTotal" = $scopeAllCount
            "PATNever"      = $neverCount
            "PAT90"         = $inactiveCount
        }

		return $results
    }
    catch {
        Write-Error "Failed to retrieve network config: $_"
        return $null
    }
}