function Get-PasswordPolicyData {
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Accept", "application/json")
    $headers.Add("Authorization", "Bearer $script:accessToken")
    $headers.Add("X-SailPoint-Experimental", "true")
    
    try {
        # $listofpolicies = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/password-policies" -Method 'GET' -Headers $headers
        $listofpolicies = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/password-policies" -Method 'GET' -Headers $headers

		# Prepare output array
		$policyData = @()
        foreach ($policy in $listofpolicies) {

            $joinedSources = ""
            foreach ($source in $policy.sourceIds) {

                # $sourceName = Invoke-RestMethod "https://$tenant.api.identitynow-demo.com/v2024/sources/$source" -Method 'GET' -Headers $headers
                $sourceName = Invoke-RestMethod "https://$tenant.api.identitynow.com/v2024/sources/$source" -Method 'GET' -Headers $headers
                $joinedSources = $joinedSources + $sourceName.name + "`r`n"

            }

            $policyData += [PSCustomObject]@{
                "Requirement" = "Setting"
                "Name" = $policy.name
                "Maximum Length" = $policy.maxLength
                "Minimum Length" = $policy.minLength
                "Minimum Letters" = $policy.minAlpha
                "Minimum Uppercase" = $policy.minUpper
                "Minimum Lowercase" = $policy.minLower
                "Minimum Digits" = $policy.minNumeric
                "Minimum Special Characters" = $policy.minSpecial
                "Minimum Character Types" = $policy.minCharacterTypes
                "Maximum Consecutive Characters" = $policy.maxRepeatedChars
                "Prevent use of account attributes" = $policy.useAccountAttributes
                "Prevent use of identity attributes" = $policy.useIdentityAttributes
                "Disallow display name fragments" = $policy.validateAgainstAccountName
                "Disallow account ID fragments" = $policy.validateAgainstAccountId
                "Prevent use of words in this site's password dictionary" = $policy.useDictionary
                "Associated Sources" = $joinedSources
            }
        }

		return $policyData

    } catch {
        Write-Error "Failed to retrieve Password Policies: $_"
        return $null
    }
}