#Load functions from the same directory
. "$PSScriptRoot\Functions\wordFunctions.ps1"
. "$PSScriptRoot\Functions\securityFunctions.ps1"
. "$PSScriptRoot\Functions\integrationFunctions.ps1"
. "$PSScriptRoot\Functions\sourceFunctions.ps1"
. "$PSScriptRoot\Functions\passwordFunctions.ps1"
. "$PSScriptRoot\Functions\operationalFunctions.ps1"
. "$PSScriptRoot\Functions\dataFunctions.ps1"


# Prompt for user input
# $consultant = Read-Host "Enter Your Name (e.g., Graham Richards)"
# $clientName = Read-Host "Enter Client Name (e.g., Calvary Health Care)"
# $script:tenant = Read-Host "Enter Tenant Name (e.g., calvarycare)"
# $clientId = Read-Host "Enter Client ID"
# $clientSecret = Read-Host "Enter Secret"

#TEMP Graham work around
$consultant = "Loky To"
$clientName = "APAM"
$script:tenant = "apac"
$clientId = Get-Content "C:\Users\LokyTo\OneDrive - Rowe Consulting Services Pty Ltd\Documents\Loky's repo\Healthcheck\APAM_client_id.txt"
$clientSecret = Get-Content "C:\Users\LokyTo\OneDrive - Rowe Consulting Services Pty Ltd\Documents\Loky's repo\Healthcheck\APAM_client_secret.txt"

# Define word doc paths
$underscoreClientName = $clientName -replace '\s+', '_'
$templatePath = "$PSScriptRoot\SailPoint_HealthCheck_TEMPLATE.docx"
$outputPath = "$PSScriptRoot\SailPoint_HealthCheck_${underscoreClientName}_$(Get-Date -Format 'yyMMdd_HHmm').docx"

# Define the colour codes for cell shading
$script:colourCodeRed = (0x9F -shl 16) -bor (0x9F -shl 8) -bor 0xFF
$script:colourCodeYellow = (204 -shl 16) -bor (242 -shl 8) -bor 255
$script:colourCodeGreen = (217 -shl 16) -bor (239 -shl 8) -bor 226
$script:colourCodeTransparent = $null

#Generate an access token that will be used for all API calls
Write-Host "Getting Token"
# $jsonToken = Invoke-RestMethod -Method Post -Uri "https://$script:tenant.api.identitynow-demo.com/oauth/token?grant_type=client_credentials&client_id=$clientId&client_secret=$clientSecret"
$jsonToken = Invoke-RestMethod -Method Post -Uri "https://$script:tenant.api.identitynow.com/oauth/token?grant_type=client_credentials&client_id=$clientId&client_secret=$clientSecret"
$script:accessToken = $jsonToken.access_token


# Copy the template to the output path
Copy-Item -Path $templatePath -Destination $outputPath -Force

# Create Word COM object
$word = New-Object -ComObject Word.Application
$word.Visible = $false
# Open the document
$doc = $word.Documents.Open($outputPath)


#### ----- SECTION 1 EXECUTIVE SUMMARY & SECTION 2 TENANT SECURITY SETTINGS ----- ####

#Get the "Tenant Information"
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Accept", "application/json")
$headers.Add("Authorization", "Bearer $accessToken")

# $identityCount = Invoke-WebRequest "https://$script:tenant.api.identitynow-demo.com/v2024/identities?count=true&limit=1" -Method 'GET' -Headers $headers
# $sourceCount = Invoke-WebRequest "https://$script:tenant.api.identitynow-demo.com/v2024/sources?count=true&limit=1" -Method 'GET' -Headers $headers
$identityCount = Invoke-WebRequest "https://$script:tenant.api.identitynow.com/v2024/identities?count=true&limit=1" -Method 'GET' -Headers $headers
$sourceCount = Invoke-WebRequest "https://$script:tenant.api.identitynow.com/v2024/sources?count=true&limit=1" -Method 'GET' -Headers $headers
$headers.Add("X-SailPoint-Experimental", "true")
# $appCount = Invoke-WebRequest "https://$script:tenant.api.identitynow-demo.com/v2024/source-apps/all?count=true&limit=1" -Method 'GET' -Headers $headers
$appCount = Invoke-WebRequest "https://$script:tenant.api.identitynow.com/v2024/source-apps/all?count=true&limit=1" -Method 'GET' -Headers $headers

#Get the "Tenant Access Controls" data
Write-Host "Getting Lockout Data"
$authOrgLockoutConfig = Get-Auth-Org-Lockout-Config
Write-Host "Getting Session Data"
$authOrgSessionConfig = Get-Auth-Org-Session-Config
Write-Host "Getting SSO Data"
$authOrgSSOConfig = Get-Auth-Org-SSO-Config
Write-Host "Getting Network Data"
$authOrgNetworkConfig = Get-Auth-Org-Network-Config

#Get the "Admin Access" data
Write-Host "Getting Admin Data"
$adminAccessData = Get-Admin-Access
#Get the "Personal Access Tokens" data
Write-Host "Getting PAT Data"
$patData = Get-PAT-Data


#### ----- SECTION 3 INTEGRATIONS ----- ####

#Get the "Virtual Appliances" data
Write-Host "Getting VA Data"
$vaData = Get-VA-Data
#Get the "Service Desk Integration" data

#Get the "Credential Providers" data


#### ----- SECTION 4 SOURCES ----- ####
Write-Host "Getting Source Data"
$adSourceData = Get-AD-Source-Data
$iqServiceData = Get-IQ-Service-Data
$entraSourceData = Get-Entra-Source-Data
$jbdcSourceData = Get-JDBC-Source-Data
$webServiceSourceData = Get-WebService-Source-Data
$delimitedSourceData = Get-Delimited-Source-Data
$miscSourceData = Get-Other-Source-Data

# splitting other source data into duplicate types and misc
$otherSourceData = $miscSourceData.MiscSources
$duplicateTypeOtherSourceData = $miscSourceData.DuplicateTypeGroups

#### ----- SECTION 5 PASSWORD POLICIES ----- ####
Write-Host "Getting Password Policy Data"
$passwordPolicyData = Get-PasswordPolicyData

#### ----- SECTION 6 OPERATIONAL BEHAVIOUR ----- ####
Write-Host "Getting Operations Data"
$operationalData = Get-OperationalData

#Authentication Issues
$opAuthData = $operationalData.authIssues
#Aggregation Failures
$opAggData = $operationalData.aggIssues
#Password Failures
$opPwdData = $operationalData.passwordIssues
#Account Provisioning Failures
$opAccountProvSummary = $operationalData.accountProvSummary
$opAccountProvData = $operationalData.accountProvIssues
#Access Provisioning Failures
$opAccessProvSummary = $operationalData.accessProvSummary
$opAccessProvData = $operationalData.accessProvIssues

#Workflow
Write-Host "Getting Workflow Data"
$workflowData = Get-WorkflowData

#Aggregations
Write-Host "Getting Aggregation Data"
$aggregationData = Get-AggregationData

#### ----- SECTION 7 DATA ANALYSIS ----- ####
Write-Host "Getting Uncorrelated Data"
$uncorrelatedData = Get-AD-Uncorrelated-Data
Write-Host "Getting Role Data"
$roleData = Get-Role-Data
Write-Host "Getting Workgroup Data"
$workgroupData = Get-Workgroup-Data


# Gather all the value data ready to replace
$singleValues = @{
    "{{Month}}" = (Get-Date -Format "MMMM")
    "{{Year}}" = (Get-Date -Format "yyyy")
    "{{ClientName}}" = $clientName
    "{{Date}}" = (Get-Date -Format "dd/MM/yyyy")
    "{{Consultant}}" = $consultant
    "{{TenantName}}" = $script:tenant
    "{{IdentityCount}}" = if ($identityCount -and $identityCount.Headers) { $identityCount.Headers["X-Total-Count"][0].ToString() } else { "!!!ERROR!!!" }
    "{{SourceCount}}" = if ($sourceCount -and $sourceCount.Headers) { $sourceCount.Headers["X-Total-Count"][0].ToString() } else { "!!!ERROR!!!" }
    "{{AppCount}}" = if ($appCount -and $appCount.Headers) { $appCount.Headers["X-Total-Count"][0].ToString() } else { "!!!ERROR!!!" }
    "{{SSOEnabled}}" = if ($authOrgSSOConfig) { $authOrgSSOConfig.enabled } else { "!!!ERROR!!!" }
    "{{BypassSSO}}" = if ($authOrgSSOConfig) { $authOrgSSOConfig.bypassIDP } else { "!!!ERROR!!!" }
    "{{TrustedCountries}}" = if ($authOrgNetworkConfig) { $authOrgNetworkConfig.countryString } else { "!!!ERROR!!!" }
    "{{SignInMaxAttempts}}" = if ($authOrgLockoutConfig) { $authOrgLockoutConfig.MaximumAttempts } else { "!!!ERROR!!!" }
    "{{CounterResetTime}}" = if ($authOrgLockoutConfig) { $authOrgLockoutConfig.LockoutWindow } else { "!!!ERROR!!!" }
    "{{SignInLockoutTime}}" = if ($authOrgLockoutConfig) { $authOrgLockoutConfig.LockoutDuration } else { "!!!ERROR!!!" }
    "{{PasswordMaxAttempts}}" = "TESTING"
    "{{PasswordLockoutTime}}" = "TESTING"
    "{{MaxSessionLength}}" = if ($authOrgSessionConfig) { $authOrgSessionConfig.maxSessionTime } else { "!!!ERROR!!!" }
    "{{EndSession}}" = if ($authOrgSessionConfig) { $authOrgSessionConfig.rememberMe } else { "!!!ERROR!!!" }
    "{{IdleSessionExpiration}}" = if ($authOrgSessionConfig) { $authOrgSessionConfig.maxIdleTime } else { "!!!ERROR!!!" }
    "{{PATTotal}}" = if ($patData) { $patData.PATTotal } else { "!!!ERROR!!!" }
    "{{PATAdminTotal}}" = if ($patData) { $patData.PATAdminTotal } else { "!!!ERROR!!!" }
    "{{PATNever}}" = if ($patData) { $patData.PATNever } else { "!!!ERROR!!!" }
    "{{PAT90}}" = if ($patData) { $patData.PAT90 } else { "!!!ERROR!!!" }
    "{{PATNoExpiry}}" = if ($patData) { $patData.PATNoExpiry } else { "!!!ERROR!!!" }
    "{{AccModFail}}" = if ($opAccountProvSummary) { $opAccountProvSummary.AccModFail } else { "!!!ERROR!!!" }
    "{{AccCrtFail}}" = if ($opAccountProvSummary) { $opAccountProvSummary.AccCrtFail } else { "!!!ERROR!!!" }
    "{{AccDisFail}}" = if ($opAccountProvSummary) { $opAccountProvSummary.AccDisFail } else { "!!!ERROR!!!" }
    "{{AccEnaFail}}" = if ($opAccountProvSummary) { $opAccountProvSummary.AccEnaFail } else { "!!!ERROR!!!" }
    "{{EntAddFail}}" = if ($opAccessProvSummary) { $opAccessProvSummary.EntAddFail } else { "!!!ERROR!!!" }
    "{{EntRevFail}}" = if ($opAccessProvSummary) { $opAccessProvSummary.EntRevFail } else { "!!!ERROR!!!" }
    "{{RoleNoAccess}}" = if ($roleData["RoleSummary"]) { $roleData["RoleSummary"].RoleNoAccess } else { "!!!ERROR!!!" }
    "{{RoleNoUsers}}" = if ($roleData["RoleSummary"]) { $roleData["RoleSummary"].RoleNoUsers } else { "!!!ERROR!!!" }
    "{{RoleNoAccUser}}" = if ($roleData["RoleSummary"]) { $roleData["RoleSummary"].RoleNoAccUser } else { "!!!ERROR!!!" }
    "{{RoleDisable}}" = if ($roleData["RoleSummary"]) { $roleData["RoleSummary"].RoleDisable } else { "!!!ERROR!!!" }
    "{{RoleGood}}" = if ($roleData["RoleSummary"]) { $roleData["RoleSummary"].RoleGood } else { "!!!ERROR!!!" }
}

# Replace single-value data points after capturing them all above
# Move selection to start of document
$word.Selection.HomeKey(6)

foreach ($key in $singleValues.Keys) {

    if ($null -eq $singleValues[$key]) {
        Write-Warning "Value for key '$key' is null"
        continue
    }
    $newValue = $singleValues[$key].ToString()
    Write-Host "Replacing '${key}' with '${newValue}'"

    # Going through doc to replace placeholder text
    Replace-Placeholder-Text -range $doc.Content -placeholder $key -newText $newValue

    # Looping through doc to replace placeholder text in shapes/text box
    foreach ($shape in $doc.Shapes) {
        try {
            if ($shape.TextFrame -and $shape.TextFrame.HasText) {
                Replace-Placeholder-Text -range $shape.TextFrame.TextRange -placeholder $key -newText $newValue
            }
        }
        catch {
            Write-Warning "No text found"
        }
    }

}

# Update Admin Access Table
Write-Host "Updating Admin Access Table"
$headers = @("Display Name", "Identity Source", "Assigned Groups", "Login Type", "Last Login")
Populate-WordTable -doc $doc -placeholder "AdminAccessTable" -inputData $adminAccessData -headers $headers

# Update VA Table
Write-Host "Updating VA Table"
$headers = @("Cluster Name", "# of VAs", "CCG Version", "Status", "# of Connections", "Recommendations")
Populate-WordTable -doc $doc -placeholder "VATable" -inputData $vaData -headers $headers

# Update AD Source Table
Write-Host "Updating AD Source Table"
$headers = @("Source Name", "Owner", "Auth", "Status")
Populate-WordTable -doc $doc -placeholder "ADSourceTable" -inputData $adSourceData -headers $headers

# Update IQService Table
Write-Host "Updating IQService Table"
$headers = @("Source Name", "IQService Host:Port", "TLS", "Version", "Status")
Populate-WordTable -doc $doc -placeholder "IQServiceTable" -inputData $iqServiceData -headers $headers

# Update Entra Source Table
Write-Host "Updating Entra Source Table"
$headers = @("Source Name", "Owner", "Auth", "Status")
Populate-WordTable -doc $doc -placeholder "EntraSourceTable" -inputData $entraSourceData -headers $headers

# Update JDBC Source Table
Write-Host "Updating JDBC Source Table"
$headers = @("Source Name", "Owner", "Connector files", "Credential Provider", "Status")
Populate-WordTable -doc $doc -placeholder "JDBCSourceTable" -inputData $jbdcSourceData -headers $headers

# Update Web Service Source Table
Write-Host "Updating Web Service Source Table"
$headers = @("Source Name", "Owner", "Auth", "Credential Provider", "Status")
Populate-WordTable -doc $doc -placeholder "WebServiceSourceTable" -inputData $webServiceSourceData -headers $headers

# Update Delimted Source Table
Write-Host "Updating Delimited Source Table"
$headers = @("Source Name", "Owner", "Status")
Populate-WordTable -doc $doc -placeholder "DelimitedSourceTable" -inputData $delimitedSourceData -headers $headers

# Update Dynamic Other Source Table
Write-Host "Updating Dynamic Other Sources Table"
$headers = @("Source Name", "Type", "Owner", "Auth", "Status")
if ($duplicateTypeOtherSourceData.Count -gt 0) {
    Populate-DuplicateSourceTables -doc $doc -duplicateGroups $duplicateTypeOtherSourceData -headers $headers
} 
# Update Other Source Table
Write-Host "Updating Other Sources Table"
$headers = @("Source Name", "Type", "Owner", "Auth", "Status")
Populate-WordTable -doc $doc -placeholder "OtherSourceTable" -inputData $otherSourceData -headers $headers

# Update Operations Authentications Table
Write-Host "Updating Ops Auth Table"
$headers = @("Actor", "Error Type", "Information", "Count")
Populate-WordTable -doc $doc -placeholder "OpAuthTable" -inputData $opAuthData -headers $headers

# Update Operations Aggregations Table
Write-Host "Updating Ops Agg Table"
$headers = @("Source", "Error Type", "30 Day Count", "30 Day Failure Rate")
Populate-WordTable -doc $doc -placeholder "OpAggTable" -inputData $opAggData -headers $headers

# Add in the password policy tables
Write-Host "Creating Password Policy Tables"
Create-Password-Policy-WordTables -doc $doc -policyData $passwordPolicyData

# Update Operations Password Table
Write-Host "Updating Ops Password Table"
$headers = @("Source", "Error Type", "Count")
Populate-WordTable -doc $doc -placeholder "OpPwdTable" -inputData $opPwdData -headers $headers

# Update Operations Account Provisioning Table
Write-Host "Updating Ops Account Prov Table"
$headers = @("Source", "Error Type", "Trigger", "Count")
Populate-WordTable -doc $doc -placeholder "OpAccountProvTable" -inputData $opAccountProvData -headers $headers

# Update Operations Access Provisioning Table
Write-Host "Updating Ops Access Prov Table"
$headers = @("Source", "Error Type", "Trigger", "Count")
Populate-WordTable -doc $doc -placeholder "OpAccessProvTable" -inputData $opAccessProvData -headers $headers

# Update Workflows Table
Write-Host "Updating Workflow Table"
$headers = @("Name", "Trigger", "Enabled", "Execution Count", "Failed Error Rate", "Cancellation Error Rate")
Populate-WordTable -doc $doc -placeholder "WorkflowTable" -inputData $workflowData -headers $headers

# Update Operations Aggregation Table
Write-Host "Updating Ops Agg Table"
$headers = @("Source Name", "Account Schedule", "Account Avg. Time", "Entitlement Schedule", "EntitlementAvg. Time")
Populate-WordTable -doc $doc -placeholder "AggregationTable" -inputData $aggregationData -headers $headers

# Update Uncorrelated Table
Write-Host "Updating Uncorrelated Table"
$headers = @("Source Name", "Disabled", "Enabled")
Populate-WordTable -doc $doc -placeholder "UncorrelatedTable" -inputData $uncorrelatedData -headers $headers

# Update Role Table
Write-Host "Updating Role Table"
$headers = @("Role Name", "Status", "Request", "Auto", "Ents", "APs", "IDs", "Summary")
Populate-WordTable -doc $doc -placeholder "RoleTable" -inputData $roleData["AllRoles"] -headers $headers

# Update Groups with No Members Table
Write-Host "Updating Workgroups Members Table"
$headers = @("Group Name", "Owner", "Member Count", "Association Count")
Populate-WordTable -doc $doc -placeholder "GroupNoMemberTable" -inputData $workgroupData["NoMembers"] -headers $headers

# Update Groups with No Associations Table
Write-Host "Updating Workgroups Associations Table"
$headers = @("Group Name", "Owner", "Member Count", "Association Count")
Populate-WordTable -doc $doc -placeholder "GroupNoAssoTable" -inputData $workgroupData["NoAssociations"] -headers $headers

# Save and close
$doc.Save()
$doc.Close()
$word.Quit()

Write-Host "Finished"
