#Load functions from the same directory
. "$PSScriptRoot\Functions\wordFunctions.ps1"
. "$PSScriptRoot\Functions\securityFunctions.ps1"
. "$PSScriptRoot\Functions\integrationFunctions.ps1"
. "$PSScriptRoot\Functions\sourceFunctions.ps1"
. "$PSScriptRoot\Functions\passwordFunctions.ps1"
. "$PSScriptRoot\Functions\operationalFunctions.ps1"
. "$PSScriptRoot\Functions\dataFunctions.ps1"

# Prompt for user input
#$script:tenant = Read-Host "Enter Tenant Name (e.g., calvarycare)"
#$clientId = Read-Host "Enter Client ID"
#$clientSecret = Read-Host "Enter Secret"

#TEMP Graham work around
$script:tenant = "apac"
$clientId = Get-Content "C:\Users\LokyTo\OneDrive - Rowe Consulting Services Pty Ltd\Documents\Loky's documents\Healthcheck\APAM_client_id.txt"
$clientSecret = Get-Content "C:\Users\LokyTo\OneDrive - Rowe Consulting Services Pty Ltd\Documents\Loky's documents\Healthcheck\APAM_client_secret.txt"

#Generate an access token that will be used for all API calls
# $jsonToken = Invoke-RestMethod -Method Post -Uri "https://$script:tenant.api.identitynow-demo.com/oauth/token?grant_type=client_credentials&client_id=$clientId&client_secret=$clientSecret"
$jsonToken = Invoke-RestMethod -Method Post -Uri "https://$script:tenant.api.identitynow.com/oauth/token?grant_type=client_credentials&client_id=$clientId&client_secret=$clientSecret"
$script:accessToken = $jsonToken.access_token

$output = Get-VA-Data

Write-Host "----"
Write-Host $output